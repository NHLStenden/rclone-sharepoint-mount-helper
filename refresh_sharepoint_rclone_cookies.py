from __future__ import annotations

import argparse
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlsplit

import requests
from requests import Session
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options

COOKIE_HEADER_LINE_RE = re.compile(r'^\s*headers\s*=\s*(.+?)\s*$', re.IGNORECASE)
REMOTE_HEADER_RE = re.compile(r"^\[(.+?)\]\s*$")
QUOTED_TOKEN_RE = re.compile(r'"((?:[^"\\]|\\.)*)"')
TEST_ENDPOINT = "/_api/web"
EXIT_GENERAL_ERROR = 1
EXIT_REMOTE_NOT_FOUND = 2
EXIT_REFRESH_VALIDATION_FAILED = 3
EXIT_BROWSER_START_FAILED = 4
EXIT_REMOTE_DETECTION_FAILED = 5
EXIT_CONFIG_ERROR = 6
EXIT_LOCK_FAILED = 7


@dataclass(frozen=True)
class RemoteInfo:
    name: str
    lines: list[str]
    options: dict[str, str]


@dataclass(frozen=True)
class CookieValidationResult:
    is_valid: bool
    status_code: Optional[int]
    reason: str
    location: Optional[str] = None
    error: Optional[str] = None


class ScriptError(RuntimeError):
    exit_code = EXIT_GENERAL_ERROR


class RemoteNotFoundError(ScriptError):
    exit_code = EXIT_REMOTE_NOT_FOUND


class RemoteDetectionError(ScriptError):
    exit_code = EXIT_REMOTE_DETECTION_FAILED


class ConfigError(ScriptError):
    exit_code = EXIT_CONFIG_ERROR


class BrowserStartError(ScriptError):
    exit_code = EXIT_BROWSER_START_FAILED


class LockError(ScriptError):
    exit_code = EXIT_LOCK_FAILED


class Logger:
    def __init__(self, log_file: Optional[Path] = None) -> None:
        self.log_file = log_file
        if self.log_file is not None:
            self.log_file.parent.mkdir(parents=True, exist_ok=True)

    def log(self, msg: str) -> None:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        print(line)
        if self.log_file is not None:
            with self.log_file.open("a", encoding="utf-8") as handle:
                handle.write(line + "\n")


class FileLock:
    def __init__(self, lock_file: Path) -> None:
        self.lock_file = lock_file
        self.handle = None

    def __enter__(self) -> FileLock:
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.release()

    def acquire(self) -> None:
        self.lock_file.parent.mkdir(parents=True, exist_ok=True)
        try:
            handle = open(self.lock_file, "a+", encoding="utf-8")
            if os.name == "nt":
                import msvcrt

                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            handle.seek(0)
            handle.truncate()
            handle.write(
                f"pid={os.getpid()}\n"
                f"host={socket.gethostname()}\n"
                f"script={Path(__file__).resolve()}\n"
                f"time={datetime.now().isoformat()}\n"
            )
            handle.flush()
            self.handle = handle
        except OSError as exc:
            raise LockError(f"Another cookie refresh process is already running. Lock file: {self.lock_file}") from exc

    def release(self) -> None:
        if self.handle is None:
            return
        try:
            if os.name == "nt":
                import msvcrt

                self.handle.seek(0)
                msvcrt.locking(self.handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(self.handle.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        self.handle.close()
        self.handle = None


LOGGER = Logger()


def log(msg: str) -> None:
    LOGGER.log(msg)


def redact_url(value: str) -> str:
    try:
        parts = urlsplit(value)
        if not parts.scheme or not parts.netloc:
            return value
        path = parts.path.rstrip("/")
        if not path:
            return f"{parts.scheme}://{parts.netloc}/"
        segments = [segment for segment in path.split("/") if segment]
        preview = "/".join(segments[:2])
        suffix = "/..." if len(segments) > 2 else ""
        return f"{parts.scheme}://{parts.netloc}/{preview}{suffix}"
    except Exception:
        return value


def scoop_root() -> Optional[Path]:
    env = os.environ.get("SCOOP")
    if env:
        p = Path(env)
        if p.exists():
            return p
    p = Path.home() / "scoop"
    return p if p.exists() else None


def which_path(name: str) -> Optional[Path]:
    found = shutil.which(name)
    return Path(found) if found else None


def detect_rclone_exe() -> Path:
    p = which_path("rclone")
    if p:
        return p
    root = scoop_root()
    if root:
        p = root / "apps" / "rclone" / "current" / "rclone.exe"
        if p.exists():
            return p
    raise FileNotFoundError("Could not find rclone.exe")


def detect_rclone_conf(rclone_exe: Path) -> Path:
    result = subprocess.run(
        [str(rclone_exe), "config", "file"],
        capture_output=True,
        text=True,
        check=True,
    )
    stdout = result.stdout.strip()
    if not stdout:
        raise ConfigError("Could not determine rclone.conf path from 'rclone config file'.")

    candidates: list[str] = []
    for raw_line in stdout.splitlines():
        line = raw_line.strip().strip('"')
        if not line:
            continue
        candidates.append(line)
        match = re.search(r"stored at:?\s*(.+)$", line, re.IGNORECASE)
        if match:
            candidates.append(match.group(1).strip().strip('"'))

    for candidate in reversed(candidates):
        path = Path(candidate)
        if path.exists():
            return path

    raise FileNotFoundError(f"Detected rclone config does not exist. Output was: {stdout}")


def detect_chromium_binary() -> Path:
    for name in ["chromium", "chrome", "msedge"]:
        p = which_path(name)
        if p:
            return p
    root = scoop_root()
    if root:
        for rel in [
            Path("apps/chromium/current/chrome.exe"),
            Path("apps/googlechrome/current/chrome.exe"),
            Path("apps/microsoft-edge/current/msedge.exe"),
        ]:
            p = root / rel
            if p.exists():
                return p
    raise FileNotFoundError("Could not find a Chromium based browser executable")


def detect_user_data_dir(browser_binary: Path) -> Path:
    root = scoop_root()
    if root:
        name = browser_binary.parent.parent.name.lower()
        candidates: list[Path] = []
        if "edge" in name:
            candidates.append(root / "persist" / "microsoft-edge" / "User Data")
        elif "chrome" in browser_binary.name.lower() and "google" in str(browser_binary).lower():
            candidates.append(root / "persist" / "googlechrome" / "User Data")
        else:
            candidates.append(root / "persist" / "chromium" / "User Data")
        for c in candidates:
            if c.exists():
                return c

    local = Path(os.environ.get("LOCALAPPDATA", ""))
    fallbacks = [
        local / "Chromium" / "User Data",
        local / "Google" / "Chrome" / "User Data",
        local / "Microsoft" / "Edge" / "User Data",
    ]
    for p in fallbacks:
        if p.exists():
            return p
    raise FileNotFoundError("Could not find Chromium user data directory")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def parse_remote_infos(conf_text: str) -> list[RemoteInfo]:
    remotes: list[RemoteInfo] = []
    current_name: Optional[str] = None
    current_lines: list[str] = []

    def flush_current() -> None:
        nonlocal current_name, current_lines
        if current_name is None:
            return
        options: dict[str, str] = {}
        for line in current_lines[1:]:
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or stripped.startswith(";"):
                continue
            if "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            normalized_key = key.strip().lower()
            if normalized_key not in options:
                options[normalized_key] = value.strip()
        remotes.append(RemoteInfo(name=current_name, lines=current_lines[:], options=options))

    for line in conf_text.splitlines(keepends=True):
        m = REMOTE_HEADER_RE.match(line.strip())
        if m:
            flush_current()
            current_name = m.group(1)
            current_lines = [line]
        elif current_name is not None:
            current_lines.append(line)
    flush_current()
    return remotes


def get_remote_info(conf_text: str, remote_name: str) -> RemoteInfo:
    for remote in parse_remote_infos(conf_text):
        if remote.name == remote_name:
            return remote
    raise RemoteNotFoundError(f"Remote [{remote_name}] not found in rclone config")


def is_sharepoint_webdav(remote: RemoteInfo) -> bool:
    remote_type = remote.options.get("type", "").lower()
    url = remote.options.get("url", "")
    return remote_type == "webdav" and "sharepoint.com" in url.lower()


def detect_remote_name(conf_text: str, requested_remote: str) -> str:
    remotes = parse_remote_infos(conf_text)
    if not remotes:
        raise ConfigError("No remotes found in rclone config")

    if requested_remote != "auto":
        remote = get_remote_info(conf_text, requested_remote)
        if remote.options.get("type", "").lower() != "webdav":
            raise ConfigError(f"Remote [{requested_remote}] is not a WebDAV remote")
        return requested_remote

    sharepoint_webdav_remotes = [remote.name for remote in remotes if is_sharepoint_webdav(remote)]
    if len(sharepoint_webdav_remotes) == 1:
        return sharepoint_webdav_remotes[0]
    if len(sharepoint_webdav_remotes) > 1:
        names = ", ".join(sharepoint_webdav_remotes)
        raise RemoteDetectionError(
            "Multiple SharePoint WebDAV remotes found. Please specify --remote explicitly. "
            f"Candidates: {names}"
        )

    webdav_remotes = [remote.name for remote in remotes if remote.options.get("type", "").lower() == "webdav"]
    if len(webdav_remotes) == 1:
        return webdav_remotes[0]
    if len(webdav_remotes) > 1:
        names = ", ".join(webdav_remotes)
        raise RemoteDetectionError(
            "Multiple WebDAV remotes found. Please specify --remote explicitly. "
            f"Candidates: {names}"
        )

    raise RemoteDetectionError("Could not auto detect a suitable WebDAV remote in rclone config")


def get_remote_host(remote: RemoteInfo) -> str:
    remote_type = remote.options.get("type", "").lower()
    if remote_type != "webdav":
        raise ConfigError(f"Remote [{remote.name}] is not a WebDAV remote")
    url = remote.options.get("url")
    if not url:
        raise ConfigError(f"Remote [{remote.name}] does not define a url")
    parts = urlsplit(url)
    if not parts.scheme or not parts.netloc:
        raise ConfigError(f"Remote [{remote.name}] contains an invalid url: {url}")
    if "sharepoint.com" not in parts.netloc.lower():
        raise ConfigError(f"Remote [{remote.name}] is not a SharePoint URL: {url}")
    return f"{parts.scheme}://{parts.netloc}"


def parse_cookie_blob(cookie_blob: str) -> tuple[Optional[str], Optional[str]]:
    fedauth = None
    rtfa = None
    for part in cookie_blob.split(";"):
        part = part.strip()
        if not part:
            continue
        key, sep, value = part.partition("=")
        if not sep:
            continue
        if key == "FedAuth":
            fedauth = value
        elif key in {"rtFa", "rtFA"}:
            rtfa = value
    return fedauth, rtfa


def parse_header_tokens_from_line(line: str) -> Optional[list[str]]:
    match = COOKIE_HEADER_LINE_RE.match(line.rstrip("\r\n"))
    if not match:
        return None
    raw_value = match.group(1)
    tokens = [bytes(token, "utf-8").decode("unicode_escape") for token in QUOTED_TOKEN_RE.findall(raw_value)]
    if not tokens:
        return []
    if len(tokens) * 2 != len(re.findall(r'"', raw_value)):
        return None
    return tokens


def build_header_line(tokens: list[str]) -> str:
    escaped_tokens = []
    for token in tokens:
        escaped = token.replace("\\", "\\\\").replace('"', '\\"')
        escaped_tokens.append(f'"{escaped}"')
    return f"headers = {','.join(escaped_tokens)}\n"


def upsert_cookie_header_tokens(tokens: list[str], fedauth: str, rtfa: str) -> list[str]:
    if len(tokens) % 2 != 0:
        raise ConfigError("Malformed headers line: expected header name/value pairs")

    updated = tokens[:]
    cookie_value = f"FedAuth={fedauth};rtFa={rtfa};"
    cookie_indices: list[int] = []

    index = 0
    while index + 1 < len(updated):
        if updated[index].lower() == "cookie":
            cookie_indices.append(index)
        index += 2

    if len(cookie_indices) > 1:
        raise ConfigError("Malformed headers line: multiple Cookie headers found")

    if cookie_indices:
        updated[cookie_indices[0] + 1] = cookie_value
    else:
        updated.extend(["Cookie", cookie_value])
    return updated


def parse_cookie_values_from_lines(lines: list[str]) -> tuple[Optional[str], Optional[str]]:
    for line in lines:
        tokens = parse_header_tokens_from_line(line)
        if tokens is None or len(tokens) % 2 != 0:
            continue
        index = 0
        while index + 1 < len(tokens):
            if tokens[index].lower() == "cookie":
                return parse_cookie_blob(tokens[index + 1])
            index += 2
    return None, None


def get_cookie_value(driver: webdriver.Chrome, name: str) -> Optional[str]:
    cookie = driver.get_cookie(name)
    return cookie["value"] if cookie else None


def build_driver(binary: Path, user_data_dir: Path, profile_directory: str, headless: bool) -> webdriver.Chrome:
    options = Options()
    options.binary_location = str(binary)
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument(f"--profile-directory={profile_directory}")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-first-run-ui")
    if headless:
        options.add_argument("--headless=new")
    try:
        driver = webdriver.Chrome(options=options)
    except WebDriverException as exc:
        raise BrowserStartError(
            "Failed to start browser automation. Close existing browser instances using the same profile or pass a different profile directory. "
            f"Original error: {exc}"
        ) from exc
    driver.set_page_load_timeout(60)
    return driver


def create_session() -> Session:
    session = requests.Session()
    session.headers.update({"Accept": "application/json;odata=verbose"})
    return session


def test_sharepoint_cookie(host: str, fedauth: str, rtfa: str, timeout: int = 15) -> CookieValidationResult:
    url = host.rstrip("/") + TEST_ENDPOINT
    headers = {"Cookie": f"FedAuth={fedauth};rtFa={rtfa};"}
    try:
        with create_session() as session:
            response = session.get(url, headers=headers, timeout=timeout, allow_redirects=False)
            location = response.headers.get("Location")
            if response.status_code == 200:
                return CookieValidationResult(True, response.status_code, "valid", location=location)
            if response.status_code in {301, 302, 303, 307, 308}:
                return CookieValidationResult(False, response.status_code, "redirected", location=location)
            if response.status_code in {401, 403}:
                return CookieValidationResult(False, response.status_code, "unauthorized", location=location)
            return CookieValidationResult(False, response.status_code, "unexpected_status", location=location)
    except requests.Timeout as exc:
        return CookieValidationResult(False, None, "timeout", error=str(exc))
    except requests.ConnectionError as exc:
        return CookieValidationResult(False, None, "connection_error", error=str(exc))
    except requests.RequestException as exc:
        return CookieValidationResult(False, None, "request_error", error=str(exc))


def format_validation_result(result: CookieValidationResult) -> str:
    details = [f"reason: {result.reason}"]
    if result.status_code is not None:
        details.append(f"HTTP status: {result.status_code}")
    if result.location:
        details.append(f"location: {result.location}")
    if result.error:
        details.append(f"error: {result.error}")
    return ", ".join(details)


def fetch_sharepoint_cookies(
    driver: webdriver.Chrome,
    host: str,
    timeout_seconds: int = 60,
    non_interactive: bool = False,
) -> tuple[str, str]:
    driver.get(host)
    log("Waiting for SharePoint cookies...")
    fedauth = None
    rtfa = None
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        fedauth = get_cookie_value(driver, "FedAuth")
        rtfa = get_cookie_value(driver, "rtFA") or get_cookie_value(driver, "rtFa")
        if fedauth and rtfa:
            break
        time.sleep(1)

    if not (fedauth and rtfa):
        if non_interactive:
            raise ScriptError(
                "Cookies were not found before timeout in non interactive mode. Complete sign in manually or rerun without --non-interactive."
            )
        input("Cookies not found yet. Finish signing in, then press Enter to try once more... ")
        fedauth = get_cookie_value(driver, "FedAuth")
        rtfa = get_cookie_value(driver, "rtFA") or get_cookie_value(driver, "rtFa")

    if not fedauth or not rtfa:
        raise ScriptError("Could not retrieve FedAuth and/or rtFa cookie")
    return fedauth, rtfa


def backup_file(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = path.with_suffix(path.suffix + f".{stamp}.bak")
    shutil.copy2(path, backup)
    return backup


def atomic_write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", delete=False, dir=path.parent, encoding="utf-8", newline="") as handle:
        handle.write(content)
        handle.flush()
        os.fsync(handle.fileno())
        temp_path = Path(handle.name)
    os.replace(temp_path, path)


def update_remote_cookie_header(conf_path: Path, remote_name: str, fedauth: str, rtfa: str) -> bool:
    conf_text = read_text(conf_path)
    remotes = parse_remote_infos(conf_text)
    changed = False
    output_lines: list[str] = []
    found_remote = False

    for remote in remotes:
        lines = remote.lines[:]
        if remote.name != remote_name:
            output_lines.extend(lines)
            continue

        found_remote = True
        header_found = False
        for index, line in enumerate(lines):
            tokens = parse_header_tokens_from_line(line)
            if tokens is None:
                continue
            if len(tokens) % 2 != 0:
                raise ConfigError(f"Malformed headers line in remote [{remote_name}]")
            header_found = True
            new_tokens = upsert_cookie_header_tokens(tokens, fedauth, rtfa)
            updated_header_line = build_header_line(new_tokens)
            if line != updated_header_line:
                lines[index] = updated_header_line
                changed = True
            break

        if not header_found:
            if lines and not lines[-1].endswith("\n"):
                lines[-1] += "\n"
            lines.append(build_header_line(["Cookie", f"FedAuth={fedauth};rtFa={rtfa};"]))
            changed = True

        output_lines.extend(lines)

    if not found_remote:
        raise RemoteNotFoundError(f"Remote [{remote_name}] not found in rclone config")

    if changed:
        backup = backup_file(conf_path)
        log(f"Backup created: {backup}")
        atomic_write_text(conf_path, "".join(output_lines))
        log(f"Updated cookie header in [{remote_name}]")
    else:
        log(f"Cookie header in [{remote_name}] is already up to date")
    return changed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Refresh SharePoint cookies in rclone.conf only when needed.")
    parser.add_argument("--remote", default="auto", help="Remote name to update, or 'auto' to detect it")
    parser.add_argument("--profile-directory", default="Default")
    parser.add_argument("--force-refresh", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--non-interactive", action="store_true")
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--login-timeout", type=int, default=60)
    parser.add_argument("--rclone-exe")
    parser.add_argument("--rclone-conf")
    parser.add_argument("--chromium-binary")
    parser.add_argument("--user-data-dir")
    parser.add_argument("--log-file")
    parser.add_argument("--lock-file")
    return parser.parse_args()


def main() -> int:
    global LOGGER
    args = parse_args()
    LOGGER = Logger(Path(args.log_file) if args.log_file else None)

    try:
        with FileLock(Path(args.lock_file)) if args.lock_file else nullcontext():
            rclone_exe = Path(args.rclone_exe) if args.rclone_exe else detect_rclone_exe()
            rclone_conf = Path(args.rclone_conf) if args.rclone_conf else detect_rclone_conf(rclone_exe)
            chromium_binary = Path(args.chromium_binary) if args.chromium_binary else detect_chromium_binary()
            user_data_dir = Path(args.user_data_dir) if args.user_data_dir else detect_user_data_dir(chromium_binary)

            conf_text = read_text(rclone_conf)
            remote_name = detect_remote_name(conf_text, args.remote)
            remote_info = get_remote_info(conf_text, remote_name)
            host = get_remote_host(remote_info)

            log(f"rclone.exe: {rclone_exe}")
            log(f"rclone.conf: {rclone_conf}")
            log(f"browser binary: {chromium_binary}")
            log(f"browser user data: {user_data_dir}")
            log(f"browser profile directory: {args.profile_directory}")
            log(f"target remote: [{remote_name}]")
            log(f"derived host: {host}")
            log(f"remote url: {redact_url(remote_info.options.get('url', ''))}")

            current_fedauth, current_rtfa = parse_cookie_values_from_lines(remote_info.lines)
            if not args.force_refresh and current_fedauth and current_rtfa:
                log("Testing current cookie from rclone.conf...")
                validation = test_sharepoint_cookie(host, current_fedauth, current_rtfa)
                if validation.is_valid:
                    log("Current cookie is still valid. No refresh needed.")
                    return 0
                log(f"Current cookie failed validation. {format_validation_result(validation)}")
            elif args.force_refresh:
                log("Force refresh requested.")
            else:
                log("No existing cookie header found. Refreshing...")

            if args.dry_run:
                log("Dry run requested. No browser will be started and no config file will be modified.")
                return 0

            driver = build_driver(chromium_binary, user_data_dir, args.profile_directory, args.headless)
            try:
                fresh_fedauth, fresh_rtfa = fetch_sharepoint_cookies(
                    driver,
                    host,
                    timeout_seconds=args.login_timeout,
                    non_interactive=args.non_interactive,
                )
            finally:
                driver.quit()

            log("Validating freshly captured cookie...")
            fresh_validation = test_sharepoint_cookie(host, fresh_fedauth, fresh_rtfa)
            if not fresh_validation.is_valid:
                log(f"Fresh cookie failed validation. {format_validation_result(fresh_validation)}")
                return EXIT_REFRESH_VALIDATION_FAILED

            update_remote_cookie_header(rclone_conf, remote_name, fresh_fedauth, fresh_rtfa)
            log("Done.")
            return 0
    except ScriptError as exc:
        log(str(exc))
        return exc.exit_code
    except (FileNotFoundError, subprocess.CalledProcessError, OSError) as exc:
        log(str(exc))
        return EXIT_GENERAL_ERROR


class nullcontext:
    def __enter__(self):
        return None

    def __exit__(self, exc_type, exc, tb):
        return False


if __name__ == "__main__":
    sys.exit(main())
