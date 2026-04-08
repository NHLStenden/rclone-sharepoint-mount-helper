# rclone-sharepoint-mount-helper

A small Windows helper for keeping an `rclone` SharePoint mount working.

It refreshes expired SharePoint session cookies from an already signed-in browser profile, updates the matching `rclone` remote, and can start the mount again when needed.

## What it helps with

`rclone` mounts that rely on SharePoint WebDAV cookies can stop working once those cookies expire.

This project automates the annoying part: refreshing the cookies and getting the mount back up without manually editing `rclone.conf`.

## Features

- Refreshes SharePoint cookies only when needed
- Updates the correct `rclone` remote automatically
- Creates a backup before changing `rclone.conf`
- Can start or recover an `rclone` mount
- Works with Chrome, Chromium, and Edge on Windows

## Requirements

- Windows
- Python 3.10 or newer
- `rclone`
- Chrome, Chromium, or Edge
- A browser profile that is already signed in to the SharePoint site you want to use

## Install

Install the Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Usage

Run the launcher:

```powershell
.\Start-RcloneSharePointMount.ps1
```

Use a specific remote:

```powershell
.\Start-RcloneSharePointMount.ps1 -RemoteName MySharePoint
```

Preview what would happen without changing anything:

```powershell
.\Start-RcloneSharePointMount.ps1 -DryRun
```

Force a refresh:

```powershell
.\Start-RcloneSharePointMount.ps1 -ForceRefresh
```

# Contributing

Contributions are welcome! Please fork the repository and submit pull requests.

# License

This project is licensed under the MIT License.

# Acknowledgements

NHL Stenden: For providing the foundational code and utilities.
Martin Bosgra: Author and primary maintainer of the project.
