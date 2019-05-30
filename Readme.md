# NYPTI Outlook Drag Drop

*Drag and drop Outlook items as files into any application*

## Overview

NYPTI Outlook Drag Drop is an add-in for Outlook 2013, 2016, and 365 that allows dragging and dropping Outlook items (messages, attachments, contacts, tasks, appointments, meetings, etc) to applications and web browsers (i.e. Chrome, Edge, FireFox) that allow physical files (saved-on-computer) to be dropped into them. Dragging an Outlook item (e.g. attachment) to a browser will temporarily save the file on the computer and remove it once the file has been uploaded.

## How Does it Work?

When you try to drag and drop from Outlook, Outlook correctly identifies the
format as virtual files (CFSTR_FILEDESCRIPTORW) since the files do not exist
directly on disk.  Instead, they are contained in a PST file, OST file, or on
an Exchange server.

However, many applications do not support this format, such as web browers and
most .NET/Java applications.

To work around this issue, NYPTI Outlook Drag Drop hooks the Outlook drag and drop
process and adds support for physical files (CF_HDROP).  When the receiving
application asks for the physical files, the files are saved to a temp folder
and those filenames are returned to the application.  The application processes
the files (such as uploading them).  NYPTI Outlook Drag Drop deletes the temp files
later in a cleanup process.

## Features

- Works with Chrome, Firefox, Internet Explorer, Edge, and other applications that accept files to be dropped
- Allows drag and drop into HTML5-based web applications
- Drag e-mails, attachments, contacts, calendar items, and more
- Drag multiple items at once
- Supports Unicode characters

## Installation

To install, run the installer that matches your Windows build by donwloading the installer from [the latest github release page](https://github.com/NYPTI/Drag-Drop/releases/latest) ~~or from the Microsoft Store.~~ (not yet)

- The installer for 64-bit Windows works with Outlook 32-bit or 64-bit
- The installer for 32-bit Windows works with Outlook 32-bit or 64-bit

After installing, restart Outlook for the add-in to take effect.

## Automated (Silent) Installation

For administrators, NYPTI Outlook Drag Drop supports automated (silent) installation and uninstallation using `msiexec` with command line parameters.

### Silent Installation

To silently install NYPTI Outlook Drag Drop, use this command:

`msiexec.exe /i <pathtomsi> /qn /log <pathtolog>`

- `<pathtomsi>`: Path to MSI file
- `<pathtolog>`: Path to log file (if folder is not specified, MSI path is used)

Example:

`msiexec.exe C:\Install\NYPTIOutlookDragDropSetup.msi /qn /log C:\Logs\NYPTIOutlookDragDropInstall.log`

After installing, restart Outlook for the add-in to take effect.

### Silent Uninstallation

To silently uninstall NYPTI Outlook Drag Drop, use this command:

`msiexec.exe /x <productcode> /qn /log <pathtolog>`

- `<productcode>` for 64-bit version: `{5E00C640-C7C4-43AB-AD7F-68DA4A72B4A7}`
- `<productcode>` for 32-bit version: `{B1E4DA5D-16B7-45E5-AFB6-3DCE3A24B083}`
- `<logfile>`: Path to log file

Example:

`msiexec.exe /x {B1E4DA5D-16B7-45E5-AFB6-3DCE3A24B083} /qn /log C:\Logs\NYPTIOutlookDragDropUninstall.log`

## Acknowledgements

NYPTI Outlook Drag Drop is based on [Outlook File Drag by tonyfederer](https://github.com/tonyfederer/OutlookFileDrag) using his MIT license, and was repackaged by NYPTI intern [Julien Cheng](https://github.com/julien-cheng) (R.P.I. - Class of 2020). See [Outlook File Drag's](https://github.com/tonyfederer/OutlookFileDrag) license below:
```
Copyright (c) 2018 Tony Federer

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
NYPTI Outlook Drag Drop also uses these open source projects:

- [Easyhook](https://easyhook.github.io/)
- [log4net](http://logging.apache.org/log4net/)

## Feedback/Contribute

You can view the source code, report issues, and contribute on [Github](https://github.com/NYPTI/Outlook-Drag-Drop).


## Version History

### 2.0.0
- Signed Installer Binaries for Windows x86
- Renamed Software
- Removed x64 installer solution
- Added icons and banners to installer
- Added capability to drag attachments from RTF email bodies

### 1.0.10
- Fixed System.ArgumentException bug in ReadHGlobalIntoStream method when reading more than 4 KB introduced in version 1.0.8.

### 1.0.9
- If files were dropped and drop effect was "move", then override to "copy" so original item is not deleted

### 1.0.8
- Fixed releasing of unmanaged resources
- Memory usage improvements
- Added more details to log file

### 1.0.5
- Fixed crash when dragging calendar items

### 1.0.4
- Added additional debug logging
- Fixed issue where STGMEDIUM was not being released after reading filenames
- Fixed issue that where reading filenames sometimes failed
- Fixed hooking process to allow starting and stopping hook without disposing and recreating hook

### 1.0.3
- Fixed issue that prevented dragging items from one group to another

### 1.0.2
- Fixed PathTooLong exception when temporary filename was longer than MAX_PATH

### 1.0.1
- Fixed issues with 64-bit Outlook
- Added self-signed certificate

### 1.0
- Initial Release

## Copyright
NYPTI Outlook Drag Drop is copyright (c) 2019 by [NYPTI](https://github.com/NYPTI) / [New York Prosecutors Training Institute](https://www.nypti.org/) and released under the MIT License.
