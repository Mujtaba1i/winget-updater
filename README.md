# WinGet Updater

## Overview

This tool automates the process of updating installed applications using [WinGet](https://github.com/microsoft/winget-cli). It simplifies the update process, checking for available updates, and seamlessly upgrading applications.

## Features

- **Automatic WinGet Installation:** Checks if WinGet is installed and offers automatic installation if not.
- **Interactive Update:** User-friendly update process. Run the executable, hit "Yes" to grant admin privileges, and let the script do the work.
- **Temporary App Skipping:** Allows users to temporarily skip updating specific applications.
- **Detailed Log:** Generates a log file on the desktop with comprehensive information about the update process, including successes, skips, and failures.

## Prerequisites

- Windows OS
- Run as Administrator

## Usage

1. **Run the Executable:**
   - Double-click the `Winget_Updater_v.1.1.5.exe` file.
   - When prompted, click "Yes" to grant administrator privileges.

2. **Follow Instructions:**
   - The script will guide you through WinGet installation (if needed) and updating applications.

3. **Review Log (Optional):**
   - After the update process, a log file is generated on the desktop for detailed information.

## Note

- Ensure the script has the necessary permissions to execute. Right-click the executable and choose "Run as Administrator" if needed.

## License

This tool is released under the [MIT License](LICENSE).
