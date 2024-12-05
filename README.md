# VaultBackup Tool

## Overview
VaultBackup is a robust console application designed to automate the backup process for Autodesk Vault. It supports both full and incremental backups, with features like automatic log creation, backup archiving, and email notifications for backup status.

## Features
- **Full and Incremental Backups**  
  Automates both types of backups based on configured schedules.

- **Log Management**  
  Generates detailed logs and backup reports in CSV format.

- **Service Monitoring**  
  Automatically checks and restarts critical services like IIS and Autodesk Data Management Job Dispatch.

- **Email Notifications**  
  Sends detailed backup reports, including drive storage status and backup logs.

## Prerequisites
- .NET Framework
- Autodesk Vault Server
- SMTP Server for email notifications (optional)

## Configuration
All configurations are managed through `App.config`. Below are the configuration keys:

| Key                     | Description                                                                                     |
|-------------------------|-------------------------------------------------------------------------------------------------|
| `vUser`, `vPass`        | Vault credentials.                                                                              |
| `sqlSaUser`, `sqlSaPass`| SQL Server credentials.                                                                         |
| `mailenabled`           | Enable or disable email notifications (`true`/`false`).                                         |
| `mailserver`            | SMTP server address.                                                                            |
| `fromMail`, `toMail`    | Sender and recipient email addresses.                                                           |
| `port`                  | SMTP server port.                                                                               |
| `NotificationSubject`   | Email subject line. Supports placeholders like `%COMPUTERNAME%`.                                |
| `target`, `oldTarget`   | Paths for current and archived backups. Supports placeholders like `%TEMP%` and `%DATE%`.       |
| `dayOfFullBackup`       | Day for full backup (0=Sunday, 1=Monday, ..., 6=Saturday, -1=Always).                           |
| `dayOfMoveOldBackup`    | Day for moving or deleting old backups.                                                         |
| `ArchieveBackup`        | Determines whether old backups should be archived (`true`) or deleted (`false`).                |
| `runIncrementalBackup`  | Enable incremental backups (`true`/`false`).                                                    |
| `runFullOnIncrFail`     | Perform a full backup if incremental backup fails (`true`/`false`).                             |
| `admsExe`               | Path to the ADMS Console executable.                                                            |
| `log`, `VBkuplog`, `BRep` | Paths for log, vault backup log, and CSV report. Supports `%TEMP%` and `%DATE%`.             |

## Usage
1. **Compile and Run**  
   Build the project and run the executable.

2. **Backup Execution**  
   The program will determine whether to perform a full or incremental backup based on the day of the week and configuration.

3. **Email Notification** (Optional)  
   If enabled, the tool will send a detailed report upon backup completion.

## Logs and Reports
- **Log File**: Contains detailed runtime logs.
- **Backup Report (CSV)**: Summarizes backup steps with timestamps, statuses, and remarks.

## Troubleshooting
- Ensure the configuration values in `App.config` are correct.
- Verify that all required services (IIS, Autodesk Job Dispatch) are running.

## Future Improvements
- Add support for cloud storage backups.
- Introduce a web-based dashboard for monitoring backup status.

## License
This project is licensed under the MIT License. See `LICENSE` for more details.

## Contact
For any issues or feature requests, contact Anirudra Choudhury at choudhury.anirudra@gmail.com.
