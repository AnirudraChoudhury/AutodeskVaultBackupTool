using System.Configuration;
using System.Diagnostics;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;

namespace VaultBackup
{
    class Program
    {
        static void Main(string[] args)
        {
            // Read configuration values
            string tempPath = Environment.GetEnvironmentVariable("TEMP");
            string computerName = Environment.GetEnvironmentVariable("COMPUTERNAME");
            string vUser = ConfigurationManager.AppSettings["vUser"];
            string vPass = ConfigurationManager.AppSettings["vPass"];
            string sqlSaUser = ConfigurationManager.AppSettings["sqlSaUser"];
            string sqlSaPass = ConfigurationManager.AppSettings["sqlSaPass"];
            bool mailenabled = bool.Parse(ConfigurationManager.AppSettings["mailenabled"]);
            string mailserver = ConfigurationManager.AppSettings["mailserver"];
            string fromMail = ConfigurationManager.AppSettings["fromMail"];
            string toMail = ConfigurationManager.AppSettings["toMail"];
            int mailport = int.Parse(ConfigurationManager.AppSettings["port"]);
            string notificationSubject = ConfigurationManager.AppSettings["NotificationSubject"].Replace("%COMPUTERNAME%",computerName);
            string target = ConfigurationManager.AppSettings["target"].Replace("%TEMP%", tempPath).Replace("%DATE%", DateTime.Now.ToString("yyyyMMddHHmmss"));
            string oldTarget = ConfigurationManager.AppSettings["oldTarget"].Replace("%TEMP%", tempPath).Replace("%DATE%", DateTime.Now.ToString("yyyyMMddHHmmss"));
            int dayOfFullBackup = int.Parse(ConfigurationManager.AppSettings["dayOfFullBackup"]);
            int dayOfMoveOldBackup = int.Parse(ConfigurationManager.AppSettings["dayOfMoveOldBackup"]);
            bool archieveBackup = bool.Parse(ConfigurationManager.AppSettings["ArchieveBackup"]);
            bool runIncrementalBackup = bool.Parse(ConfigurationManager.AppSettings["runIncrementalBackup"]);
            bool runFullOnIncrFail = bool.Parse(ConfigurationManager.AppSettings["runFullOnIncrFail"]);
            string admsExe = ConfigurationManager.AppSettings["admsExe"];
            string log = ConfigurationManager.AppSettings["log"].Replace("%TEMP%", tempPath).Replace("%DATE%", DateTime.Now.ToString("yyyyMMddHHmmss"));
            string vBkuplog = ConfigurationManager.AppSettings["VBkuplog"].Replace("%TEMP%", tempPath).Replace("%DATE%", DateTime.Now.ToString("dd.MMM.yyyy-HH.mm"));
            string bRep = ConfigurationManager.AppSettings["BRep"].Replace("%TEMP%", tempPath).Replace("%DATE%", DateTime.Now.ToString("yyyyMMddHHmmss"));
            

            // Create log file
            File.Create(log).Close();

            //initiating the csv report
            AppendToCsv(bRep, "Step", "Description","StartTime","EndTime", "Status", "Remarks");
            // Run backup
            string backupType = ADSKVaultBackup(admsExe, target, oldTarget, log, vUser, vPass, sqlSaUser, sqlSaPass, vBkuplog, bRep, dayOfMoveOldBackup, dayOfFullBackup, archieveBackup, runIncrementalBackup, runFullOnIncrFail);

            if (mailenabled)
            {
                // Check services and send email
                CheckServices(log, bRep);
                string driveStatus = CheckDriveStorage();
                string mailMessage = CreateMailMessage(backupType, driveStatus, target, oldTarget, bRep, vBkuplog, dayOfMoveOldBackup, dayOfFullBackup, archieveBackup, runIncrementalBackup, runFullOnIncrFail);
                SendMail(mailserver, mailport, fromMail, toMail, notificationSubject, mailMessage, vBkuplog);
            }
        }

        static string ADSKVaultBackup(string admsExe, string target, string oldTarget, string log, string vUser, string vPass, string sqlSaUser, string sqlSaPass, string vBkuplog, string bRep, int dayOfMoveOldBackup, int dayOfFullBackup, bool archieveBackup, bool runIncrementalBackup, bool runFullOnIncrFail)
        {
            int currentDay = (int)DateTime.Now.DayOfWeek;
            bool isFullBackupDate = dayOfFullBackup == -1 || dayOfFullBackup == currentDay;

            if (dayOfMoveOldBackup == currentDay || dayOfFullBackup == -1)
            {
                if (archieveBackup)
                {
                    MoveOldBackup(target, oldTarget, bRep);
                }
                else
                {
                    DeleteLastBackup(target, bRep);
                }
            }

            if (isFullBackupDate)
            {
                FullBackup(admsExe, target, log, vUser, vPass, sqlSaUser, sqlSaPass, vBkuplog, bRep);
                return "A Full Vault Backup";
            }
            else if (runIncrementalBackup)
            {
                IncrementalBackup(admsExe, target, log, vUser, vPass, sqlSaUser, sqlSaPass, vBkuplog, bRep, runFullOnIncrFail);
                return "An Incremental Vault Backup";
            }
            else
            {
                return "No Backup";
            }
        }

        static void MoveOldBackup(string target, string oldTarget, string bRep)
        {
            string _startTime = DateTime.Now.ToString();
            try
            {
                if (Directory.Exists(target) && Directory.Exists(oldTarget))
                {
                    Directory.Delete(oldTarget, true);
                    Directory.Move(target, oldTarget);
                    Directory.CreateDirectory(target);
                    AppendToCsv(bRep, "Archieve Last Backup", "Delete Archieved Backup and Archieve Last Backup",_startTime, DateTime.Now.ToString(), "Success", "Old Backup Deleted Last Backup Renamed as Old");
                }
                else if (Directory.Exists(target) && !Directory.Exists(oldTarget))
                {
                    Directory.Move(target, oldTarget);
                    Directory.CreateDirectory(target);
                    AppendToCsv(bRep, "Archieve Last Backup", "Delete Archieved Backup and Archieve Last Backup", _startTime, DateTime.Now.ToString(), "Success", "Old Backup Not Found Last Backup Renamed as Old");
                }
                else
                { 
                    Directory.CreateDirectory(target);
                    AppendToCsv(bRep, "Archieve Last Backup", "Delete Archieved Backup and Archieve Last Backup", _startTime, DateTime.Now.ToString(), "Success", "No Last back up found");
                }
            }
            catch (Exception ex)
            {
                AppendToCsv(bRep, "Archieve Last Backup", "Delete Archieved Backup and Archieve Last Backup", _startTime, DateTime.Now.ToString(), "Failed", ex.Message);
            }
        }

        static void DeleteLastBackup(string target, string bRep)
        {
            string _startTime = DateTime.Now.ToString();
            try
            {
                if (Directory.Exists(target))
                {
                    long lastBackupSizeInBytes = Directory.GetFiles(target, "*", SearchOption.AllDirectories).Sum(file => new FileInfo(file).Length);
                    Directory.Delete(target, true);
                    Directory.CreateDirectory(target);
                    double lastBackupSize = Math.Round(lastBackupSizeInBytes / (double)(1024 * 1024 * 1024), 2);
                    AppendToCsv(bRep, "Delete Last Backup", "Delete Last Backup instead of archieving as space is constrained", _startTime, DateTime.Now.ToString(), "Success", $"Last Backup with {lastBackupSize} GB Deleted and the space is freed up to start new backup");
                }
                else
                {
                    Directory.CreateDirectory(target);
                    AppendToCsv(bRep, "Delete Last Backup", "Delete Last Backup instead of archieving as space is constrained", _startTime, DateTime.Now.ToString(), "Success", "No Last back up found so the folder is created");
                }
            }
            catch (Exception ex)
            {
                AppendToCsv(bRep, "Delete Last Backup", "Delete Last Backup instead of archieving as space is constrained", _startTime, DateTime.Now.ToString(), "Failed", ex.Message);
            }
        }

        static void FullBackup(string admsExe, string target, string log, string vUser, string vPass, string sqlSaUser, string sqlSaPass, string vBkuplog, string bRep)
        {
            string _startTime = DateTime.Now.ToString();
            try
            {
                AppendToLog(log, "Stopping Process: Connectivity.ADMSConsole");
                StopProcessIfRunning("Connectivity.ADMSConsole.exe");
                AppendToLog(log, $"=== Backup script start {DateTime.Now}");
                AppendToLog(log, $"Backup will now begin {DateTime.Now}");
                AppendToLog(log, $"====== Full backup started {DateTime.Now}:");

                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = admsExe,
                    Arguments = $"-Obackup -B{target} -VU{vUser} -VP{vPass} -DBU{sqlSaUser} -DBP{sqlSaPass} -S -L{vBkuplog} -DBSC -INRF",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process process = Process.Start(processStartInfo))
                {
                    process.WaitForExit();
                    AppendToLog(log, $"=== BACKUP SCRIPT FINISHED {DateTime.Now}");
                    long backupFolderSizeInBytes = Directory.GetFiles(target, "*", SearchOption.AllDirectories).Sum(file => new FileInfo(file).Length);
                    double backupSize = Math.Round(backupFolderSizeInBytes / (double)(1024 * 1024 * 1024), 2);
                    AppendToCsv(bRep, "Full Backup", "Running ADMS Console Backup Command to do the backup", _startTime, DateTime.Now.ToString(), "Success", $"Backup completed and the size of backup = {backupSize} GB");
                }
            }
            catch (Exception ex)
            {
                AppendToCsv(bRep, "Full Backup", "Running ADMS Console Backup Command to do the backup", _startTime, DateTime.Now.ToString(), "Failed", ex.Message);
            }
        }

        static void IncrementalBackup(string admsExe, string target, string log, string vUser, string vPass, string sqlSaUser, string sqlSaPass, string vBkuplog, string bRep, bool runFullOnIncrFail)
        {
            string _startTime = DateTime.Now.ToString();
            try
            {
                AppendToLog(log, "Stopping Process: Connectivity.ADMSConsole");
                StopProcessIfRunning("Connectivity.ADMSConsole.exe");
                AppendToLog(log, $"=== Backup script start {DateTime.Now}");
                AppendToLog(log, $"Backup will now begin {DateTime.Now}");
                AppendToLog(log, $"====== Incremental backup started {DateTime.Now}:");

                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = admsExe,
                    Arguments = $"-Obackup -B{target} -VU{vUser} -VP{vPass} -DBU{sqlSaUser} -DBP{sqlSaPass} -S -L{vBkuplog} -DBSC -INC -INRF",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process process = Process.Start(processStartInfo))
                {
                    process.WaitForExit();
                    AppendToLog(log, $"=== BACKUP SCRIPT FINISHED {DateTime.Now}");

                    bool incrError = File.ReadAllLines(vBkuplog).Any(line => line.Contains("A Full Backup must occur") || line.Contains("change is not supported by an Incremental Backup"));

                    if (incrError)
                    {
                        AppendToCsv(bRep, "Incremental Backup", "Running ADMS Console Backup Command to do the backup", _startTime, DateTime.Now.ToString(), "Failed", "Incremental Backup Failed.");
                        if (runFullOnIncrFail)
                        {
                            AppendToLog(log, $"=== Starting Full backup as incremental backup failed {DateTime.Now}");
                            FullBackup(admsExe, target, log, vUser, vPass, sqlSaUser, sqlSaPass, vBkuplog, bRep);
                        }
                    }
                    else
                    {
                        long backupFolderSizeInBytes = Directory.GetFiles(target, "*", SearchOption.AllDirectories).Sum(file => new FileInfo(file).Length);
                        double backupSize = Math.Round(backupFolderSizeInBytes / (double)(1024 * 1024 * 1024), 2);
                        AppendToCsv(bRep, "Incremental Backup", "Running ADMS Console Backup Command to do the backup", _startTime, DateTime.Now.ToString(), "Success", $"Backup completed and the size of backup = {backupSize} GB");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendToCsv(bRep, "Incremental Backup", "Running ADMS Console Backup Command to do the backup", _startTime, DateTime.Now.ToString(), "Failed", ex.Message);
            }
        }

        static void CheckServices(string log, string bRep)
        {
            bool status = true;
            string _startTime = DateTime.Now.ToString();
            StringBuilder remarks = new StringBuilder();

            try
            {
                ServiceController iisAdmin = new ServiceController("IISADMIN");
                ServiceController w3svc = new ServiceController("W3SVC");
                ServiceController jobDispatch = new ServiceController("Autodesk Data Management Job Dispatch");

                if (w3svc.Status == ServiceControllerStatus.Stopped)
                {
                    w3svc.Start();
                    AppendToLog(log, "W3SVC Service has been restarted");
                    remarks.Append("W3SVC Started");
                }
                else
                {
                    AppendToLog(log, "World Wide Web service is running");
                    remarks.Append("W3SVC Running");
                }

                if (iisAdmin.Status == ServiceControllerStatus.Stopped)
                {
                    Process.Start("iisreset", "/start")?.WaitForExit();
                    AppendToLog(log, "IISADMIN service has been restarted");
                    remarks.Append(" | IIS Service Started");
                }
                else
                {
                    AppendToLog(log, "IIS Admin service is running");
                    remarks.Append(" | IIS Service Running");
                }

                if (jobDispatch.Status == ServiceControllerStatus.Stopped)
                {
                    jobDispatch.Start();
                    AppendToLog(log, "AUTODESK Job Dispatch has been restarted");
                    remarks.Append(" | Autodesk job dispatch Service Started");
                }
                else
                {
                    AppendToLog(log, $"=== Job Dispatch service is running {DateTime.Now}");
                    remarks.Append(" | Job Dispatch service is running");
                }

                AppendToCsv(bRep, "Check Service Status", "Checking IIS & W3SVC & Autodesk job dispatch Service", _startTime, DateTime.Now.ToString(), "Success", remarks.ToString());
            }
            catch (Exception ex)
            {
                status = false;
                AppendToCsv(bRep, "Check Service Status", "Checking IIS & W3SVC & Autodesk job dispatch Service", _startTime, DateTime.Now.ToString(), "Failed", ex.Message);
            }
        }

        static string CheckDriveStorage()
        {
            StringBuilder notificationMessage = new StringBuilder();
            notificationMessage.Append($"<h1>Drive wise storage for {Environment.MachineName}</h1>");
            notificationMessage.Append("<table border='1'><tr><th>Drive</th><th>Total Space</th><th>Free Space</th></tr>");

            foreach (DriveInfo drive in DriveInfo.GetDrives().Where(d => d.IsReady))
            {
                double totalSpace = Math.Round(drive.TotalSize / (double)(1024 * 1024 * 1024), 2);
                double freeSpace = Math.Round(drive.TotalFreeSpace / (double)(1024 * 1024 * 1024), 2);
                notificationMessage.Append($"<tr><td>{drive.Name}</td><td>{totalSpace} GB</td><td>{freeSpace} GB</td></tr>");
            }

            notificationMessage.Append("</table>");
            return notificationMessage.ToString();
        }

        static string CreateMailMessage(string backupType, string driveStatus, string target, string oldTarget, string backupReport, string vBkuplog, int dayOfMoveOldBackup, int dayOfFullBackup, bool archieveBackup, bool runIncrementalBackup, bool runFullOnIncrFail)
        {
            StringBuilder mailMessage = new StringBuilder();
            mailMessage.Append($"Greetings! <p>{backupType} is finished on the <strong>{Environment.MachineName}</strong>.</p> <p>The Log of the backup is attached with this mail.</p>");
            mailMessage.Append(GetBackupSettings(target, oldTarget, backupReport, vBkuplog, dayOfMoveOldBackup, dayOfFullBackup, archieveBackup, runIncrementalBackup, runFullOnIncrFail));
            mailMessage.Append("<h1> Vault Backup Summary</h1>");
            mailMessage.Append(CsvToHtmlTable(backupReport));
            mailMessage.Append(driveStatus);
            return mailMessage.ToString();
        }

        static string GetBackupSettings(string target, string oldTarget, string log, string vBkuplog, int dayOfMoveOldBackup, int dayOfFullBackup, bool archieveBackup, bool runIncrementalBackup, bool runFullOnIncrFail)
        {
            StringBuilder settingsMessage = new StringBuilder();
            settingsMessage.Append("<h1> Here are the backup Settings for " + Environment.MachineName + " </h1>");
            settingsMessage.Append("<table border='1'><tr><th>Parameter</th><th>Value</th></tr>");
            settingsMessage.Append($"<tr><td>Backup Location</td><td>{target}</td></tr>");
            settingsMessage.Append($"<tr><td>Archeive the last backup</td><td>{archieveBackup}</td></tr>");
            settingsMessage.Append($"<tr><td>Archeive Location</td><td>{oldTarget}</td></tr>");
            settingsMessage.Append($"<tr><td>Full Backup Day</td><td>{dayOfFullBackup}</td></tr>");
            settingsMessage.Append($"<tr><td>Move Old Backup Day</td><td>{dayOfMoveOldBackup}</td></tr>");
            settingsMessage.Append($"<tr><td>Incremental Backup Enabled</td><td>{runIncrementalBackup}</td></tr>");
            settingsMessage.Append($"<tr><td>Full Backup on Failed incremental Backup</td><td>{runFullOnIncrFail}</td></tr>");
            settingsMessage.Append($"<tr><td>Script Log Path</td><td>{log}</td></tr>");
            settingsMessage.Append($"<tr><td>Vault Backup log path</td><td>{vBkuplog}</td></tr>");
            settingsMessage.Append("</table>");
            return settingsMessage.ToString();
        }

        static string CsvToHtmlTable(string csvPath)
        {
            StringBuilder html = new StringBuilder();
            html.Append("<table border='1'>");

            using (StreamReader reader = new StreamReader(csvPath))
            {
                string line;
                bool isHeader = true;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] values = line.Split(',');
                    html.Append("<tr>");
                    foreach (string value in values)
                    {
                        html.Append(isHeader ? $"<th>{value}</th>" : $"<td>{value}</td>");
                    }
                    html.Append("</tr>");
                    isHeader = false;
                }
            }

            html.Append("</table>");
            return html.ToString();
        }

        static void AppendToLog(string logPath, string message)
        {
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine(message);
            }
        }

        static void AppendToCsv(string csvPath, string step, string description, string startTime, string endTime, string status, string remarks)
        {
            using (StreamWriter writer = new StreamWriter(csvPath, true))
            {
                writer.WriteLine($"{step},{description},{startTime},{endTime},{status},{remarks}");
            }
        }

        static void SendMail(string server, int port, string fromEmail, string toEmail, string subject, string body, string attachmentPath)
        {
            try
            {
                SmtpClient smtpClient = new SmtpClient(server, port);
                MailMessage mailMessage = new MailMessage(fromEmail, toEmail)
                {
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true
                };

                if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
                {
                    mailMessage.Attachments.Add(new Attachment(attachmentPath));
                }

                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void StopProcessIfRunning(string processName)
        {
            var processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(processName));
            if (processes.Any())
            {
                Process.Start("taskkill", $"/F /IM {processName}")?.WaitForExit();
            }
        }
    }
}