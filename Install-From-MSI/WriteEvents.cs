using System;
using System.Collections.Generic;
using System.IO;

namespace Install_From_MSI
{
    class WriteEvents
    {

        private string pathToDirLog;
        public WriteEvents(string pathToDirLog) {
            this.pathToDirLog = pathToDirLog;
        }
        
        public void WriteLog(object sender, InstallEvent e) {

            // Set VARIABLES
            string logFileName = $"{Environment.MachineName}.log";
            string appName = ((NewApp)sender).Name.Replace("\"", "").Replace(",", "").Replace(":", "");
            string tmpFileLog = ((NewApp)sender).PathToTempFileLog;
            string fileNameLog = String.Empty;
            string olderrorlog = String.Empty;
            string pathToDirLogApp = Path.Combine(pathToDirLog, appName);

            Dictionary<ResultJob, string> dicPathName = new Dictionary<ResultJob, string>() {
                { ResultJob.Install, Path.Combine(pathToDirLogApp,"install")},
                { ResultJob.ErrorInstall, Path.Combine(pathToDirLogApp, "error-install")},
                { ResultJob.Uninstall,Path.Combine(pathToDirLogApp, "uninstall")},
                { ResultJob.ErrorUninstall,Path.Combine(pathToDirLogApp, "error-uninstall")}
            };

            // END set VARIABLES

            #region Create Directories for store Logs

            if (!Directory.Exists(pathToDirLogApp))
                try{
                    Directory.CreateDirectory(pathToDirLogApp);
                }
                catch (Exception expException){
                    Console.WriteLine(expException.Message);
                    Environment.Exit(0);
                }           

            #endregion

            if (e.RJ == ResultJob.Install | e.RJ == ResultJob.Uninstall)
            {
                fileNameLog = Path.Combine(dicPathName[e.RJ], logFileName);
                olderrorlog = Path.Combine(dicPathName[e.RJ + 2], logFileName);
                if (File.Exists(olderrorlog))
                    File.Delete(olderrorlog);
            }
            else
                fileNameLog = Path.Combine(dicPathName[e.RJ], logFileName);

            File.Copy(tmpFileLog, fileNameLog, true);
            File.Delete(tmpFileLog);
        }

    }
}
