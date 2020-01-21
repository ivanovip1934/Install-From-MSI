using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Xml.XPath;

namespace Install_From_MSI
{
    class WriteEvents
    {

        private string pathToDirLog;

        public WriteEvents(string pathToDirLog) {
            this.pathToDirLog = pathToDirLog;
        }
        
        public void WriteLog(object sender, InstallEvent e) {

            string logFileName = $"{Environment.MachineName}.log";
            string appName = ((NewApp)sender).Name.Replace("\"", "").Replace(",", "").Replace(":", "");

            #region Create Directories for store Logs

            string  pathToDirLogApp = Path.Combine(pathToDirLog, appName);
            if (!Directory.Exists(pathToDirLogApp))
                try{
                    Directory.CreateDirectory(pathToDirLogApp);
                }
                catch (Exception expException){
                    Console.WriteLine(expException.Message);
                    Environment.Exit(0);
                }

            Dictionary<string, string> dicPathDirAndName = new Dictionary<string, string>() {
                { "pathToDirLogAppInstall", "install"},
                { "pathToDirLogAppInstallError", "error-install"},
                { "pathToDirLogAppUninstall","uninstall"},
                { "pathToDirLogAppUninstallError","error-uninstall"}
            };
            Dictionary<string, string> dicPathName = new Dictionary<string, string>();

            foreach (KeyValuePair<string, string> entry in dicPathDirAndName)
            {
                dicPathName.Add(entry.Key, Path.Combine(pathToDirLogApp, entry.Value));
                if (!Directory.Exists(dicPathName[entry.Key]))
                {
                    try
                    {
                        Directory.CreateDirectory(dicPathName[entry.Key]);
                    }
                    catch (Exception ecxException)
                    {
                        Console.WriteLine(ecxException.Message);
                        Environment.Exit(0);
                    }
                }
            }

            #endregion  
            
            string fileNameLog = String.Empty;
            string errorlog = String.Empty;

            switch (e.RJ) {

                case ResultJob.Install:
                    fileNameLog = Path.Combine(dicPathName["pathToDirLogAppInstall"], logFileName);
                    errorlog = Path.Combine(dicPathName["pathToDirLogAppInstallError"], logFileName);
                    if (File.Exists(errorlog))
                        File.Delete(errorlog);
                    break;
                case ResultJob.Uninstall:
                    fileNameLog = Path.Combine(dicPathName["pathToDirLogAppUninstall"], logFileName);
                    errorlog = Path.Combine(dicPathName["pathToDirLogAppUninstallError"], logFileName);
                    if (File.Exists(errorlog))
                        File.Delete(errorlog);
                    break;
                case ResultJob.ErrorInstall:
                    fileNameLog = Path.Combine(dicPathName["pathToDirLogAppInstallError"], logFileName);
                    break;
                case ResultJob.ErrorUninstall:
                    fileNameLog = Path.Combine(dicPathName["PathToDirLogAppUninstallError"], logFileName);
                    break;
                default:
                    break;
            }   
            
            File.Copy(e.TmpFileLog, fileNameLog, true);
            File.Delete(e.TmpFileLog);
        }

        int Cmd(string line) {

            Process newpc = Process.Start(
                new ProcessStartInfo {
                    FileName = "cmd",
                    Arguments = $"/c {line}",
                    WindowStyle = ProcessWindowStyle.Hidden
                });
            newpc.WaitForExit();
            int PcessExitCod = newpc.ExitCode;
            Console.WriteLine(PcessExitCod.ToString());
            return PcessExitCod;
        }



    }
}
