using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace Install_From_MSI
{
    class WriteEvents
    {

        private string appname;
        private string pathToDirLog;
        private string pathToDirLogApp;
        private string pathToDirLogAppInstall;
        private string pathToDirLogAppInstallError;
        private string pathToDirLogAppUninstall;
        private string PathToDirLogAppUninstallError;
        private string logFileName;


        public WriteEvents(string pathToDirLog ,string appName) {
            this.pathToDirLog = pathToDirLog;            
          
            
            

            appname = appName;

            pathToDirLogApp = Path.Combine(pathToDirLog, appName);
            if (!Directory.Exists(pathToDirLogApp))
                try {
                    Directory.CreateDirectory(pathToDirLogApp);
                }
                catch (Exception e) {
                    Console.WriteLine(e.Message);
                    Environment.Exit(0);
                }

            pathToDirLogAppInstall = Path.Combine(pathToDirLogApp, "install");
            if (!Directory.Exists(pathToDirLogAppInstall))
                try {
                    Directory.CreateDirectory(pathToDirLogAppInstall);
                }
                catch (Exception e) {
                    Console.WriteLine(e.Message);
                    Environment.Exit(0);
                }
            pathToDirLogAppInstallError = Path.Combine(pathToDirLogApp, "error-install");
            if (!Directory.Exists(pathToDirLogAppInstallError))
                try {
                    Directory.CreateDirectory(pathToDirLogAppInstallError);
                }
                catch (Exception e) {
                    Console.WriteLine(e.Message);
                    Environment.Exit(0);
                }

            pathToDirLogAppUninstall = Path.Combine(pathToDirLogApp, "uninstall");
            if (!Directory.Exists(pathToDirLogAppUninstall))
                try {
                    Directory.CreateDirectory(pathToDirLogAppUninstall);
                }
                catch (Exception e) {
                    Console.WriteLine(e.Message);
                    Environment.Exit(0);
                }
            PathToDirLogAppUninstallError = Path.Combine(pathToDirLogApp, "error-uninstall");
            if (!Directory.Exists(PathToDirLogAppUninstallError))
                try {
                    Directory.CreateDirectory(PathToDirLogAppUninstallError);
                }
                catch (Exception e) {
                    Console.WriteLine(e.Message);
                    Environment.Exit(0);
                }

            logFileName = $"{Environment.MachineName}.log";
                                      
        }


        public void WriteLog(object sender, InstallEvent e) {

            string fileNameLog = "";

            switch (e.RJ) {

                case ResultJob.Install:
                    fileNameLog = Path.Combine(pathToDirLogAppInstall, logFileName);
                    break;
                case ResultJob.Uninstall:
                    fileNameLog = Path.Combine(pathToDirLogAppUninstall, logFileName);
                    break;
                case ResultJob.ErrorInstall:
                    using (StreamWriter writer = new StreamWriter(@"c:\programs\erronum.txt")) {
                        writer.WriteLine(e.Message);
                    }
                    fileNameLog = Path.Combine(pathToDirLogAppInstallError, logFileName);
                    break;
                case ResultJob.ErrorUninstall:
                    using (StreamWriter writer = new StreamWriter(@"c:\programs\erronumuninstall.txt", true)) {
                        writer.WriteLine(e.Message);
                    }
                    fileNameLog = Path.Combine(PathToDirLogAppUninstallError, logFileName);
                    break;
                default:
                    break;
            }   
            
            File.Copy(e.TmpFileLog, fileNameLog, true);
            File.Delete(e.TmpFileLog);
        }


        #region OLD METHOD Install, InstallError, Uninstall, UninstallError

        //public void Install(object sender, InstallEvent e) {

        //    string FileNameLog = Path.Combine(pathToDirLogAppInstall, LogFileName);
        //    File.Copy(e.TmpFileLog, FileNameLog, true);
        //    File.Delete(e.TmpFileLog);
        //}

        //public void InstallError(object sender, InstallEvent e) {

        //    using (StreamWriter writer = new StreamWriter(@"c:\programs\erronum.txt")) {
        //        writer.WriteLine(e.Message);
        //    }
        //    string FileNameLog = Path.Combine(pathToDirLogAppInstallError, LogFileName);
        //    File.Copy(e.TmpFileLog, FileNameLog, true);
        //    File.Delete(e.TmpFileLog);
        //}

        //public void Uninstall(object sender, InstallEvent e) {

        //    string FileNameLog = Path.Combine(pathToDirLogAppUninstall, LogFileName);
        //    File.Copy(e.TmpFileLog, FileNameLog, true);
        //    File.Delete(e.TmpFileLog); 
        //}

        //public void UninstallError(object sender, InstallEvent e) {

        //    using (StreamWriter writer = new StreamWriter(@"c:\programs\erronumuninstall.txt",true)) {
        //        writer.WriteLine(e.Message);
        //    }
        //    string FileNameLog = Path.Combine(PathToDirLogAppUninstallError, LogFileName);
        //    File.Copy(e.TmpFileLog, FileNameLog, true);
        //    File.Delete(e.TmpFileLog);
        //}

        #endregion










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
