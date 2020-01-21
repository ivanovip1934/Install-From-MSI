using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Install_From_MSI
{
    [Serializable]
    class NewApp {

        


        public string Name { get; private set; }
        public string PathMsi { get; private set; }
        public bool x86AppOnX64Os { get; private set; }
        public string Version { get; private set; }
        public bool Update { get; private set; }
        public bool ForceRestart { get; private set; }
        public string Property { get; set; }

        private Random rnd;
        public string PathToTempFileLog{ get; private set; }

        public event EventHandler<InstallEvent> EventLog;
        
        public NewApp(string pathX64, string pathX86,string version, bool update, bool forceRestart, string property) {
            Version = version;
            PathMsi = GetPathMSI(pathX64, pathX86);
            Update = update;
            ForceRestart = forceRestart;
            Property = property;
        }

        public NewApp(Options options) {

            PathMsi = GetPathMSI(options.PathX64, options.PathX86);
            x86AppOnX64Os = options.x86AppOnX64Os;
            Version = (!string.IsNullOrEmpty(options.Version)) ? options.Version : GetInfoMSI(PathMsi, "ProductVersion");
            Name = GetInfoMSI(PathMsi, "ProductName");
            //Version = (!string.IsNullOrEmpty(options.Version)) ? options.Version: GetMSIInfo(PathMsi, "ProductVersion");
            Update = options.Update;
            ForceRestart = options.ForceRestart;
            Property = options.Property;
            rnd = new Random();
            PathToTempFileLog = Path.Combine(Environment.GetEnvironmentVariable("TEMP", EnvironmentVariableTarget.Machine), $"LogInstallApp{rnd.Next(1000, 2000)}.log");
            

        }


        string GetPathMSI(string pathX64, string pathX86) {
            bool IsOSX64 = Environment.Is64BitOperatingSystem;
            switch (IsOSX64) {
                case true:
                    if (!String.IsNullOrEmpty(pathX64)) {
                        if (File.Exists(pathX64))
                            return pathX64;
                        else {
                            Console.WriteLine($"MSI: {pathX64} NOT exists");
                            Environment.Exit(0);
                        }

                    } else {
                        Console.WriteLine("Ошибка: Путь до MSI пакета для x64 OS не задан");
                        Console.WriteLine("Ошибка: Установка не возможна.");
                        Environment.Exit(0);

                    }
                    break;
                case false:
                    if (!String.IsNullOrEmpty(pathX86)) {
                        if (File.Exists(pathX86))
                            return pathX86;
                        else {
                            Console.WriteLine($"MSI: {pathX86} NOT exists");
                            Environment.Exit(0);
                        }
                    } else {
                        Console.WriteLine("Ошибка: Путь до MSI пакета для x86 OS не задан");
                        Console.WriteLine("Ошибка: Установка не возможна.");
                        Environment.Exit(0);
                    }
                    break;

                default:
                    break;
            }            
            return "privet";
        }

        string GetInfoMSI(string fileName, string property) {
            string STR = String.Empty;
            GetMsiInfo msiinfo = new GetMsiInfo();
            msiinfo.GetMSIInfo(fileName, property);
            //Thread tread = new Thread(() => msiinfo.GetMSIInfo(fileName, property));
            //tread.Start();
            //Thread.Sleep(300);
            
            
            STR = msiinfo.ValueProperty;
            
            return STR;

        }

        public bool IsNeedInstall(CurApp curapp) {
            string[] arrVersionCurApp = curapp.DisplayVersion.Trim().Split('.');
            string[] arrversionMSI = Version.Trim().Split('.');
            int count = (arrVersionCurApp.Count() >= arrversionMSI.Count()) ? arrversionMSI.Count() : arrVersionCurApp.Count();

            for (int i = 0; i < count; i++) {
                Console.WriteLine($"arrversionMSI[{i}] = {int.Parse(arrversionMSI[i])} arrvDisplayVersion[{i}] = {int.Parse(arrVersionCurApp[i])}");
                if (int.Parse(arrversionMSI[i]) > int.Parse(arrVersionCurApp[i])) {
                    return true;
                }
                    
                else if ((int.Parse(arrversionMSI[i]) < int.Parse(arrVersionCurApp[i])))
                    return false;
            }

            return false;
        }


        public void InstallApp() {
            // 1. Проверяем: есть ли установленная версия данной программы?
            CurApp curapp = ListProg();


            if (curapp != null) {
                Console.WriteLine($"Cur app version {curapp.DisplayVersion}");
                Console.WriteLine($"Cur app UninstallString {curapp.UninstallString}");
                if (!IsNeedInstall(curapp)) {
                    // Если версия установленного приложения больше или равно версии msi - то просто выходим из приложения.
                    Console.WriteLine("программа не требует обновления  - завершаем работу");
                    Environment.Exit(0);
                }
                else {
                    // если Update = false - то удаляем текущую установленную программу.
                    if (!Update) {
                        UninstallApp(curapp);
                        if (ForceRestart)
                            Environment.Exit(0);
                    }
                }
            }




            string installString = $"msiexec /i \"{PathMsi}\" /quiet";
            if (ForceRestart)
                installString = installString + " /forcerestart";
            installString = installString + $" /log {PathToTempFileLog}";

            int numinst = Cmd(installString);
            switch (numinst) { 
                case 0:
                EventLog?.Invoke(this, new InstallEvent("Event: программа установлена", PathToTempFileLog, ResultJob.Install));
                    break;
                case 1641:
                    EventLog?.Invoke(this, new InstallEvent("Event: программа установлена и требуется перезагрузка", PathToTempFileLog, ResultJob.Install));
                    Environment.Exit(0);
                    break;
                default:
                    EventLog?.Invoke(this, new InstallEvent($"Event: ошибка № {numinst} при установки программы", PathToTempFileLog, ResultJob.ErrorInstall));
                    Environment.Exit(0);
                    break;
            }

            if (ForceRestart)
                Environment.Exit(0);

        }

                        


         void UninstallApp(CurApp curapp) {
            string uninstallString = curapp.UninstallString.ToLower();

            #region Удаляем Msi пакеты используя msiexec
            if (!string.IsNullOrEmpty(uninstallString) && uninstallString.Contains("msiexec")) {
                Console.WriteLine($"Будем удалять {curapp.DisplayName} версии {curapp.DisplayVersion} используя msiexec");
                if (uninstallString.Contains("/i")) {
                    Console.WriteLine("Vmesto udaleniy zapustinza ustanovka");
                    uninstallString = uninstallString.ToLower().Replace("/i", "/x");
                    Console.WriteLine($"New uninstallstring: {uninstallString}");

                    Console.ReadKey();

                }

                if (!uninstallString.Contains("quiet"))
                    uninstallString = uninstallString + " /quiet";

                if (ForceRestart & !uninstallString.Contains("forcerestart"))
                    uninstallString = uninstallString + " /forcerestart";
                uninstallString = uninstallString + $" /log {PathToTempFileLog}";
                int numuninst = Cmd(uninstallString);

                switch (numuninst) {
                    case 0:
                        EventLog?.Invoke(this, new InstallEvent("Event: программа удалена", PathToTempFileLog, ResultJob.Uninstall));
                        break;
                    case 1641:
                        EventLog?.Invoke(this, new InstallEvent("Event: программа установлена и требуется перезагрузка", PathToTempFileLog, ResultJob.Uninstall));
                        Environment.Exit(0);
                        break;
                    default:
                        EventLog?.Invoke(this, new InstallEvent($"Event: ошибка № {numuninst} при удалении программы", PathToTempFileLog, ResultJob.ErrorUninstall));
                        Environment.Exit(0);
                        break;
                }                

            }
            #endregion

        }


        CurApp ListProg() {
            string registry_key = String.Empty;
            registry_key = x86AppOnX64Os?
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall":
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            
            if (string.IsNullOrEmpty(registry_key) || string.IsNullOrEmpty(this.Name)) {
                Console.WriteLine("reg path or appName is null");
                Environment.Exit(0);
            }

            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key)) {
                
                foreach (string subkey_name in key.GetSubKeyNames()) {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name)) {

                        if (subkey.GetValue("DisplayName") != null && subkey.GetValue("DisplayName").ToString().Contains(Name)) {

                            string displayName = subkey.GetValue("DisplayName").ToString();
                            string displayVersion = (subkey.GetValue("DisplayVersion") != null) ? subkey.GetValue("DisplayVersion").ToString() : "";
                            string uninstallString = (subkey.GetValue("UninstallString") != null) ? subkey.GetValue("UninstallString").ToString() : "";
                            string installLocation = (subkey.GetValue("InstallLocation") != null) ? subkey.GetValue("InstallLocation").ToString() : "";

                            return new CurApp(displayName, displayVersion, installLocation, uninstallString);
                        }

                    }
                }
            }

            return null;

        }

        int Cmd(string line) {

            Process newpc = Process.Start(new ProcessStartInfo { FileName = "cmd", Arguments = $"/c {line}", WindowStyle = ProcessWindowStyle.Hidden });
            newpc.WaitForExit();
            int PcessExitCod = newpc.ExitCode;
            Console.WriteLine(PcessExitCod.ToString());
            return PcessExitCod;
        }

    }

     
}
