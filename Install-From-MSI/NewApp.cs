using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;

namespace Install_From_MSI
{
    [Serializable]
    class NewApp
    {

        public string Name { get; private set; }
        public string PathMsi { get; private set; }
        public bool x86AppOnX64Os { get; private set; }
        public string Version { get; private set; }
        public bool Update { get; private set; }
        public bool ForceRestart { get; private set; }
        public string Property { get; set; }

        private Random rnd;
        private Dictionary<string,string> ptrnnames;
        public string PathToTempFileLog { get; private set; }

        public event EventHandler<InstallEvent> EventLog;

        public NewApp(string pathX64, string pathX86, string version, bool update, bool forceRestart, string property)
        {
            Version = version;
            PathMsi = GetPathMSI(pathX64, pathX86);
            Update = update;
            ForceRestart = forceRestart;
            Property = property;
        }

        public NewApp(Options options, Dictionary<string,string> ptrnNames)
        {

            x86AppOnX64Os = options.x86AppOnX64Os;
            PathMsi = GetPathMSI(options.PathX64, options.PathX86);
            Version = (!string.IsNullOrEmpty(options.Version)) ? options.Version : GetInfoMSI(PathMsi, "ProductVersion");
            Name = GetInfoMSI(PathMsi, "ProductName");
            //Version = (!string.IsNullOrEmpty(options.Version)) ? options.Version: GetMSIInfo(PathMsi, "ProductVersion");
            Update = options.Update;
            ForceRestart = options.ForceRestart;
            Property = options.Property;
            rnd = new Random(StartVal());
            ptrnnames = ptrnNames;
            PathToTempFileLog = Path.Combine(Environment.GetEnvironmentVariable("TEMP", EnvironmentVariableTarget.Machine), $"LogInstallApp{rnd.Next(1000, 2000)}.log");

        }


        string GetPathMSI(string pathX64, string pathX86)
        {
            bool IsOSX64 = Environment.Is64BitOperatingSystem;
            switch (IsOSX64)
            {
                case true:
                    if (this.x86AppOnX64Os)
                    {
                        if (!String.IsNullOrEmpty(pathX86))
                        {
                            if (File.Exists(pathX86))
                                return pathX86;
                            else
                            {
                                Console.WriteLine($"MSI: {pathX86} NOT exists");
                                Environment.Exit(0);
                            }
                        }
                        else
                        {
                            Console.WriteLine("Ошибка: Путь до MSI пакета для x86 OS не задан");
                            Console.WriteLine("Ошибка: Установка не возможна.");
                            Environment.Exit(0);
                        }
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(pathX64))
                        {
                            if (File.Exists(pathX64))
                                return pathX64;
                            else
                            {
                                Console.WriteLine($"MSI: {pathX64} NOT exists");
                                Environment.Exit(0);
                            }

                        }
                        else
                        {
                            Console.WriteLine("Ошибка: Путь до MSI пакета для x64 OS не задан");
                            Console.WriteLine("Ошибка: Установка не возможна.");
                            Environment.Exit(0);
                        }
                    }
                    break;
                case false:
                    if (!String.IsNullOrEmpty(pathX86))
                    {
                        if (File.Exists(pathX86))
                            return pathX86;
                        else
                        {
                            Console.WriteLine($"MSI: {pathX86} NOT exists");
                            Environment.Exit(0);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Ошибка: Путь до MSI пакета для x86 OS не задан");
                        Console.WriteLine("Ошибка: Установка невозможна.");
                        Environment.Exit(0);
                    }
                    break;

                default:
                    break;
            }
            return String.Empty;
        }

        string GetInfoMSI(string fileName, string property)
        {
            GetMsiInfo msiinfo = new GetMsiInfo();
            msiinfo.GetMSIInfo(fileName, property);
            return msiinfo.ValueProperty;
        }

        public bool IsNeedInstall(CurApp curapp)
        {
            string[] arrVersionCurApp = curapp.DisplayVersion.Trim().Split('.');
            string[] arrversionMSI = Version.Trim().Split('.');
            int count = Math.Min(arrVersionCurApp.Count(), arrversionMSI.Count());
            int valueCurApp = 0;
            int valueVerMsi = 0;

            for (int i = 0; i < count; i++) {
                valueCurApp = int.Parse(arrVersionCurApp[i]);
                valueVerMsi = int.Parse(arrversionMSI[i]);
                if (valueVerMsi > valueCurApp)
                    return true;
                else if (valueVerMsi < valueCurApp)
                    return false;
            }            
            return false;
        }


        public void InstallApp()
        {
            // 1. Проверяем: есть ли установленная версия данной программы?
            CurApp curapp = ListProg();
            if (curapp != null)
            {
                Console.WriteLine($"Cur App version {curapp.DisplayVersion}");
                Console.WriteLine($"Cur App UninstallString {curapp.UninstallString}");
                if (!IsNeedInstall(curapp))
                {
                    // Если версия установленного приложения больше или равно версии msi - то просто выходим из приложения.
                    Console.WriteLine("программа не требует обновления - завершаем работу");
                    Environment.Exit(0);
                }
                else
                {
                    // если Update = false - то удаляем текущую установленную программу.
                    if (!Update)
                    {
                        UninstallApp(curapp);
                        if (ForceRestart)
                            Environment.Exit(0);
                    }
                }
            }

            string installString = $"msiexec /i \"{PathMsi}\" /quiet";
            if (!String.IsNullOrEmpty(Property))
                installString += $" {Property}"; 
            if (ForceRestart)
                installString += " /forcerestart";
            installString += $" /log {PathToTempFileLog}";

            int numinst = Cmd(installString);
            switch (numinst)
            {
                case 0:
                    EventLog?.Invoke(this, new InstallEvent("Event: программа установлена", ResultJob.Install));
                    break;
                case 1641:
                    EventLog?.Invoke(this, new InstallEvent("Event: программа установлена и требуется перезагрузка", ResultJob.Install));
                    Environment.Exit(0);
                    break;
                default:
                    EventLog?.Invoke(this, new InstallEvent($"Event: ошибка № {numinst} при установки программы", ResultJob.ErrorInstall));
                    Environment.Exit(0);
                    break;
            }
        }


        void UninstallApp(CurApp curapp)
        {
            string uninstallString = curapp.UninstallString.ToLower();

            #region Удаляем Msi пакеты используя msiexec
            if (!string.IsNullOrEmpty(uninstallString) && uninstallString.Contains("msiexec"))
            {
                Console.WriteLine($"Будем удалять {curapp.DisplayName} версии {curapp.DisplayVersion} используя msiexec");
                if (uninstallString.Contains("/i"))
                {
                    Console.WriteLine("Vmesto udaleniy zapustinza ustanovka");
                    uninstallString = uninstallString.ToLower().Replace("/i", "/x");
                }

                if (!uninstallString.Contains("quiet"))
                    uninstallString += " /quiet";

                if (ForceRestart & !uninstallString.Contains("forcerestart"))
                    uninstallString += " /forcerestart";
                uninstallString += $" /log {PathToTempFileLog}";
                int numuninst = Cmd(uninstallString);

                switch (numuninst)
                {
                    case 0:
                        EventLog?.Invoke(this, new InstallEvent("Event: программа удалена", ResultJob.Uninstall));
                        break;
                    case 1641:
                        EventLog?.Invoke(this, new InstallEvent("Event: программа установлена и требуется перезагрузка", ResultJob.Uninstall));
                        Environment.Exit(0);
                        break;
                    default:
                        EventLog?.Invoke(this, new InstallEvent($"Event: ошибка № {numuninst} при удалении программы", ResultJob.ErrorUninstall));
                        Environment.Exit(0);
                        break;
                }

            }
            #endregion

        }


        CurApp ListProg()
        {
            string registry_key = String.Empty;
            if (Environment.Is64BitOperatingSystem)
                registry_key = this.x86AppOnX64Os ?
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" :
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            else
                registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            
            if (string.IsNullOrEmpty(registry_key) || string.IsNullOrEmpty(this.Name))
            {
#if DEBUG
                Console.WriteLine("reg path or appName is null");
#endif
                Environment.Exit(0);
            }

            string prntName = GetPattern(this.Name);
            Regex rgx = new Regex(prntName, RegexOptions.IgnoreCase);

            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
            {
                var appfromreg = from subkey in key.GetSubKeyNames()
                                 let displayName = key.OpenSubKey(subkey).GetValue("DisplayName")?.ToString()
                                 where displayName != null
                                 where rgx.IsMatch(displayName)
                                 select new CurApp
                                 (
                                     displayName,
                                     key.OpenSubKey(subkey).GetValue("DisplayVersion")?.ToString(),
                                     key.OpenSubKey(subkey).GetValue("InstallLocation")?.ToString(),
                                     key.OpenSubKey(subkey).GetValue("UninstallString")?.ToString()                                     
                                 );

                if (appfromreg.Count() == 1)
                { 
#if DEBUG
                
                    Console.WriteLine($"DisplayName: {appfromreg.First().DisplayName}\n" +
                                              $"DisplayVersion: {appfromreg.First().DisplayVersion}\n" +
                                              $"UninstallString: {appfromreg.First().UninstallString}\n" +
                                              $"InstallLocation: {appfromreg.First().InstallLocation}");
#endif
                    return appfromreg.First();
                }
                return null;
            }
        }


        string GetPattern(string name)
        {
            if (ptrnnames.ContainsKey(name))
                return ptrnnames[name];

            foreach (KeyValuePair<string,string> ptrn in this.ptrnnames)
            {
                if (IsName(name, ptrn))
                    return ptrn.Value;
            }
            return name;
        }

        static bool IsName(string appname, KeyValuePair<string, string> patter)
        {
            Regex rgx = new Regex(patter.Value, RegexOptions.IgnoreCase);
            return rgx.IsMatch(appname);
        }

        int Cmd(string line)
        {
            int startvalue = StartVal();
            Thread.Sleep(startvalue);
            Random rnd = new Random(startvalue);

            do
            {
                Thread.Sleep(rnd.Next(1500, 10000));
            } while (InstallerServiceIsUsing()); 
            
            Process newpc = Process.Start(new ProcessStartInfo {
                                                                FileName = "cmd",
                                                                Arguments = $"/c {line}",
                                                                WindowStyle = ProcessWindowStyle.Hidden });
            newpc.WaitForExit();
            int PrcExitCod = newpc.ExitCode;
            Console.WriteLine(PrcExitCod.ToString());
            return PrcExitCod;
        }

        public static bool InstallerServiceIsUsing()
        {

            try
            {
                using (var mutex = Mutex.OpenExisting(@"Global\_MSIExecute"))
                {
                    return true;
                }
            }
            catch (Exception)
            {
                // Mutex not found; MSI isn't running
            }
            return false;

        }
        private  int StartVal()
        {
            int startvalue = 0;
            RandomNumberGenerator rng = new RNGCryptoServiceProvider();
            byte[] tokenData = new byte[1];
            rng.GetBytes(tokenData);
            string Varstr = String.Empty;
            for (int i = 0; i <= tokenData.Length - 1; i++)
            {
                Varstr += tokenData[i].ToString();
            }
            startvalue = int.Parse(Varstr);
            return startvalue;
            
        }
    }
}
