using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Install_From_MSI
{
    [Serializable]
    class NewApp {

        #region using msi.dll
        [DllImport("msi.dll", SetLastError = true)]
        static extern uint MsiOpenDatabase(string szDatabasePath, IntPtr phPersist, out IntPtr phDatabase);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiDatabaseOpenViewW(IntPtr hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out IntPtr phView);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiViewExecute(IntPtr hView, IntPtr hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern uint MsiViewFetch(IntPtr hView, out IntPtr hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiRecordGetString(IntPtr hRecord, int iField,
           [Out] StringBuilder szValueBuf, ref int pcchValueBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern IntPtr MsiCreateRecord(uint cParams);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern uint MsiCloseHandle(IntPtr hAny);

        #endregion


        public string Name { get; private set; }
        private string PathX64 { get; set; }
        private string PathX86 { get; set; }
        public string PathMsi { get; private set; }
        public string Version { get; private set; }
        public bool Update { get; private set; }
        public bool ForceRestart { get; private set; }
        public string Property { get; set; }

        private Random rnd;
        public string pathToTempFileLog{ get; private set; }

        public event EventHandler<InstallEvent> EventLog;

        public NewApp(string pathX64, string pathX86,string version, bool update, bool forceRestart, string property) {
            PathX64 = pathX64;
            PathX86 = pathX86;
            Version = version;
            PathMsi = GetPathMSI(PathX64, PathX86);
            Update = update;
            ForceRestart = forceRestart;
            Property = property;
        }

        public NewApp(Options options) {


            PathMsi = GetPathMSI(options.PathX64, options.PathX86);
            Name = GetMSIInfo(PathMsi, "ProductName");
            Version = (!string.IsNullOrEmpty(options.Version)) ? options.Version: GetMSIInfo(PathMsi, "ProductVersion");;
            Update = options.Update;
            ForceRestart = options.ForceRestart;
            Property = options.Property;
            rnd = new Random();
            pathToTempFileLog = Path.Combine(Environment.GetEnvironmentVariable("TEMP", EnvironmentVariableTarget.Machine), $"LogInstallApp{rnd.Next(1000, 2000)}.log");
            
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

        string GetMSIInfo(string fileName, string Property) {

            try {
                using (Stream stream = new FileStream(fileName, FileMode.Open)) {
                }
            }
            catch (Exception e) {
                //check here why it failed and ask user to retry if the file is in use.
                Console.WriteLine(e.Message);
            }
            string sqlStatement = "SELECT * FROM Property WHERE Property = '" + Property + "'";
            IntPtr phDatabase = IntPtr.Zero;
            IntPtr phView = IntPtr.Zero;
            IntPtr hRecord = IntPtr.Zero;

            StringBuilder szValueBuf = new StringBuilder();
            int pcchValueBuf = 255;

            // Open the MSI database in the input file 
            uint val = MsiOpenDatabase(fileName, IntPtr.Zero, out phDatabase);

            hRecord = MsiCreateRecord(1);

            // Open a view on the Property table for the version property 
            int viewVal = MsiDatabaseOpenViewW(phDatabase, sqlStatement, out phView);

            // Execute the view query 
            int exeVal = MsiViewExecute(phView, hRecord);

            // Get the record from the view 
            uint fetchVal = MsiViewFetch(phView, out hRecord);

            // Get the version from the data 
            int retVal = MsiRecordGetString(hRecord, 2, szValueBuf, ref pcchValueBuf);

            uint uRetCode;
            uRetCode = MsiCloseHandle(phDatabase);
            uRetCode = MsiCloseHandle(phView);
            uRetCode = MsiCloseHandle(hRecord);

            return szValueBuf.ToString();
        }

        public bool IsNeedInstall(CurApp curapp) {
            string[] arrVersionCurApp = curapp.DisplayVersion.Trim().Split('.');
            string[] arrversionMSI = Version.Trim().Split('.');
            int count = (arrVersionCurApp.Count() >= arrversionMSI.Count()) ? arrversionMSI.Count() : arrVersionCurApp.Count();

            for (int i = 0; i <= count; i++) {
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

            
            if (curapp != null)
                if (!IsNeedInstall(curapp)) {
                    // Если версия установленного приложения больше или равно версии msi - то просто выходим из приложения.
                    Console.WriteLine("программа не требует обновления  - завершаем работу");
                    Environment.Exit(0);
                } else {
                    // если Update = false - то удаляем текущую установленную программу.
                    if (!Update) {
                        UninstallApp(curapp);
                        if (ForceRestart)
                            Environment.Exit(0);
                    }
                }

           


            string installString = $"msiexec /i {PathMsi} /quiet";
            if (ForceRestart)
                installString = installString + " /forcerestart";
            installString = installString + $" /log {pathToTempFileLog}";

            int numinst = Cmd(installString);
            switch (numinst) { 
                case 0:
                EventLog?.Invoke(this, new InstallEvent("Event: программа установлена", pathToTempFileLog, ResultJob.Install));
                    break;
                case 1641:
                    EventLog?.Invoke(this, new InstallEvent("Event: программа установлена и требуется перезагрузка", pathToTempFileLog, ResultJob.Install));
                    Environment.Exit(0);
                    break;
                default:
                    EventLog?.Invoke(this, new InstallEvent($"Event: ошибка № {numinst} при установки программы", pathToTempFileLog, ResultJob.ErrorInstall));
                    Environment.Exit(0);
                    break;
            }

            if (ForceRestart)
                Environment.Exit(0);

        }

                        


         void UninstallApp(CurApp curapp) {
            string uninstallString = curapp.UninstallString;

            #region Удаляем Msi пакеты используя msiexec
            if (!string.IsNullOrEmpty(uninstallString) && uninstallString.ToLower().Contains("msiexec")) {
                Console.WriteLine($"Будем удалять {curapp.DisplayName} версии {curapp.DisplayVersion} используя msiexec");
                if (!uninstallString.ToLower().Contains("quiet"))
                    uninstallString = uninstallString + " /quiet";
                if (ForceRestart & !uninstallString.ToLower().Contains("forcerestart"))
                    uninstallString = uninstallString + " /forcerestart";
                uninstallString = uninstallString + $" /log {pathToTempFileLog}";
                int numuninst = Cmd(uninstallString);

                switch (numuninst) {
                    case 0:
                        EventLog?.Invoke(this, new InstallEvent("Event: программа удалена", pathToTempFileLog, ResultJob.Uninstall));
                        break;
                    case 1641:
                        EventLog?.Invoke(this, new InstallEvent("Event: программа установлена и требуется перезагрузка", pathToTempFileLog, ResultJob.Uninstall));
                        Environment.Exit(0);
                        break;
                    default:
                        EventLog?.Invoke(this, new InstallEvent($"Event: ошибка № {numuninst} при удалении программы", pathToTempFileLog, ResultJob.ErrorUninstall));
                        Environment.Exit(0);
                        break;
                }                

            }
            #endregion

        }


        CurApp ListProg() {
            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

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
