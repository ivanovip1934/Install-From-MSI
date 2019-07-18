using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Install_From_MSI
{
    class Program
    {
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


        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;


            //Params Par1 = new Params ( @"C:\Programs\GoogleChrome\gchromex64.msi", @"C:\Programs\GoogleChrome\gchromex86.msi", true, false,"");
            //Console.WriteLine($"PathX86 = {Par1.PathX86}");

            //XmlSerializer writer = new System.Xml.Serialization.XmlSerializer(typeof(Params));
            //var path = @"C:\Programs\GoogleChrome\config_chrome.xml";
            //TextWriter file = new StreamWriter(path);
            //writer.Serialize(file,Par1);
            //file.Close();
            var path = @"C:\Programs\GoogleChrome\config_chrome.xml";
            Params Parm2;

            using (var sr = new StreamReader(path))
            {
                 XmlSerializer writer1 = new System.Xml.Serialization.XmlSerializer(typeof(Params));
                Parm2 = (Params)writer1.Deserialize(sr);
                 Console.WriteLine($"PathX64 = {Parm2.PathX64}");
                Console.WriteLine($"Update = {Parm2.Update}");

            }
            string PathToMSI = GetPathMSI(Parm2.PathX64, Parm2.PathX86);
            Console.WriteLine("Path to MSI: " + PathToMSI );
             
            Console.WriteLine(GetMSIInfo(PathToMSI, "ProductVersion"));
            Console.WriteLine(GetMSIInfo(PathToMSI, "ProductName"));


            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
            {
                Console.WriteLine(key.GetType());
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        //  Console.WriteLine(subkey.GetValue("DisplayName"));
                    }
                }
            }


            Console.ReadKey();


        }

        static string GetPathMSI(string pathX64, string pathX86)
        {
            bool IsOSX64 = Environment.Is64BitOperatingSystem;
            switch (IsOSX64)
            {
                case true:
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
                        Console.WriteLine("Ошибка: Установка не возможна.");
                        Environment.Exit(0);
                    }
                    break;

                default:
                    break;
            }


            return "privet";
        }

        //public static string GetMSIProperty(string msiFile, string msiProperty)
        //{
        //    string retVal = string.Empty;

        //    Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
        //    Object installerObj = Activator.CreateInstance(classType);
        //    WindowsInstaller.Installer installer = installerObj as WindowsInstaller.Installer;

        //    Database database = installer.OpenDatabase("C:\\DataP\\sqlncli.msi", 0);

        //    string sql = String.Format("SELECT Value FROM Property WHERE Property=’{0}’", msiProperty);

        //    View view = database.OpenView(sql);

        //    Record record = view.Fetch();

        //    if (record != null)
        //    {
        //        retVal = record.get_StringData(1);
        //    }
        //    else
        //        retVal = "Property Not Found";

        //    return retVal;
        //}

        public static string GetMSIInfo(string fileName, string Property)
        {
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


        //Warning 
        public static DataRow DataRowInstalledApplication(this RegistryKey rgkey, string keyName)
        {
            RegistryKey key = rgkey.OpenSubKey(keyName, false);
            try
            {
                //Application Name is mandetory for a given key.
                if (key == null || key.RegToString("DisplayName", false) == null) return null;
                //Build a sanitised data row
                var rowBuilder = new DataTable().NewRow();
                rowBuilder["DisplayName"] = key.RegToString("DisplayName");
                rowBuilder["UninstallString"] = key.RegToString("UninstallString");
                rowBuilder["InstallLocation"] = key.RegToString("InstallLocation");
                rowBuilder["Publisher"] = key.RegToString("Publisher");
                rowBuilder["DisplayIcon"] = key.RegToString("DisplayIcon");

                return rowBuilder;
            }
            finally
            {
                if (key != null) key.Close();
            }

        }





    }
}
