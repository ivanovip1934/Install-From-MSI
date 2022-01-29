using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Serialization;


namespace Install_From_MSI
{
    class Program
    {

#if DEBUG
        static void Main(string[] args)
        {
#else
            #region  DLL import for Hidden Window

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("Kernel32")]
        private static extern IntPtr GetConsoleWindow();

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

            #endregion

            static void Main(string[] args)
        {
            #region Anable Hidden windows
            
            IntPtr hwnd;
            hwnd = GetConsoleWindow();
            ShowWindow(hwnd, SW_HIDE);
            
            #endregion
#endif
            Options Parm2 = new Options();
            string path = String.Empty;
            GeneralConfig Conf = GeneralConfig.GetInstance();

            #region Check args
            if (args.Count() != 1)
            {
#if DEBUG
                Console.WriteLine("Program is close");
#endif
                Environment.Exit(1);
            }
            #endregion Check args

#if DEBUG
            Console.WriteLine($"Programs is running...{args.Count()}");
#endif
            path = Path.Combine(Conf.PathToDirXML, $"{args[0]}.xml");

            if (!File.Exists(path))
            {
#if DEBUG
                Console.WriteLine($"Config file: {args[0]}.xml not exists");
#endif
                Environment.Exit(1);
            }

            try
            {
                using (var sr = new StreamReader(path))
                {
                    XmlSerializer writer1 = new System.Xml.Serialization.XmlSerializer(typeof(Options));
                    Parm2 = (Options)writer1.Deserialize(sr);
                }
            }
            catch (Exception e)
            {
#if DEBUG
                Console.WriteLine($"Not able reade config file {args[0]}.xml");
                Console.WriteLine($"Error:\n{e.Message}");
#endif
                Environment.Exit(1);
            }

            if (Parm2 != null)
            {
                NewApp newapp = new NewApp(Parm2, Conf.PtrnNames);
                WriteEvents wrevent = new WriteEvents(Conf.PathToDirLog);

                newapp.EventLog += wrevent.WriteLog;
#if DEBUG
                Console.WriteLine("Path to MSI: " + newapp.PathMsi);
                Console.WriteLine("x86AppOnX64Os is: " + newapp.x86AppOnX64Os);
                Console.WriteLine("Name app in MSI: " + newapp.Name);
                Console.WriteLine("Version MSI app: " + newapp.Version);
#endif
                newapp.InstallApp();
            }
        }
    }
}
