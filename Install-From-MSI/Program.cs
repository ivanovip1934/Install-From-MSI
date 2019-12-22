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
using System.Management;
using System.Diagnostics;

namespace Install_From_MSI
{
    class Program
    {
       


        static void Main(string[] args)
        {
            //  Console.OutputEncoding = System.Text.Encoding.UTF8;

            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            string registry_keyx6432 = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            string path = @"C:\Programs\GoogleChrome\config_chrome.xml";
            Options Parm2;


            //Params Par1 = new Params ( @"C:\Programs\GoogleChrome\gchromex64.msi", @"C:\Programs\GoogleChrome\gchromex86.msi", true, false,"");
            //Console.WriteLine($"PathX86 = {Par1.PathX86}");

            //XmlSerializer writer = new System.Xml.Serialization.XmlSerializer(typeof(Params));
            //var path = @"C:\Programs\GoogleChrome\config_chrome.xml";
            //TextWriter file = new StreamWriter(path);
            //writer.Serialize(file,Par1);
            //file.Close();


            
            

            using (var sr = new StreamReader(path))
            {
                 XmlSerializer writer1 = new System.Xml.Serialization.XmlSerializer(typeof(Options));
                Parm2 = (Options)writer1.Deserialize(sr);
            }
            
            NewApp newapp = new NewApp(Parm2);
            WriteEvents wrevent = new WriteEvents(Parm2.PathToDirLog ,"TestNameApp");

            newapp.EventLog += wrevent.WriteLog;
            //newapp.Install += wrevent.Install;
            //newapp.ErrorInstall += wrevent.InstallError;
            //newapp.Uninstall += wrevent.Uninstall;
            //newapp.ErrorUninstall += wrevent.UninstallError;
            Console.WriteLine("Path to MSI: " + newapp.PathMsi);
            newapp.InstallApp();
                
            Console.ReadKey();          


        }
               

    }
}
