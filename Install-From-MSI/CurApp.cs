using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Install_From_MSI
{
    class CurApp {
        public string DisplayName { get; set; }
        public string DisplayVersion { get; set; }
        public string InstallLocation { get; set; }
        public string UninstallString { get; set; }
        private string tmpnamelog;
        private string tmplog;
        private Random rnd;
       

        public CurApp(string displayName, string displayVersion, string installLocation, string uninstallString) {
            DisplayName = displayName;
            DisplayVersion = displayVersion;
            InstallLocation = installLocation;
            UninstallString = uninstallString;
            rnd = new Random();
            tmpnamelog = "curapp" + rnd.Next(1000, 2000).ToString() + ".log";
            tmplog = Path.Combine(Environment.GetEnvironmentVariable("Temp"),tmpnamelog);

        }





    }
}
