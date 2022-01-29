using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Install_From_MSI
{



    public class Options

    {
        public string PathX64 { get; set; }
        public string PathX86 { get; set; }
        public bool x86AppOnX64Os { get; set; }
        public string Version { get; set; }
        public bool Update { get; set; }
        public bool ForceRestart { get; set; }
        public string Property { get; set; }
        //public string PathToDirLog { get; set; }

        public Options() { }

    }
}
