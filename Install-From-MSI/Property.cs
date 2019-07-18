using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Install_From_MSI
{
    [Serializable]
    public class Params
    {
        public string PathX64 { get; set; }
        public string PathX86 { get; set; }
        public bool Update { get; set; }
        public bool ForceRestart { get; set; }
        public string Property { get; set; }

        private Params() { }

        public Params(string pathX64, string pathX86, bool update, bool forceRestart, string property) {
            PathX64 = pathX64;
            PathX86 = pathX86;
            Update = update;
            ForceRestart = forceRestart;
            Property = property;
        }

    }
}
