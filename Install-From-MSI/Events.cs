using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Install_From_MSI
{
    class InstallEvent:EventArgs
    {
        public string Message { get; private set; }
        public string TmpFileLog { get; private set; }
        public ResultJob RJ { get; private set; }

        public InstallEvent(string message, ResultJob rj) {
            Message = message;
            RJ = rj;
        }
    }


}
