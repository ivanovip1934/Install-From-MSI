using System;
using System.IO;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;

namespace Install_From_MSI {
    internal class GetMsiInfo {
        #region using msi.dll

        [DllImport("msi.dll", SetLastError = true)]
        private static extern uint MsiOpenDatabase(string szDatabasePath, IntPtr szPersist, out IntPtr phDatabase);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        private static extern int MsiDatabaseOpenViewW(IntPtr hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out IntPtr phView);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        private static extern int MsiViewExecute(IntPtr hView, IntPtr hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        private static extern uint MsiViewFetch(IntPtr hView, out IntPtr hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        private static extern int MsiRecordGetString(IntPtr hRecord, int iField, [Out] StringBuilder szValueBuf, ref int pcchValueBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        private static extern IntPtr MsiCreateRecord(uint cParams);

        [DllImport("msi.dll", ExactSpelling = true)]
        private static extern uint MsiCloseHandle(IntPtr hAny);

        [DllImport("msi.dll", CharSet = CharSet.Auto)]
        public static extern uint MsiDatabaseCommit(IntPtr phDatabas);

        #endregion

        public string ValueProperty { get; private set; }


        //public void GetMSIInfo(string fileName, string property) {
        public void GetMSIInfo(string fileName, string property)
        {
            try
            {
                using (Stream stream = new FileStream(fileName, FileMode.Open))
                {
                }
            }
            catch (Exception e)
            {
                //check here why it failed and ask user to retry if the file is in use.
                Console.WriteLine(e.Message);
            }

            var sqlStatement = "SELECT * FROM Property WHERE Property = '" + property + "'";
            var phDatabase = IntPtr.Zero;
            var phView = IntPtr.Zero;
            var hRecord = IntPtr.Zero; 
            var szPersist = IntPtr.Zero;

            var pcchValueBuf = 255;
            var szValueBuf = new StringBuilder(pcchValueBuf);
            

            // Open the MSI database in the input file 
            var val = MsiOpenDatabase(fileName, szPersist, out phDatabase);

            hRecord = MsiCreateRecord(1);

            // Open a view on the Property table for the version property 
            var viewVal = MsiDatabaseOpenViewW(phDatabase, sqlStatement, out phView);

            // Execute the view query 
            var exeVal = MsiViewExecute(phView, hRecord);

            // Get the record from the view 
            var fetchVal = MsiViewFetch(phView, out hRecord);

            // Get the version from the data 
            var retVal = MsiRecordGetString(hRecord, 2, szValueBuf, ref pcchValueBuf);

            uint uRetCode;

            uRetCode = MsiDatabaseCommit(phDatabase);
            uRetCode = MsiCloseHandle(phView);
            uRetCode = MsiCloseHandle(hRecord);
            uRetCode = MsiCloseHandle(phDatabase);

            ValueProperty = szValueBuf.ToString();
        }
    }
}