using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class PathFilesIO
    {
        public string PathInput { get; set; }
        public string PathOutput { get; set; }
        public string PathTemp { get; set; }
        public string PathErr { get; set; }
        public string PathSource_F { get; set; }
        public string PathSource_S { get; set; }
        public string PathSuccess_O { get; set; }
        public string PathLogTime { get; set; }
        public string PathLogFileRun { get; set; }
        public string PathFileRun { get; set; }
        public string TypeRunning { get; set; }
        public string DateTimeFolderName { get; set; }
        public string PathConfigExcel { get; set; }
        public string TypePrinting { get; set; }
        public string Printer { get; set; }
        public string TypePrintPreview { get; set; }
        public string TypeDoc { get; set; }
        public string LogTimeProcess { get; set; }
        public string BCP_Folder { get; set; }
    }
}
