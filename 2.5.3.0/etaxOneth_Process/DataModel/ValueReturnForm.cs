using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class ValueReturnForm
    {
        public bool StatusRunning { set; get; }
        public int CountFileRun { set; get; }
        public int AmountAllFile { set; get; }
        public bool StatusFindPDF { set; get; }
        public string pathPrint { set; get; }
    }
}
