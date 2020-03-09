using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class DataOutput
    {
        public string MessageResultPDF { get; set; }
        public string MessageResultXML { get; set; }
        public string MessageResultError { get; set; }
        public string MessageError { get; set; }
        public string MessageLogTime { get; set; }
        public string Message_Content { get; set; }
        public bool StatusCallAPI { get; set; }
        
    }
}
