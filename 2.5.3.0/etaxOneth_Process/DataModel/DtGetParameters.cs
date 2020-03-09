using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class DtGetParameters
    {
        public string PathInput { set; get; }
        public string PathOutput { set; get; }
        public string SellerTaxID { set; get; }
        public string BranchID { set; get; }
        public string APIKey { set; get; }
        public string UserCode { set; get; }
        public string AccessKey { set; get; }
        public string ServiceCode { set; get; }
        public string AmountFile { set; get; }
        public string ServiceURL { set; get; }
        public string PathConfigExcel { set; get; }
    }
}
