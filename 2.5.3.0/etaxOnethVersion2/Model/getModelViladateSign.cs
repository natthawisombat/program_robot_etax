using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOnethVersion2.Model
{
    public class getModelViladateSign
    {
        public string SellerTaxId { set; get; }
        public string SellerBranchId { set; get; }
        public string UserCode { set; get; }
        public string AccessKey { set; get; }
        public string APIKey { set; get; }
        public string ServiceCode { set; get; }
        public string TextContent { set; get; }
    }
    public class getModelOutPutViladateSign
    {
        public string MessageResultError { get; set; }
        public bool StatusCallAPI { get; set; }
        public string TransactionCode { set; get; }
        public string Status { set; get; }
        public string ErrorCode { set; get; }
        public string ErrorMessage { set; get; }
        public string ResponseMessage { set; get; }
        public string MessageResultStatus { get; set; }
    }
}
