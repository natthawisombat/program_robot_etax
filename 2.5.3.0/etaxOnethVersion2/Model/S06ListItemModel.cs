using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOnethVersion2.Model
{
    class S06ListItemModel
    {
        public string ListHeader_start { get; set; }
        public string ListHeader_end { get; set; }
        public string sellertaxid { get; set; }
        public string sellerbranchid { get; set; }
        public string document_name { get; set; }
        public string document_id { get; set; }
        public string document_issue_dtm { get; set; }
        public string document_remark { get; set; }
        public string BUYER_NAME { get; set; }
        public string BUYER_TAX_ID { get; set; }
        public string BUYER_BRANCH_ID { get; set; }
        public string BUYER_URIID { get; set; }
        public string BUYER_ADDRESS { get; set; }
        public string BUYER_Country_PostCode { get; set; }
    }
}
