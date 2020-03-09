using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class HGroup
    {
        public string Data_Type { get; set; }
        public string Doc_Type_Code { get; set; }
        public string Doc_Name { get; set; }
        public string Doc_ID { get; set; }
        public string Doc_Issue_Dtm { get; set; }
        public string Create_Purpose_Code { get; set; }
        public string Create_Purpose { get; set; }
        public string Add_Ref_Assign_ID { get; set; }
        public string Add_Ref_Issue_Dtm { get; set; }
        public string Add_Ref_Type_Code { get; set; }
        public string Delivery_Type_Code { get; set; }
        public string Buyer_Order_Assign_ID { get; set; }
        public string Buyer_Order_Issue_Dtm { get; set; }
        public string Buyer_Order_Ref_Type_Code { get; set; }
        public string DOCUMENT_REMARK { get; set; }

    }
}
