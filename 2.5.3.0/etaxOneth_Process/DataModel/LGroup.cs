using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class LGroup
    {
        public string Data_Type { get; set; }
        public string Line_ID { get; set; }
        public string Product_ID { get; set; }
        public string Product_Name { get; set; }
        public string Product_Desc { get; set; }
        public string Product_Batch_ID { get; set; }
        public string Product_Expire_Dtm { get; set; }
        public string Product_Class_Code { get; set; }
        public string Product_Class_Name { get; set; }
        public string Product_OriCountry_ID { get; set; }
        public string Product_Charge_Amount { get; set; }
        public string Product_Charge_Curr_Code { get; set; }
        public string Product_Al_Charge_IND { get; set; }
        public string Product_Al_Actual_Amount { get; set; }
        public string Product_Al_Actual_Curr_Code { get; set; }
        public string Product_Al_Reason_Code { get; set; }
        public string Product_Al_Reason { get; set; }
        public string Product_Quantity { get; set; }
        public string Product_Unit_Code { get; set; }
        public string Product_Quan_Per_Unit { get; set; }
        public string Line_Tax_Type_Code { get; set; }
        public string Line_Tax_Cal_Rate { get; set; }
        public string Line_Basis_Amount { get; set; }
        public string Line_Basis_Curr_Code { get; set; }
        public string Line_Tax_Cal_Amount { get; set; }
        public string Line_Tax_Cal_Curr_Code { get; set; }
        public string Line_AL_Charge_IND { get; set; }
        public string Line_AL_Actual_Amount { get; set; }
        public string Line_AL_Actual_Curr_Code { get; set; }
        public string Line_AL_Reason_Code { get; set; }
        public string Line_AL_Reason { get; set; }
        public string Line_Tax_Total_Amount { get; set; }
        public string Line_Tax_Total_Curr_Code { get; set; }
        public string Line_Net_Total_Amount { get; set; }
        public string Line_Net_Total_Curr_Code { get; set; }
        public string Line_Net_Include_Amount { get; set; }
        public string Line_Net_Include_Curr_Code { get; set; }
        public string Product_Remark { get; set; }
    }
}
