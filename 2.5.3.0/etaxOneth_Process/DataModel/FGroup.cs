using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOneth_Process.DataModel
{
    public class FGroup
    {
        public string Data_Type { get; set; }
        public string Line_Total_Count { get; set; }
        public string Delivery_Occur { get; set; }
        public string Invoice_Curr_Code { get; set; }
        public string Tax_Type_Code { get; set; }
        public string Tax_Cal_Rate { get; set; }
        public string Basis_Amount { get; set; }
        public string Basis_Curr_Code { get; set; }
        public string Tax_Cal_Amount { get; set; }
        public string Tax_Cal_Curr_Code { get; set; }
        public string Al_Charge_IND { get; set; }
        public string Al_Actual_Amount { get; set; }
        public string Al_Actual_Curr_Code { get; set; }
        public string Al_Reason_Code { get; set; }
        public string Al_Reason { get; set; }
        public string Payment_Type_Code { get; set; }
        public string Payment_Discription { get; set; }
        public string Payment_Due_Dtm { get; set; }
        public string Original_Total_Amount { get; set; }
        public string Original_Total_Curr_Code { get; set; }
        public string LINE_TOTAL_AMOUNT { get; set; }
        public string LINE_TOTAL_CURRENCY_CODE { get; set; }
        public string Adjusted_Inform_Amount { get; set; }
        public string Adjusted_Inform_Curr_Code { get; set; }
        public string Al_Total_Amount { get; set; }
        public string Al_Total_Curr_Code { get; set; }
        public string Charge_Total_Amount { get; set; }
        public string Charge_Total_Curr_Code { get; set; }
        public string Tax_Basis_Amount { get; set; }
        public string Tax_Basis_Curr_Code { get; set; }
        public string Tax_Total_Amount { get; set; }
        public string Tax_Total_Curr_Code { get; set; }
        public string Grand_Total_Amount { get; set; }
        public string Grand_Total_Curr_Code { get; set; }
    }
}
