using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using etaxOnethVersion2.Model;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace etaxOnethVersion2.API
{
    public class APIManage
    {
        public getModelOutPutViladateSign ViladateSignAPI(getModelViladateSign inputFides)
        {
            Console.WriteLine(inputFides);
            getModelOutPutViladateSign strMessageExecute = new getModelOutPutViladateSign();
            JObject oKeepResponeExecute = new JObject();
            string etaxgetdocumentviladateSign = "https://uatetaxsp.one.th/etaxdocumentws/etaxvalidatesigndocument";
            try
            {
                var client = new RestClient(etaxgetdocumentviladateSign);
                var request = new RestRequest(Method.POST);
                request.AddHeader("Cache-Control", "no-cache");
                request.AddParameter("SellerTaxId", inputFides.SellerTaxId);
                request.AddParameter("SellerBranchId", inputFides.SellerBranchId);
                request.AddParameter("UserCode", inputFides.UserCode);
                request.AddParameter("AccessKey", inputFides.AccessKey);
                request.AddParameter("APIKey", inputFides.APIKey);
                request.AddParameter("ServiceCode", inputFides.ServiceCode);
                request.AddFile("TextContent", inputFides.TextContent);
                IRestResponse response = client.Execute(request);
                HttpStatusCode statusCode = response.StatusCode;
                int numericStatusCode = (int)statusCode;
                oKeepResponeExecute = JObject.Parse(response.Content);
                if(numericStatusCode == 200)
                {
                    strMessageExecute.ResponseMessage = oKeepResponeExecute.ToString();
                    strMessageExecute.MessageResultStatus = oKeepResponeExecute["status"].ToString();
                    strMessageExecute.StatusCallAPI = true;
                }
                else if (numericStatusCode == 500)
                {
                    strMessageExecute.StatusCallAPI = false;
                    strMessageExecute.MessageResultError = "Internal Server Error มีข้อผิดพลาดบางอย่างภายใน ไม่ทราบสาเหตุ";
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                strMessageExecute.StatusCallAPI = false;
                strMessageExecute.ResponseMessage = oKeepResponeExecute.ToString();
                return strMessageExecute;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return strMessageExecute;
        }
    }
}
