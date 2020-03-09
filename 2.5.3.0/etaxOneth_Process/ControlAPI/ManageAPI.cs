using etaxOneth_Process.DataModel;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace etaxOneth_Process.ControlAPI
{
    public class ManageAPI
    {
        public DataOutput CallAPI(DtGetParameters dtInput, string strPathTxtContent, string strPathPDFContent)
        {
            DataOutput strMessageExecute = new DataOutput();
            JObject oKeepResponeExecute = new JObject();
            JObject oKeepResponeExecute_docstatus = new JObject();
            string etaxgetdocumentstatus_url = "";
            if (dtInput.ServiceURL == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
            {
                etaxgetdocumentstatus_url = "https://uatetaxsp.one.th/etaxdocumentws/etaxgetdocumentstatus";
            }
            else
            {
                etaxgetdocumentstatus_url = "https://etaxsp.one.th/etaxdocumentws/etaxgetdocumentstatus";
            }
            try
            {
                var client = new RestClient(dtInput.ServiceURL);
                var request = new RestRequest(Method.POST);
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                request.AddHeader("Cache-Control", "no-cache");
                request.AddParameter("SellerTaxId", dtInput.SellerTaxID);
                request.AddParameter("SellerBranchId", dtInput.BranchID);
                request.AddParameter("APIKey", dtInput.APIKey);
                request.AddParameter("UserCode", dtInput.UserCode);
                request.AddParameter("AccessKey", dtInput.AccessKey);
                request.AddParameter("ServiceCode", dtInput.ServiceCode);
                request.AddFile("TextContent", strPathTxtContent);
                if (!strPathPDFContent.Equals(""))
                {
                    request.AddFile("PDFContent", strPathPDFContent);
                }
                request.Timeout = 31000;

                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                IRestResponse response = client.Execute(request);
                HttpStatusCode statusCode = response.StatusCode;
                int numericStatusCode = (int)statusCode;
                Console.WriteLine(numericStatusCode + " check");
                stopWatch.Stop();
                oKeepResponeExecute = JObject.Parse(response.Content);
                if (numericStatusCode == 200)
                {
                    if(oKeepResponeExecute["status"].ToString() == "PC")
                    {
                        pc_status_doc:
                        Thread.Sleep(1000);
                        var client_docstatus = new RestClient(etaxgetdocumentstatus_url);
                        var request_docstatus = new RestRequest(Method.POST);
                        request_docstatus.AddHeader("Cache-Control", "no-cache");
                        request_docstatus.AddParameter("SellerTaxId", dtInput.SellerTaxID);
                        request_docstatus.AddParameter("SellerBranchId", dtInput.BranchID);
                        request_docstatus.AddParameter("APIKey", dtInput.APIKey);
                        request_docstatus.AddParameter("UserCode", dtInput.UserCode);
                        request_docstatus.AddParameter("AccessKey", dtInput.AccessKey);
                        request_docstatus.AddParameter("TransactionCode", oKeepResponeExecute["transactionCode"].ToString());
                        IRestResponse response_docstatus = client.Execute(request_docstatus);
                        HttpStatusCode statusCode_docstatus = response_docstatus.StatusCode;
                        oKeepResponeExecute_docstatus = JObject.Parse(response_docstatus.Content);
                        if(oKeepResponeExecute_docstatus["status"].ToString() == "PC")
                        {
                            goto pc_status_doc;
                        }
                        else
                        {
                            strMessageExecute.MessageLogTime = stopWatch.ElapsedMilliseconds.ToString() + " ms";
                            strMessageExecute.MessageResultPDF = oKeepResponeExecute_docstatus["pdfURL"].ToString();
                            strMessageExecute.MessageResultXML = oKeepResponeExecute_docstatus["xmlURL"].ToString();
                            strMessageExecute.Message_Content = oKeepResponeExecute_docstatus.ToString();
                            strMessageExecute.StatusCallAPI = true;
                        }
                        
                    }
                    else
                    {
                        strMessageExecute.MessageLogTime = stopWatch.ElapsedMilliseconds.ToString() + " ms";
                        strMessageExecute.MessageResultPDF = oKeepResponeExecute["pdfURL"].ToString();
                        strMessageExecute.MessageResultXML = oKeepResponeExecute["xmlURL"].ToString();
                        strMessageExecute.Message_Content = oKeepResponeExecute.ToString();
                        strMessageExecute.StatusCallAPI = true;
                    }
                    
                }
                else if(numericStatusCode == 500)
                {
                    strMessageExecute.StatusCallAPI = false;
                    strMessageExecute.MessageResultError = "Internal Server Error มีข้อผิดพลาดบางอย่างภายใน ไม่ทราบสาเหตุ";
                }
                else
                {
                    strMessageExecute.StatusCallAPI = false;
                    strMessageExecute.MessageResultError = "Service Etax was problem";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                strMessageExecute.MessageResultError = oKeepResponeExecute.Root.ToString();
                strMessageExecute.StatusCallAPI = false;

                //return strMessageExecute;
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
