using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using etaxOnethVersion2;
using RestSharp;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace etaxOnethVersion2.API
{
    class getchat
    {
        public string idpord;
        public string urlprod;
        public string urluat;
        public string uidchatprod;
        String uidchatprod1;
        public void mainrestfulgetid(etaxOneth from)
        {

            try
            {
                if (from.lbUrl.Text == "Production")
                {
                    this.production(from);
                }
                else if (from.lbUrl.Text == "ทดสอบระบบ")
                {
                    this.uatpoc(from);
                }
            }
            catch (NullReferenceException a)
            {

            }


        }

        public void production(etaxOneth from)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {

                var client = new RestClient("https://etaxgateway.one.th:8550/api/v1/getuser_id");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("content-type", "application/json");
                request.AddParameter("application/json", "{\n\t\"username\":\"" + from.txtUserCode.Text + "\",\n\t\"type\":\"PROD\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                if (oKeepResponeExecute["message"].ToString() == "success")
                {
                    idpord = (oKeepResponeExecute["text"]["id"]).ToString();
                    uidchatprod1 = this.checktofrdprod(idpord);
                    from.txt_uid = uidchatprod1;
                    from.txt_email = oKeepResponeExecute["text"]["thai_email"].ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        public string checktofrdprod(string idpord)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client = new RestClient("https://chat-manage.one.th:8997/api/v1/searchfriend");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Authorization", "Bearer A16185216830056b1946f138905230c3c633dbeec596d4e8d962971c40269af89a5b101b00a02411db4d741312cee67d5");
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("undefined", "{\n    \"bot_id\": \"B0df132b88f9a526691a0576bfdb24196\",\n    \"key_search\": \"" + idpord + "\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                if (oKeepResponeExecute["status"].ToString() == "success")
                {
                    uidchatprod = oKeepResponeExecute["friend"]["user_id"].ToString();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return uidchatprod;
        }
        public string checktofrd_UAT(string iduat)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client = new RestClient("https://uatchat-manage.one.th:8997/api/v1/searchfriend");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Authorization", "Bearer A114be672146a57b690973a5b600f446187e8f9b094e84ae5a955780e477ae54a5cd8f63d965a462a822a7cad1b970374");
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("undefined", "{\n    \"bot_id\": \"B0df132b88f9a526691a0576bfdb24196\",\n    \"key_search\": \"" + iduat + "\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                Console.WriteLine(oKeepResponeExecute);
                if (oKeepResponeExecute["status"].ToString() == "success")
                {
                    uidchatprod = oKeepResponeExecute["friend"]["user_id"].ToString();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return uidchatprod;
        }
        public void testc()
        {
            Console.WriteLine(" *--*-*-*-*-** " + uidchatprod1);
        }
        public void sendtochat(string urlpor, etaxOneth from)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client1 = new RestClient("https://chat-public.one.th:8034/api/v1/push_message");
                var request1 = new RestRequest(Method.POST);
                request1.AddHeader("cache-control", "no-cache");
                request1.AddHeader("Authorization", "Bearer A16185216830056b1946f138905230c3c633dbeec596d4e8d962971c40269af89a5b101b00a02411db4d741312cee67d5");
                request1.AddHeader("Content-Type", "application/json");
                //request1.AddParameter("undefined", "{\n    \"to\": \"" + uidchatprod  + "\",\n    \"bot_id\": \"B0df132b88f9a526691a0576bfdb24196\",\n    \"type\": \"template\",\n    \"elements\": [\n        {\n            \"image\": \"http://clustem.eu/wp-content/uploads/icona-download-pdf.png\",\n            \"title\": \"PDF File\",\n            \"detail\": \"ตรวจสอบเอกสาร PDF\",\n            \"choice\": [\n                {\n                    \"label\": \"ตรวจสอบ\",\n                    \"type\": \"file\",\n                    \"file\": \"" + urlpor + "\",\n                    \"payload\": {\n                        \"keyword\": \"savefile\",\n                        \"msgtxt\": \"pdffile\"\n                    }\n                }\n            ]\n        }\n    ]\n}", ParameterType.RequestBody);
                request1.AddParameter("undefined", "{\n    \"to\": \"" + from.txt_uid + "\",\n    \"bot_id\": \"B0df132b88f9a526691a0576bfdb24196\",\n\"type\":\"template\",\n\"elements\": [\n {\n\"image\": \"http://clustem.eu/wp-content/uploads/icona-download-pdf.png\",\n \"title\": \"PDF File\",\n\"detail\": \"ชื่อเอกสาร " + from.txt_namefile + "Tracking " + from.txt_trackingid + "\",\n \"choice\": [\n {\n                    \"label\": \"ดาวน์โหลด\",\n                    \"type\": \"file\",\n                    \"file\": \"" + urluat + "\",\n                    \"payload\": {\n                        \"keyword\": \"savefile\",\n                        \"msgtxt\": \"pdffile\",\n   \"namefile\": \"" + from.txt_namefile + "\",\n \"tracking_id\":\"" + from.txt_trackingid + "\"\n                 }\n                }\n            ]\n        }\n    ]\n}", ParameterType.RequestBody);
                IRestResponse response1 = client1.Execute(request1);
                oKeepResponeExecute = JObject.Parse(response1.Content);
                Console.WriteLine(oKeepResponeExecute);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
        public void sendtochat_UAT(string urluat, etaxOneth from)
        {

            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client1 = new RestClient("https://uatchat-public.one.th:8034/api/v1/push_message");
                var request1 = new RestRequest(Method.POST);
                request1.AddHeader("cache-control", "no-cache");
                request1.AddHeader("Authorization", "Bearer A114be672146a57b690973a5b600f446187e8f9b094e84ae5a955780e477ae54a5cd8f63d965a462a822a7cad1b970374");
                request1.AddHeader("Content-Type", "application/json");
                //request1.AddParameter("undefined", "{\n    \"to\": \"" + uidchatprod  + "\",\n    \"bot_id\": \"B0df132b88f9a526691a0576bfdb24196\",\n    \"type\": \"template\",\n    \"elements\": [\n        {\n            \"image\": \"http://clustem.eu/wp-content/uploads/icona-download-pdf.png\",\n            \"title\": \"PDF File\",\n            \"detail\": \"ตรวจสอบเอกสาร PDF\",\n            \"choice\": [\n                {\n                    \"label\": \"ตรวจสอบ\",\n                    \"type\": \"file\",\n                    \"file\": \"" + urlpor + "\",\n                    \"payload\": {\n                        \"keyword\": \"savefile\",\n                        \"msgtxt\": \"pdffile\"\n                    }\n                }\n            ]\n        }\n    ]\n}", ParameterType.RequestBody);
                request1.AddParameter("undefined", "{\n    \"to\": \"" + from.txt_uid + "\",\n    \"bot_id\": \"B0df132b88f9a526691a0576bfdb24196\",\n\"type\":\"template\",\n\"elements\": [\n {\n\"image\": \"http://clustem.eu/wp-content/uploads/icona-download-pdf.png\",\n \"title\": \"PDF File\",\n\"detail\": \"ชื่อเอกสาร : " + from.txt_namefile + " Tracking : " + from.txt_trackingid + "\",\n \"choice\": [\n {\n                    \"label\": \"ดาวน์โหลด\",\n                    \"type\": \"file\",\n                    \"file\": \"" + urluat + "\",\n                    \"payload\": {\n                        \"keyword\": \"savefile\",\n                        \"msgtxt\": \"pdffile\",\n   \"namefile\": \"" + from.txt_namefile + "\",\n \"tracking_id\":\"" + from.txt_trackingid + "\"\n                 }\n                }\n            ]\n        }\n    ]\n}", ParameterType.RequestBody);
                IRestResponse response1 = client1.Execute(request1);
                oKeepResponeExecute = JObject.Parse(response1.Content);
                Console.WriteLine(oKeepResponeExecute);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private void uatpoc(etaxOneth from)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client = new RestClient("https://etaxgateway.one.th:8550/api/v1/getuser_id");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("content-type", "application/json");
                request.AddParameter("application/json", "{\n\t\"username\":\"" + from.txtUserCode.Text + "\",\n\t\"type\":\"UAT\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                Console.WriteLine(oKeepResponeExecute);
                if (oKeepResponeExecute["message"].ToString() == "success")
                {
                    idpord = (oKeepResponeExecute["text"]["id"]).ToString();
                    uidchatprod1 = this.checktofrd_UAT(idpord);
                    from.txt_uid = uidchatprod1;
                    from.txt_email = oKeepResponeExecute["text"]["thai_email"].ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        public string uploadfile_prod(string namefile, etaxOneth from)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client = new RestClient("https://chat-public.one.th:8034/api/v1/upload_file_for_bot");
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Bearer A16185216830056b1946f138905230c3c633dbeec596d4e8d962971c40269af89a5b101b00a02411db4d741312cee67d5");
                request.AddParameter("bot_id", "B0df132b88f9a526691a0576bfdb24196");
                request.AddFile("file[]", namefile);
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                stopWatch.Stop();
                urlprod = ("https://chat-public.one.th:8034/get_botfile/bucket-botfile/B0df132b88f9a526691a0576bfdb24196/" + oKeepResponeExecute["data"][0]["message"]);
                from.txt_namefile = oKeepResponeExecute["data"][0]["file_name"].ToString();

                if (oKeepResponeExecute["status"].ToString() == "success")
                {
                    Console.WriteLine(oKeepResponeExecute["status"].ToString());
                    this.GENTRACKING(from);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return urlprod;


        }
        public string uploadfile_uat(string namefile, etaxOneth from)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client = new RestClient("https://uatchat-public.one.th:8034/api/v1/upload_file_for_bot");
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Bearer A114be672146a57b690973a5b600f446187e8f9b094e84ae5a955780e477ae54a5cd8f63d965a462a822a7cad1b970374");
                request.AddParameter("bot_id", "B0df132b88f9a526691a0576bfdb24196");
                request.AddFile("file[]", namefile);
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                stopWatch.Stop();
                urluat = ("https://uatchat-public.one.th:8034/get_botfile/bucket-botfile/B0df132b88f9a526691a0576bfdb24196/" + oKeepResponeExecute["data"][0]["message"]);
                from.txt_namefile = oKeepResponeExecute["data"][0]["file_name"].ToString();

                if (oKeepResponeExecute["status"].ToString() == "success")
                {
                    Console.WriteLine(oKeepResponeExecute["status"].ToString());
                    this.GENTRACKING(from);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return urluat;
        }
        public void GENTRACKING(etaxOneth from)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                var client = new RestClient("https://etaxgateway.one.th:8550/api/v1/tracking");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("undefined", "{\n\t\"auth\":\"oeb2019\",\n\t\"username\":\"" + from.txtUserCode.Text + "\",\n\t\"file_name\":\"" + from.txt_namefile + "\",\n\t\"email\":\"" + from.txt_email + "\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content);
                Console.WriteLine(oKeepResponeExecute);
                from.txt_trackingid = oKeepResponeExecute["tracking_id"].ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                oKeepResponeExecute = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

    }
}
