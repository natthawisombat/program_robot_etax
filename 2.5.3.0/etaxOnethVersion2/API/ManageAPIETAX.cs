using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using RestSharp;
using System.Text;
using System.Threading.Tasks;

namespace etaxOnethVersion2.API
{
    class ManageAPIETAX
    {
        JObject oKeepResponeExecute = new JObject();
        public string CallAPISENDMAIL(string taxid,string branch,string email,string path)
        {
            var client = new RestClient("http://203.151.50.53:8000/api/firsttime");
            var request = new RestRequest(Method.POST);
            request.AddHeader("Cache-Control", "no-cache");
            request.AddHeader("content-type", "application/json");
            request.RequestFormat = DataFormat.Json;
            request.AddBody(new { mailuser = path, tax_id = taxid, branch_id = branch,setpath = path });
            IRestResponse response = client.Execute(request);
            Console.Write(response);
            Console.ReadLine();
            oKeepResponeExecute = JObject.Parse(response.Content);
            return oKeepResponeExecute["id"].ToString();
        }
    }
}
