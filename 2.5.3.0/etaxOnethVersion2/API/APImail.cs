using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using etaxOnethVersion2.DATA_API;
using System.Threading;

namespace etaxOnethVersion2.API
{
    class APImail
    {
        public string email { get; set; }
        public string taxseller { get; set; }
        public string branch { get; set; }
        public string path { get; set; }
        public etaxOnethVersion2.etaxOneth form { get; set; }
        public string input { get; set; }
        public string timeuser { get; set; }
        public string typesoft { get; set; }
        public JObject jsontext { get; set; }
        public JObject jsontext_about { get; set; }
        public string numberstringgen { get; set; }
        public string err_code { get; set; }
        public string err_msg { get; set; }
        public string actionmsg { get; set; }
        public string textstring__ { get; set; }
        API_MAIL n_text_apimail = new API_MAIL();
        etaxOneth formetax;
        public string iphost = "https://etaxgateway.one.th/apiprog";
        //public string iphost = "http://localhost:8210";
        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
        bool pingeng = true;
        string textcheckupdate;
        string textversion;
        

        public static bool PingHost(string nameOrAddress)
        {
            bool pingable = false;
            Ping pinger = null;

            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }

            return pingable;
        }
        public void _checkversion_program()
        {
            Thread t_check = new Thread(new ThreadStart(_checkversion_program_thread));
            try
            {
                t_check.Start();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }
        public void _checkversion_program_thread()
        {
            
            //etaxOneth form;
            //Console.WriteLine(n_text_apimail);
            string textversion = "";
            //pingeng = PingHost("devinet-etax.one.th");
            JObject oKeepResponeExecute = new JObject();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            try
            {
                if (pingeng == true)
                {
                    string version = fvi.FileVersion;
                    var client = new RestClient(iphost+"/api/version");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("undefined", "{\n    \"version_this\": \"" + version + "\"\n}", ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                    Console.WriteLine(oKeepResponeExecute["dataversion"]);
                }

            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
                
            }
            try
            {
                
                if (pingeng == true && oKeepResponeExecute["type"].ToString() == "OK-Con")
                {
                    textversion = oKeepResponeExecute["dataversion"].ToString();
                    textstring__ = "true\\" + oKeepResponeExecute["dataversion"].ToString() + "\\" + oKeepResponeExecute["noc_phone"].ToString();
                }
                else if (pingeng == true && oKeepResponeExecute["type"].ToString() == "Fail-Con")
                {
                    textversion = oKeepResponeExecute["dataversion"].ToString();
                    textstring__ =  "false\\" + oKeepResponeExecute["dataversion"].ToString() + "\\" + oKeepResponeExecute["noc_phone"].ToString();
                }
                else if (pingeng == false)
                {
                    textstring__ =  "return\\no\\no";
                }
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            try
            {
                if (textstring__.Split('\\')[0] == "true")
                {
                    form.label25.Text = "โปรแกรมเป็นเวอร์ชั่นล่าสุด\n" + "เวอร์ชั่น : " + textstring__.Split('\\')[1];
                    form.button5.Visible = false;
                    form.label25.Update();
                }
                else if (textstring__.Split('\\')[0] == "false")
                {
                    form.txtversioncurrent = textstring__.Split('\\')[1];
                    form.label25.Text = "กรุณาอัพเดทโปรแกรม";
                    form.button5.Visible = true;
                    form.label25.Update();
                }
                else
                {
                    form.label25.Text = "ไม่สามารถเชื่อมต่อ\nระบบอัพเดทโปรแกรมได้";
                    form.label25.Update();
                    //button5.Visible = true;
                }
            }
            catch (Exception ea)
            {

            }

            //return "0";

        }
        public void start_prog()
        {
            Thread t_start = new Thread(new ThreadStart(start_prog_t));
            t_start.Start();
        }
        public void start_prog_t()
        {
            string email = this.email;
            string taxseller = this.taxseller;
            string branch = this.branch;
            string path = this.path;
            string input = this.input;
            string timeuser = this.timeuser;
            string typesoft = this.typesoft;
            string idcom;
            etaxOneth form = this.form;
            JObject oKeepResponeExecute = new JObject();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            string type_console = "start_program";
            idcom = form.numberstringgen;
            Console.WriteLine(form.inputnameFile.Text + "-" + idcom);
            try
            {
                //pingeng = PingHost("devinet-etax.one.th");
                
                if (pingeng)
                {
                    string str = "{\n\t\"mailuser\":\"" + email + "\",\n\t\"tax_id\":\"" + taxseller + "\",\n\t\"branch_id\":\"" + branch + "\",\n\t\"setpath\":\"" + path + "\",\n\t\"inputpath\":\"" + input + "\",\n\t\"ipaddres\":\"" + GetLocalIPAddress() + "\",\n" +
                        "\t\"namehost\":\"" + Environment.MachineName + "\",\n\t\"typesoft\":\"" + typesoft + "\",\n\t\"timeuser\":\"" + timeuser + "\",\n\t\"versionpro\":\"" + version + "\",\n\t\"idcomputer\":\"" + form.inputnameFile.Text + "-" + idcom + "\",\n\t\"type_console\":\"" + type_console + "\"\n}";
                    JObject json = JObject.Parse(str);
                    Console.WriteLine(json + " ==== start");
                    var client = new RestClient(iphost + "/api/v1/forprogram");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("undefined", str, ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                    Console.WriteLine(oKeepResponeExecute);
                }
                
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
                
        }
        public void first_export()
        {
            Thread t1 = new Thread(new ThreadStart(strat_export));
            t1.Start();
        }
        public void next_export()
        {
            Thread t2 = new Thread(new ThreadStart(coun_export));
            t2.Start();
        }
        public void strat_export()
        {
            string email = this.email;
            string taxseller = this.taxseller;
            string branch = this.branch;
            string path = this.path;
            string input = this.input;
            string timeuser = this.timeuser;
            string typesoft = this.typesoft;
            etaxOneth form = this.form;
            JObject oKeepResponeExecute = new JObject();
            JObject oKeepResponeExecute_ver = new JObject();
            string idcom;
            string type_console = "starttimeauto_program";
            try
            {                
                //pingeng = PingHost("devinet-etax.one.th");
                idcom = form.numberstringgen;
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                Console.WriteLine(pingeng);
                if (pingeng && form.emailtxt.Text.Length > 0)
                {
                    var client = new RestClient(iphost + "/api/v1/forprogram");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Content-Type", "application/json");
                    string str = "{\n    \"mailuser\": \"" + email + "\",\n    \"tax_id\": \"" + taxseller + "\",\n    \"branch_id\": \"" + branch + "\",\n    \"setpath\": \"" + path + "\",\n    \"room\": \"" + path + "\",\n    \"inputpath\": \"" + input + "\",\n    \"ipaddres\": \"" + GetLocalIPAddress() + "\",\n    " +
                        "\"namehost\": \"" + Environment.MachineName + "\",\n    \"typesoft\": \"" + typesoft + "\",\n    \"timeuser\": " + timeuser + ",\n    \"versionpro\": \"" + version + "\",\n    \"idcomputer\": \"" + form.inputnameFile.Text +"-" + idcom + "\",\n    \"type_console\": \"" + type_console + "\"\n}";
                    Console.WriteLine(str + " ==========");
                    JObject json = JObject.Parse(str);
                    request.AddParameter("undefined", str, ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    Console.WriteLine(response.Content.ToString());
                    oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                    form.label20.Text = oKeepResponeExecute.ToString();
                    //n_text_apimail.jsontext = oKeepResponeExecute;
                    //n_text_apimail.jsontext_about = json;
                    if (form.label18.Text.Length == 0)
                    {
                        form.label18.Text = json.ToString();
                    }
                    form.label25.Text = "กำลังตรวจสอบ เวอร์ชั่น";
                    var client_ver = new RestClient(iphost + "/api/v1/nxt_version");
                    var request_ver = new RestRequest(Method.POST);
                    request_ver.AddHeader("cache-control", "no-cache");
                    request_ver.AddHeader("Content-Type", "application/json");
                    request_ver.AddParameter("undefined", "{\n    \"this_version\": \"" + version + "\"\n}", ParameterType.RequestBody);
                    IRestResponse response_ver = client_ver.Execute(request_ver);
                    oKeepResponeExecute_ver = JObject.Parse(response_ver.Content.ToString());
                    Console.WriteLine(oKeepResponeExecute_ver);
                    textcheckupdate = oKeepResponeExecute_ver["msg"].ToString();
                    textversion = oKeepResponeExecute_ver["version"].ToString();
                    if (textcheckupdate == "ok")
                    {
                        form.label25.Text = "โปรแกรมเป็นเวอร์ชั่นล่าสุด\n" + "เวอร์ชั่น : " + textversion;
                        form.button5.Visible = false;
                    }
                    else
                    {
                        form.label25.Text = "กรุณาอัพเดทโปรแกรม";
                        form.button5.Visible = true;
                    }
                }
                else
                {
                    this.form.label25.Text = "ไม่สามารถเชื่อมต่อ\nกับระบบแจ้งเตือนได้\nเนื่องจากไม่พบ Email";
                }
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public void coun_export()
        {
            etaxOneth form = this.form;
            JObject oKeepResponeExecute = new JObject();
            JObject oKeepResponeExecute_ver = new JObject();
            string type_console = "next_api";
            try
            {
                //pingeng = PingHost("devinet-etax.one.th");
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                Console.WriteLine(pingeng);
                string str = form.label18.Text;
                string _sid = form.label20.Text;
                JObject json = JObject.Parse(str);
                JObject json_sid = JObject.Parse(_sid);
                Console.WriteLine(json);
                if (pingeng && json["mailuser"].ToString().Length > 0)
                {
                    var client = new RestClient(iphost + "/api/v1/forprogram");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Content-Type", "application/json");        
                    json["type_console"] = type_console;
                    json["sid"] = json_sid["sid"];
                    Console.WriteLine(json.ToString());
                    //string str = "{\n    \"mailuser\": \"" + email + "\",\n    \"tax_id\": \"" + taxseller + "\",\n    \"branch_id\": \"" + branch + "\",\n    \"setpath\": \"" + path + "\",\n    \"room\": \"" + path + "\",\n    \"inputpath\": \"" + input + "\",\n    \"ipaddres\": \"" + GetLocalIPAddress() + "\",\n    " +
                    //    "\"namehost\": \"" + Environment.MachineName + "\",\n    \"typesoft\": \"" + typesoft + "\",\n    \"timeuser\": " + timeuser + ",\n    \"versionpro\": \"" + version + "\",\n    \"idcomputer\": \"" + form.inputnameFile.Text + "-" + idcom + "\",\n    \"type_console\": \"" + type_console + "\"\n}";
                    //JObject json = JObject.Parse(str);
                    request.AddParameter("undefined", json.ToString(), ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                    Console.WriteLine(oKeepResponeExecute);
                    if(oKeepResponeExecute["version"].ToString() == version)
                    {
                        form.label25.Text = "โปรแกรมเป็นเวอร์ชั่นล่าสุด\n" + "เวอร์ชั่น : " + textversion;
                        form.button5.Visible = false;
                    }
                    else
                    {
                        form.label25.Text = "กรุณาอัพเดทโปรแกรม";
                        form.button5.Visible = true;
                    }
                }
                else
                {
                    this.form.label25.Text = "ไม่สามารถเชื่อมต่อ\nกับระบบแจ้งเตือนได้\nเนื่องจากไม่พบ Email";
                }
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public void stop_program_auto()
        {
            Thread stop_auto = new Thread(new ThreadStart(stop_thread_auto));
            stop_auto.Start();
        }
        public void stop_thread_auto()
        {
            //etaxOneth form = n_text_apimail.form;
            
            JObject oKeepResponeExecute = new JObject();
            try
            {
                //pingeng = PingHost("devinet-etax.one.th");
                if (pingeng)
                {
                    string _idcomputer = this.form.label20.Text;
                    JObject json__idcomputer = JObject.Parse(_idcomputer);
                    Console.WriteLine(json__idcomputer + " ==========");
                    var client = new RestClient(iphost + "/api/v1/forprogram");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("undefined", "{\n\t\"cid_computer\":\"" + json__idcomputer["data"].ToString() + "\",\n\t\"sid\":\"" + json__idcomputer["sid"].ToString() + "\",\n\t\"type_console\":\"end_auto\"\n}", ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                    Console.WriteLine(oKeepResponeExecute);
                }                    
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public void send_err_service()
        {
            Thread sendmail_service = new Thread(new ThreadStart(send_err_service_thread));
            sendmail_service.Start();
        }
        public void send_err_service_thread()
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                string str = "{'hostname': '" + Environment.MachineName + "','ipaddress': '" + GetLocalIPAddress() + "','inputpath': '" + this.input.Split('\\')[this.input.Split('\\').Length - 2] + "','outputpath': '" + this.path.Split('\\')[this.path.Split('\\').Length - 2] + "','err_code': '" + this.err_code + "','email': '" + this.email + "','taxid': '" + this.taxseller + "','err_msg': '" + this.err_msg + "','actionmsg':'" + this.actionmsg + "'}";
                Console.WriteLine(str + " TEST to JSON");
                //string _idcomputer = this.form.label20.Text;
                //JObject json__idcomputer = JObject.Parse(_idcomputer);
                var client = new RestClient(iphost + "/api/v1/mail_err");
                var request = new RestRequest(Method.POST);
                JObject json = JObject.Parse(str);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Content-Type", "application/json");

                request.AddParameter("undefined", json.ToString(), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                Console.WriteLine(oKeepResponeExecute + " ERR Service");
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public void close_program__()
        {
            Thread close_program__service = new Thread(new ThreadStart(close_program__thread));
            close_program__service.Start();
        }
        public void close_program__thread()
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                string _idcomputer = this.form.label20.Text;
                JObject json__idcomputer = JObject.Parse(_idcomputer);
                var client = new RestClient(iphost + "/api/v1/forprogram");
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("Content-Type", "application/json");
                //request.AddParameter("undefined", "{\n\t\"cid_computer\":\"" + json__idcomputer["data"].ToString() + "\",\n\t\"sid\":\"" + json__idcomputer["sid"].ToString() + "\",\n\t\"type_console\":\"exitprogram\"\n}", ParameterType.RequestBody);
                request.AddParameter("undefined", "{\n\t\"cid_computer\":\"" + json__idcomputer["data"].ToString() + "\",\n\t\"sid\":\"" + json__idcomputer["sid"].ToString() + "\",\n\t\"tax_id\":\"" + taxseller +"\",\n\t\"email\":\"" + email +"\",\n\t\"datainput_path\":\"" + input +"\",\n\t\"data_path\":\"" + path +"\",\n\t\"gethostname\":\"" + Environment.MachineName + "\",\n\t\"type_console\":\"exitprogram\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                oKeepResponeExecute = JObject.Parse(response.Content.ToString());
                Console.WriteLine(oKeepResponeExecute + " CLOSE");
            }
            catch (Exception ea)
            {
                Console.WriteLine(ea);
            }
        }
        public static string GetLocalIPAddress()
        {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            throw new Exception("No network adapters with an IPv4 address in the system!");
        }

    }
}
