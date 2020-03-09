using Quobject.SocketIoClientDotNet.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using etaxOnethVersion2;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using MsgBox;
using System.Threading;
using System.Net;
using System.Diagnostics;
using RestSharp;
using System.Net.NetworkInformation;
//using SocketIO.Client;

namespace etaxOnethVersion2.API
{
    class GetSocket : IDisposable
    {

        //SocketIO.Client.Socket socketio;
        public string ip = "http://etaxgateway.one.th:8000/";
        private static string emailpublic, pathpublic;        
        string sid;
        string datastring;
        string[] sidsum;
        string pathIn;
        private bool disposed;
        string textcheckupdate;
        public string textversion;
        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
        string idcom;
        //var socket = IO.Socket(ip);
        Socket socket = IO.Socket("https://etaxgateway.one.th:8100/");
        /*https://etaxgateway.one.th:8100/*/
        bool pingeng = false;
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
        public string API_checkversion()
        {
            pingeng = PingHost("etaxgateway.one.th");
            JObject oKeepResponeExecute = new JObject();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            try
            {
                if(pingeng == true)
                {
                    string version = fvi.FileVersion;
                    var client = new RestClient("https://etaxgateway.one.th:8100/api/version");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("undefined", "{\n    \"version_this\": \"" + version + "\"\n}", ParameterType.RequestBody);
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
            if (pingeng == true && oKeepResponeExecute["type"].ToString() == "OK-Con")
            {
                textversion = oKeepResponeExecute["dataversion"].ToString();
                return "true\\" + oKeepResponeExecute["dataversion"].ToString() + "\\" + oKeepResponeExecute["noc_phone"].ToString();
            }
            else if(pingeng == true && oKeepResponeExecute["type"].ToString() == "Fail-Con")
            {
                textversion = oKeepResponeExecute["dataversion"].ToString();
                return "false\\" + oKeepResponeExecute["dataversion"].ToString() + "\\" + oKeepResponeExecute["noc_phone"].ToString();
            }
            else if(pingeng == false)
            {
                return "return\\no\\no";
            }
            return "0";
            



        }
        public void connect_thisSocket(etaxOneth form)
        {
            socket.Connect();
        }
        public void connectSocket(etaxOneth form)
        {
            JObject oKeepResponeExecute = new JObject();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            if(form.label25.Text.ToString() == "กรุณาอัพเดทโปรแกรม")
            {
                try
                {
                    string version = fvi.FileVersion;
                    socket.Connect();
                    form.label25.Text = "กำลังตรวจสอบ เวอร์ชั่น";
                    form.button5.Visible = false;
                    //Console.WriteLine(socket.Connect());
                    socket.On(Socket.EVENT_CONNECT, () =>
                    {
                        Console.WriteLine("Connect");
                    });
                    socket.Emit("checkversioncurrent", version);
                    socket.On("version", (data) =>
                    {
                        Console.WriteLine(data);
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        textcheckupdate = oKeepResponeExecute["type"].ToString();
                        textversion = oKeepResponeExecute["dataversion"].ToString();
                        if (textcheckupdate == "OK-Con")
                        {
                            form.label25.Text = "โปรแกรมเป็นเวอร์ชั่นล่าสุด\n" + "เวอร์ชั่น : " + textversion;
                            form.button5.Visible = false;
                        }
                        else
                        {
                            form.label25.Text = "กรุณาอัพเดทโปรแกรม";
                            form.button5.Visible = true;
                        }
                    });
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
            else if (form.label25.Text.ToString() == "ปิดระบบแจ้งเตือน")
            {

                try
                {
                    string status_socket = this.API_checkversion();
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
                        Console.WriteLine(count);
                        GC.Collect();
                    }
                    GC.WaitForPendingFinalizers();
                    GC.SuppressFinalize(this);
                }
            }
            //else if (form.metroToggle1.Checked == true)
            //{
            //    string version = fvi.FileVersion;
            //    Console.WriteLine(version);
            //    form.label25.Text = "กำลังตรวจสอบ เวอร์ชั่น";
            //    form.button5.Visible = false;
            //    //Console.WriteLine(socket.Connect());
            //    socket.Emit("checkversioncurrent", version);
            //    socket.On("version", (data) =>
            //    {
            //        oKeepResponeExecute = JObject.Parse(data.ToString());
            //        Console.WriteLine(oKeepResponeExecute);
            //        textcheckupdate = oKeepResponeExecute["type"].ToString();
            //        textversion = oKeepResponeExecute["dataversion"].ToString();
            //        if (textcheckupdate == "OK-Con")
            //        {
            //            form.label25.Text = "โปรแกรมเป็นเวอร์ชั่นล่าสุด\n" + "เวอร์ชั่น : " + textversion;
            //            form.button5.Visible = false;
            //        }
            //        else
            //        {
            //            form.label25.Text = "กรุณาอัพเดทโปรแกรม";
            //            form.button5.Visible = true;
            //        }
            //    });
            //}     
            
        }
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// The virtual dispose method that allows
        /// classes inherithed from this one to dispose their resources.
        /// </summary>
        /// <param name = "disposing" ></ param >
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources here.
                }

                // Dispose unmanaged resources here.
            }

            disposed = true;
        }
        public void checkSocket()
        {
            //socket.Connect();
            socket.On(Socket.EVENT_CONNECT, (data) =>
            {

            });
        }


        public void startSocketAuto1(string email, string taxseller, string branch, string path, etaxOneth form, string input, string timeuser, string typesoft)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                idcom = form.numberstringgen;
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                string str = "{ 'mailuser': '" + email + "','tax_id':'" + taxseller + "','branch_id':'" + branch + "','setpath':'" + path + "','room':'" + path + "','inputpath' : '" + input + "','ipaddres' : '" + GetLocalIPAddress() + "','namehost' : '" + Environment.MachineName + "','timeuser': '" + timeuser + "','typesoft': '" + typesoft + "','versionpro':'" + version + "','idcomputer':'" + form.inputnameFile.Text + "-" + idcom + "' }";
                string strroom = "{ 'room': '" + form.inputnameFile.Text + "-" + idcom + "'}";
                JObject json = JObject.Parse(str);
                JObject jsonroon = JObject.Parse(strroom);
                if (form.label18.Text.Length == 0)
                {
                    form.label18.Text = json.ToString();
                }
                socket.Emit("startprogramauto", json);
                socket.Emit("enterroom", jsonroon);
                socket.On("msg", (data) =>
                {
                    oKeepResponeExecute = JObject.Parse(data.ToString());
                    //sid = oKeepResponeExecute["data"].ToString();
                });
                socket.On("msgupdate", (data) =>
                {
                    oKeepResponeExecute = JObject.Parse(data.ToString());
                    textcheckupdate = oKeepResponeExecute["msg"].ToString();
                    textversion = oKeepResponeExecute["version"].ToString();
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
                });
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
                    Console.WriteLine(count);
                    GC.Collect();
                }
                GC.SuppressFinalize(this);
                oKeepResponeExecute = null;
                oKeepResponeExecute.Remove();
            }
            

        }
        public void startSocketAuto2(string email, string taxseller, string branch, string path, string statusprogram, etaxOneth form, string input, string timeuser, string typesoft)
        {
            
            JObject oKeepResponeExecute = new JObject();
            try
            {
                idcom = form.numberstringgen;
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                string str = "{ 'mailuser': '" + email + "','tax_id':'" + taxseller + "','branch_id':'" + branch + "','setpath':'" + path + "','room':'" + path + "','inputpath' : '" + input + "','ipaddres' : '" + GetLocalIPAddress() + "','namehost' : '" + Environment.MachineName + "','timeuser': '" + timeuser + "','typesoft': '" + typesoft + "','versionpro':'" + version + "','idcomputer':'" + form.inputnameFile.Text + "-" + idcom + "' }";
                string strroom = "{ 'room': '" + form.inputnameFile.Text + "-" + idcom + "'}";
                if (statusprogram == "1")
                {
                    JObject json = JObject.Parse(str);
                    if (form.label18.Text.Length == 0)
                    {
                        form.label18.Text = json.ToString();
                    }
                    socket.Emit("starttimeauto", json);
                    socket.On("msg", (data) =>
                    {
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        //sid = oKeepResponeExecute["data"].ToString();
                    });
                    socket.On("msgupdate", (data) =>
                    {
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        textcheckupdate = oKeepResponeExecute["msg"].ToString();
                        textversion = oKeepResponeExecute["version"].ToString();
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
                    });
                }
                else if (statusprogram == "2")
                {
                    JObject json = JObject.Parse(str);
                    JObject jsonroon = JObject.Parse(strroom);
                    if (form.label18.Text.Length == 0)
                    {
                        form.label18.Text = json.ToString();
                    }
                    
                    socket.Emit("enterroom", jsonroon);
                    socket.Emit("starttimemmauto", json);
                    socket.On("msg", (data) =>
                    {
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        //idcom = oKeepResponeExecute["data"].ToString();
                    });
                    socket.On("msgupdate", (data) =>
                    {
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        textcheckupdate = oKeepResponeExecute["msg"].ToString();
                        textversion = oKeepResponeExecute["version"].ToString();
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
                    });
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
                    Console.WriteLine(count);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
                oKeepResponeExecute = null;
                oKeepResponeExecute.Remove();
            }         

        }
        public void SendMailAlert(string inputpath, string outputpath, string err_code, string email, string taxid, string err_msg, string actionmsg)
        {
            JObject jj = new JObject();
            string str = "{'hostname': '" + Environment.MachineName + "','ipaddress': '" + GetLocalIPAddress() + "','inputpath': '" + inputpath.Split('\\')[inputpath.Split('\\').Length - 2] + "','outputpath': '" + outputpath.Split('\\')[inputpath.Split('\\').Length - 2] + "','err_code': '" + err_code + "','email': '" + email + "','taxid': '" + taxid + "','err_msg': '" + err_msg + "','actionmsg':'" + actionmsg + "'}";
            
            try
            {                
                jj = JObject.Parse(str);
                //connectSocket();
                socket.Emit("warning-err", jj);
                Console.WriteLine(jj + " test");
                pathIn = "";
            }
            catch (Exception efwaf)
            {
                //MessageBox.Show(e.Message);
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    Console.WriteLine(count);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
                jj.Remove();
                jj = null;
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
        public void EndSocketAuto(string type, etaxOneth form)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                if (type == "Cancel")
                {
                    socket.Emit("endtimeauto", form.inputnameFile.Text + "-" + idcom);
                    socket.On("msg", (data) =>
                    {
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        //idcom = oKeepResponeExecute["data"].ToString();
                    });
                }
                else if (type == "CloseAndCancel")
                {
                    socket.Emit("endtimeauto", form.inputnameFile.Text + "-" + idcom);
                    socket.On("msg", (data) =>
                    {
                        oKeepResponeExecute = JObject.Parse(data.ToString());
                        //idcom = oKeepResponeExecute["data"].ToString();
                    });
                    socket.Emit("exitprogram", form.inputnameFile.Text + "-" + idcom);
                    socket.Disconnect();
                }
                else if (type == "Close")
                {
                    idcom = form.numberstringgen;
                    socket.Emit("exitprogram", form.inputnameFile.Text + "-" + idcom);
                }
                else if (type == "CloseApplication")
                {
                    socket.Emit("endprocressAll");
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
                    Console.WriteLine(count);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
                oKeepResponeExecute = null;
                oKeepResponeExecute.Remove();
            }
            

        }
        public void stopSocket()
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                socket.Emit("endtime");
                socket.On("txtmsg", (data) =>
                {
                    oKeepResponeExecute = JObject.Parse(data.ToString());
                });
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
                    Console.WriteLine(count);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
                oKeepResponeExecute = null;
                oKeepResponeExecute.Remove();
            }
                
        }
        public void closeSocket()
        {
            socket.Close();
        }
        public void nextSocket(etaxOneth form)
        {
            JObject oKeepResponeExecute = new JObject();
            try
            {
                if (form.label18.Text.Length > 0)
                {
                    socket.Emit("update", form.label18.Text);
                }
                
                socket.On("update", (data) =>
                {
                    oKeepResponeExecute = JObject.Parse(data.ToString());
                    //sid = oKeepResponeExecute["data"].ToString();
                });
                socket.On("msgupdate", (data) =>
                {
                    oKeepResponeExecute = JObject.Parse(data.ToString());
                    textcheckupdate = oKeepResponeExecute["msg"].ToString();
                    textversion = oKeepResponeExecute["version"].ToString();
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
                });
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
                    Console.WriteLine(count);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
                oKeepResponeExecute = null;
                oKeepResponeExecute.Remove();
            }
            

        }
        public void disconnectSocket()
        {
            socket.Disconnect();
        }
        
    }
    
}
