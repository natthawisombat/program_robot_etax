//using etaxOneth_Process;
using etaxOneth_Process.DataModel;
using etaxOnethVersion2.API;
using Newtonsoft.Json.Linq;
using Ookii.Dialogs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Linq;
using MsgBox;
using Quobject.SocketIoClientDotNet.Client;
using etaxOnethVersion2;
using System.Security.Cryptography;
using System.Text;
using System.Net.NetworkInformation;
using System.Timers;
using GemBox.Spreadsheet;
using System.Text.RegularExpressions;
using System.Net;
using System.Management;
using etaxOneth_Printer;
using etaxOnethVersion2.DATA_API;

namespace etaxOnethVersion2
{
    public partial class etaxOneth : Form
    {
        int TogMove;
        int MValX;
        int MValY;
        int statusProgram = 1; // 1 คือโปรแกรมเจอไฟล์ 0 คือโปรแกรมไม่เจอไฟล์
        //string strPathInOut = AppDomain.CurrentDomain.BaseDirectory + "InOutPath\\InOutPath.txt";
        string strDateFileName = string.Empty;
        string email, path;
        string statusWorker; // Do คือ โปรแกรมทำงานอยู่ Cancel คือ ยกเลิก
        public string ____url { get; set; }
        public string ___copise { get; set; }
        public string set__copise_file { get; set; }
        PathFilesIO pfIO = new PathFilesIO();
        ManageAPIETAX Send_start = new ManageAPIETAX();
        JObject oKeepResponeExecute = new JObject();
        API_MAIL n_text_apimail = new API_MAIL();
        //etaxOnethProcess toProcess = new etaxOnethProcess();
        //FormProcress procressing = new FormProcress();
        //ProcressETAX toProcess = new ProcressETAX();
        string APIKEY = "AK2-3UY8R84Q6GFXD8QXZMFYNPUJRFOKWOT614C5GAAC88OVAGLD8F1HWAPB9LW05QQSACKVSN0FBI1H8WPO0NWU29MWO9854BBMW5OJ6IZOBUJFZRHV0ZPD4CP02PXLT95YHD6QX3TH01CZX7TJ7X4NBLKZLGQP8NF3BKZIU6NSW463VL61A8LBBAKIJSQS7M1TL2S50E6AGWRE85ZSK6AIHYDYXY7C19LKLFRFWXW3FAOAB61O0E5TGPBNC3P4X72BU";
        public string textinfile;
        public string linkpath;
        public string nameFile;
        public string myTestKey;
        public ArrayList lisnameFile = new ArrayList();
        public ArrayList listfileWorker = new ArrayList();
        public int countform;
        int countform1;
        APImail _apimail = new APImail();
        public string txttool;
        public string statusFile;
        public int perc = 0;
        string texttest;
        string typeurl;
        string hash = "0105561072420_00000_Etax_One_th";
        private ProcressETAX procress = new ProcressETAX();
        private DateTime startTime = DateTime.Now;
        public string txtidSocketio;
        public int sumtxtTimeout;
        private BackgroundWorker bgworker = null;
        private Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();
        string RandomNameOfFolder;
        public string numberstringgen;
        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
        WebClient webClient;
        ProgressDlg progressDlg;
        Stopwatch sw = new Stopwatch();
        string TypeDoc = "";
        public bool pingeng = true;
        public bool net_status_ { get; set; }
        public string txtversioncurrent = "";
        public etaxOneth()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            bgworker = new BackgroundWorker();
            bgworker.WorkerReportsProgress = true;
            bgworker.WorkerSupportsCancellation = true;
            bgworker.DoWork += backgroundWorker1_DoWork;
            try
            {

                if (rdbAuto.Checked == true)
                {
                    txtInput.Enabled = false;
                    txtOutput.Enabled = false;
                    btnInput.Enabled = false;
                    btnOutput.Enabled = false;
                    txtAmountFile.Enabled = true;
                    txtTimeRun.Enabled = true;
                    txtInput.Clear();
                    txtOutput.Clear();
                }
                else
                {
                    txtInput.Enabled = true;
                    txtOutput.Enabled = true;
                    btnInput.Enabled = true;
                    btnOutput.Enabled = true;
                    txtAmountFile.Enabled = false;
                    txtTimeRun.Enabled = false;
                    txtAmountFile.Clear();
                    txtTimeRun.Clear();
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
                GC.SuppressFinalize(this);
            }




        }
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

        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private void etaxOneth_Load(object sender, EventArgs e)
        {
            NetConfig_ __NET = new NetConfig_();
            bool __NET_stuats = __NET.internet_status_();
            
            Console.WriteLine(__NET_stuats + " === NET");

            numberstringgen = RandomString(60) + System.Environment.MachineName;
            _apimail.numberstringgen = numberstringgen;
            _apimail.form = this;
            //btnExport.Enabled = false;            
            button5.Visible = false;
            countform1 = Application.OpenForms.OfType<etaxOneth>().Count();
            //this.Show();
            //try
            //{


            //}

            //catch (NullReferenceException ex)
            //{
            //    Console.WriteLine(ex);
            //}
            //finally
            //{
            //    for (int i = 0; i <= GC.MaxGeneration; i++)
            //    {
            //        int count = GC.CollectionCount(i);
            //        Console.WriteLine(count);
            //        GC.Collect();
            //    }
            //    //getsocket.Dispose();
            //    GC.WaitForPendingFinalizers();
            //    GC.SuppressFinalize(this);
            //}
            if (textinfile != null)
            {
                cboTypeDoc.SelectedIndex = 0;
                int len = (textinfile.Split(',')).Length;
                bool check = bool.Parse(textinfile.Split(',')[15]);
                bool checkAutoPrint = false;
                try
                {
                    checkAutoPrint = bool.Parse(textinfile.Split(',')[13]);
                }
                catch (Exception ex)
                {
                    checkAutoPrint = false;
                }
                if (check == true)
                {
                    rdbAuto.Checked = check;
                    inputnameFile.Enabled = false;
                    inputnameFile.SelectedItem = nameFile;
                    if (checkAutoPrint)
                    {
                        chkprintauto.Checked = checkAutoPrint;
                        cboPrinter.SelectedItem = textinfile.Split(',')[14];
                    }
                    this.export();

                }
            }
            else
            {
                Microsoft.Win32.RegistryKey rkey;
                rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
                myTestKey = (string)rkey.GetValue("PathConfigETAX");
                int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
                string[] textfile = Directory.GetFiles(myTestKey, "*.cfg").Select(Path.GetFileName).ToArray();
                for (int i = 0; i < fileCount; i++)
                {
                    inputnameFile.Items.Add(textfile[i]);
                }

                rkey.Close();



            }


            try
            {
                cboTypeDoc.SelectedIndex = 0;
                cboServiceCode.Items.Add("S03");
                cboServiceCode.Items.Add("S06");
                cboServiceCode.Items.Add("S06(Excel Only)");
                cboServiceCode.Items.Add("S03(Excel Only)");
                cboServiceCode.Items.Add("S06(Excel Only & List Item)");
                cboServiceCode.Items.Add("BCP Service");
                cboServiceURL.Items.Add("ทดสอบระบบ (UAT)");
                cboServiceURL.Items.Add("Production");
                //cboServiceURL.Items.Add("https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument");
                //cboServiceURL.Items.Add("https://etaxsp.one.th/etaxdocumentws/etaxsigndocument");
                //cboServiceURL.Items.Add("http://localhost:8960/api_create_void");
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

        }

        private void pbClose_Click(object sender, EventArgs e)
        {
            try
            {
                cancelProcress("CloseAndCancel");
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

        private void pbPowerOff_Click(object sender, EventArgs e)
        {
            try
            {
                cancelProcress("CloseAndCancel");
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
        public void cancelProcress(string status)
        {
            string[] arrPathIO = null;
            string pathfile = myTestKey + "\\" + inputnameFile.Text;
            try
            {
                StreamReader txtw = new StreamReader(pathfile);
                String line = txtw.ReadToEnd();
                byte[] data = Convert.FromBase64String(line);
                using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                {
                    byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                    using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                    {
                        ICryptoTransform transform = tripleDes.CreateDecryptor();
                        byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                        //Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                        line = UTF8Encoding.UTF8.GetString(results);
                    }
                }
                txtw.Close();
                arrPathIO = line.Split(',');
                if (status == "CloseAndCancel")
                {
                    ValueReturnForm valueReturn = new ValueReturnForm();
                    Console.WriteLine(bgworker.IsBusy);
                    if (bgworker.IsBusy)
                    {
                        bgworker.CancelAsync();
                        //procress = new ProcressETAX();                    
                        statusWorker = "cancel";
                        if (pingeng)
                        {
                            try
                            {
                                _apimail.stop_program_auto();
                                _apimail.close_program__();
                            }
                            catch (NullReferenceException ex)
                            {
                                Console.WriteLine(ex.Message);
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
                        
                        if (txtAmountFile.Text == "")
                        {
                            procress.StopWorking("0", this);
                        }
                        else
                        {
                            procress.StopWorking(txtAmountFile.Text, this);
                        }
                        this.Close();

                    }
                    else
                    {
                        Console.WriteLine("Close");
                        if (pingeng)
                        {
                            try
                            {
                                _apimail.close_program__();

                            }
                            catch (NullReferenceException ex)
                            {
                                Console.WriteLine(ex.Message);
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
                                this.Close();
                            }
                        }
                        else
                        {
                            this.Close();
                        }
                            

                    }

                }
                else if (status == "Cancel")
                {
                    bgworker.CancelAsync();
                    txtStatus.Clear();
                    txtStatus.Refresh();
                    //procress = new ProcressETAX();
                    statusWorker = "cancel";
                    if (pingeng)
                    {
                        try
                        {
                            _apimail.stop_program_auto();

                        }
                        catch (NullReferenceException ex)
                        {
                            Console.WriteLine(ex.Message);
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
                        
                    procress.StopWorking(txtAmountFile.Text, this);
                    //InputBox.SetLanguage(InputBox.Language.English);
                    //InputBox.ShowDialog("ยกเลิกการทำงาน!",
                    //"Warning",   //Text message (mandatory), Title (optional)
                    //InputBox.Icon.Question, //Set icon type (default info)
                    //InputBox.Buttons.Ok, //Set buttons (default ok)
                    //InputBox.Type.Nothing, //Set type (default nothing)
                    //null, //String field as ComboBox items (default null)
                    //true, //Set visible in taskbar (default false)
                    //new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                    ////MessageBox.Show("ยกเลิกการทำงาน!!!");
                    //this.Close();
                    resetClearValue();
                    ////frmProcess.Close();
                    //etaxOneth etaxform = new etaxOneth();
                    //etaxform.Show();

                }
                string[] arrAllFile = System.IO.Directory.GetFiles(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder);
                foreach (var item in arrAllFile)
                {
                    try
                    {
                        File.Move(item, arrPathIO[1] + "\\Fail" + "\\" + Path.GetFileName(item));
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
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
                try
                {
                    System.IO.Directory.Delete(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder);
                }
                catch (IOException e)
                {

                }
                finally
                {
                    for (int i = 0; i <= GC.MaxGeneration; i++)
                    {
                        int count = GC.CollectionCount(i);
                        Console.WriteLine(count);
                        GC.Collect();
                    }
                }

            }
            catch (DirectoryNotFoundException ewda)
            {
                _apimail.close_program__();
                this.Close();
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
        private void resetClearValue()
        {
            inputnameFile.SelectedIndex = -1;
            txtSellerTaxID.Clear();
            txtBranchID.Clear();
            //txtAPIKey.Clear();
            txtUserCode.Clear();
            txtAccessKey.Clear();
            cboServiceCode.SelectedIndex = -1;
            lbUrl.Text = "";
            cboServiceURL.SelectedIndex = -1;
            cboTypeDoc.SelectedIndex = -1;
            emailtxt.Clear();
            inputnameFile.Enabled = true;
            txtSellerTaxID.Enabled = true;
            txtBranchID.Enabled = true;
            txtUserCode.Enabled = true;
            //txtAPIKey.Enabled = true;
            txtAccessKey.Enabled = true;
            cboServiceCode.Enabled = true;
            cboServiceURL.Enabled = true;
            emailtxt.Enabled = true;
            rdbAuto.Enabled = true;
            rdbManual.Enabled = true;
            rdbManual.Checked = true;
            txtStatusRunning.BackColor = SystemColors.Control;
            btnExport.Enabled = true;
            txtStatusRunning.ResetText();
            chkprintauto.Checked = false;
            chkprintauto.Enabled = true;
            check___copies.Checked = false;
            check___copies.Enabled = true;
            chkPreview.Checked = false;
            chkPreview.Enabled = true;
            cboTypeDoc.Enabled = true;
            pictureBox2.Enabled = true;
            pictureBox4.Enabled = true;


        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (statusWorker == "Do" && rdbAuto.Checked == true)
            {
                DialogResult result = MessageBox.Show("Do you want to Cancel?", "Confirmation", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    cancelProcress("Cancel");
                    txtStatusRunning.BackColor = SystemColors.Control;
                    //metroToggle1.Enabled = true;
                    cboTypeDoc.SelectedIndex = 0;
                    chkPreview.Visible = false;
                }
                else if (result == DialogResult.No)
                {
                    //...
                }

            }
            else if (statusWorker == "Do" && rdbManual.Checked == true)
            {

            }

        }

        private void pbRestoreDown_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            pbRestoreDown.Visible = false;
            pbMaximize.Visible = true;
        }

        private void pbMaximize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            pbMaximize.Visible = false;
            pbRestoreDown.Visible = true;
        }

        private void pbMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnInput_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFileDialog = new OpenFileDialog();

            if (oFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string ext = Path.GetExtension(oFileDialog.FileName);
                //MessageBox.Show(ext);
                if (ext == ".txt" || ext == ".xlsx" || ext == ".xls" || ext == ".csv")
                {
                    txtInput.Text = oFileDialog.FileName;
                }
                else
                {
                    InputBox.SetLanguage(InputBox.Language.English);
                    InputBox.ShowDialog("กรุณาเลือกไฟล์ txt หรือ xlsx!",
                    "Warning",   //Text message (mandatory), Title (optional)
                    InputBox.Icon.Information, //Set icon type (default info)
                    InputBox.Buttons.Ok, //Set buttons (default ok)
                    InputBox.Type.Nothing, //Set type (default nothing)
                    null, //String field as ComboBox items (default null)
                    true, //Set visible in taskbar (default false)
                    new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                    //MessageBox.Show("กรุณาเลือกไฟล์ txt หรือ xlsx!!!");
                    txtInput.Text = "";
                }
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            if (txtInput.Text.Equals(""))
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("Please Select File Input!",
                "Warning",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Information, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                //MessageBox.Show("Please Select File Input!");
                return;
            }

            VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtOutput.Text = dlg.SelectedPath;

                //string strDataPath = Path.GetDirectoryName(txtInput.Text) + "," + txtOutput.Text;
                //File.WriteAllText(strPathInOut, String.Empty);
                //CreateTextFile(strPathInOut, strDataPath);
            }
        }

        private void rdbManual_CheckedChanged(object sender, EventArgs e)
        {
            label21.Visible = false;
            pictureBox5.Visible = false;
            label22.Visible = true;
            pictureBox7.Visible = true;
            cboPrinter.Enabled = true;
            //label24.Enabled = true;
            input_copies.Enabled = true;
            label19.Enabled = true;
            chkprintauto.Enabled = true;
            chkprintauto.Checked = false;
            check___copies.Enabled = true;
            check___copies.Checked = false;
            txtInput.Enabled = true;
            txtOutput.Enabled = true;
            txtSellerTaxID.Enabled = true;
            txtBranchID.Enabled = true;
            //txtAPIKey.Enabled = true;
            btnInput.Enabled = true;
            btnOutput.Enabled = true;
            txtInput.Visible = true;
            txtOutput.Visible = true;
            btnInput.Visible = true;
            btnOutput.Visible = true;
            label23.Visible = true;
            txtConfixExcel.Visible = true;
            btnConfixExcel.Visible = true;
            label1.Visible = true;
            label3.Visible = true;
            lPathFile.Visible = false;
            //txtPathFile.Visible = false;
            //btnPathFile.Visible = false;
            txtAmountFile.Enabled = false;
            txtTimeRun.Enabled = false;
            txtConfixExcel.Clear();
            txtAmountFile.Clear();
            input_copies.Clear();
            this.___copise = string.Empty;
            txtTimeRun.Clear();
            inputnameFile.Visible = false;
            inputnameFile.Text = "";
            label17.Visible = false;
            emailtxt.Visible = false;
            clearform();
        }

        private void rdbAuto_CheckedChanged(object sender, EventArgs e)
        {
            label22.Visible = false;
            pictureBox7.Visible = false;
            label21.Visible = true;
            pictureBox5.Visible = true;
            cboPrinter.Enabled = false;
            //label24.Enabled = false;
            input_copies.Enabled = true;
            label19.Enabled = true;
            label23.Visible = false;
            txtConfixExcel.Visible = false;
            btnConfixExcel.Visible = false;
            label17.Visible = true;
            emailtxt.Visible = true;
            txtInput.Visible = false;
            txtOutput.Visible = false;
            btnInput.Visible = false;
            btnOutput.Visible = false;
            label1.Visible = false;
            label3.Visible = false;
            lPathFile.Visible = true;
            //txtPathFile.Visible = true;
            //btnPathFile.Visible = true;
            txtAmountFile.Enabled = true;
            txtTimeRun.Enabled = true;
            txtInput.Clear();
            txtOutput.Clear();
            inputnameFile.Visible = true;
            inputnameFile.Items.Clear();
            inputnameFile.Enabled = true;
            chkprintauto.Enabled = false;
            check___copies.Enabled = false;
            cboServiceCode.SelectedIndex = -1;
            chkPreview.Checked = false;
            cboServiceURL.SelectedIndex = -1;
            Microsoft.Win32.RegistryKey rkey;
            rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
            myTestKey = (string)rkey.GetValue("PathConfigETAX");
            int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
            if (fileCount != 0)
            {
                string[] textfile = Directory.GetFiles(myTestKey, "*.cfg").Select(Path.GetFileName).ToArray();
                for (int i = 0; i < fileCount; i++)
                {
                    inputnameFile.Items.Add(textfile[i]);
                }



            }
            else
            {

            }

            clearform();
        }

        private void btnHome_Click(object sender, EventArgs e)
        {
            inputnameFile.Items.Clear();
            Microsoft.Win32.RegistryKey rkey;
            rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
            myTestKey = (string)rkey.GetValue("PathConfigETAX");
            int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
            if (fileCount != 0)
            {
                string[] textfile = Directory.GetFiles(myTestKey, "*.cfg").Select(Path.GetFileName).ToArray();
                for (int i = 0; i < fileCount; i++)
                {
                    inputnameFile.Items.Add(textfile[i]);
                }


            }


            txtAccessKey.Clear();
            //txtAPIKey.Clear();
            txtBranchID.Clear();
            txtInput.Clear();
            txtOutput.Clear();
            //txtServiceCode.Clear();
            txtSellerTaxID.Clear();
            txtUserCode.Clear();
            txtAmountFile.Clear();
            txtTimeRun.Clear();
            inputnameFile.SelectedText = "";
            //txtServiceURL.Clear();
            inputnameFile.Enabled = true;
            btnInput.Enabled = true;
            btnOutput.Enabled = true;
            btnExport.Enabled = true;
            btnCancel.Enabled = true;

        }
        public void export()
        {
            if (!bgworker.IsBusy)
            {
                bool chkDataTxt = true;
                if (rdbAuto.Checked == true)
                {
                    //label25.Text = "กำลังตรวจสอบ เวอร์ชั่น";
                    chkDataTxt = CheckAddDataAuto();
                    if (chkDataTxt == false)
                    {
                        return;
                    }
                    string[] arrPathIO = null;
                    string pathfile = myTestKey + "\\" + inputnameFile.Text;
                    StreamReader txtw = new StreamReader(pathfile);
                    String line = txtw.ReadToEnd();
                    byte[] data = Convert.FromBase64String(line);
                    using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                    {
                        byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                        using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                        {
                            ICryptoTransform transform = tripleDes.CreateDecryptor();
                            byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                            //Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                            line = UTF8Encoding.UTF8.GetString(results);
                        }
                    }

                    txtw.Close();
                    arrPathIO = line.Split(',');
                    this.RandomNameOfFolder = RandomNumberAndPassword();
                    System.IO.Directory.CreateDirectory(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder);
                    int lenOutput = arrPathIO[1].Split('\\').Length;
                    _apimail.email = emailtxt.Text;
                    _apimail.taxseller = txtSellerTaxID.Text;
                    _apimail.branch = txtBranchID.Text;
                    _apimail.form = this;
                    _apimail.path = arrPathIO[1].Split('\\')[lenOutput - 2];
                    _apimail.input = arrPathIO[0].Split('\\')[lenOutput - 2];
                    _apimail.timeuser = txtTimeRun.Text;
                    _apimail.typesoft = typeurl;
                    if (bool.Parse(arrPathIO[13]) == true)
                    {
                        //pingeng = PingHost("devinet-etax.one.th");
                        string status = "1";
                        try
                        {
                            Console.WriteLine("=========");
                            if (pingeng)
                            {
                                _apimail.first_export();
                            }


                        }
                        catch (NullReferenceException ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.SuppressFinalize(this);
                        }
                    }
                    else if (bool.Parse(arrPathIO[13]) == false)
                    {
                        //pingeng = PingHost("devinet-etax.one.th");
                        string status = "2";
                        try
                        {
                            if (pingeng)
                            {
                                _apimail.first_export();
                            }
                            //_apimail.first_export(emailtxt.Text, txtSellerTaxID.Text, txtBranchID.Text, arrPathIO[1].Split('\\')[lenOutput - 2], this, arrPathIO[0].Split('\\')[lenOutput - 2], txtTimeRun.Text, typeurl);

                        }
                        catch (NullReferenceException ex)
                        {
                            Console.WriteLine(ex.Message);
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
                    bgworker.RunWorkerAsync();
                    txtInput.Clear();
                    txtOutput.Clear();
                    DisableForm();
                }
                else
                {
                    chkDataTxt = CheckAddDataManual();
                    if (chkDataTxt == false)
                    {
                        return;
                    }

                    RunProcessFormM();
                    txtAmountFile.Clear();
                    txtTimeRun.Clear();
                    EnableForm();
                }
                //DisableForm();
                //EnableForm();
                statusWorker = "Do";
                GC.Collect(1, GCCollectionMode.Forced);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //string[] namePath;
                //namePath = path.Split('\\');
                //MessageBox.Show(Send_start.CallAPISENDMAIL(txtSellerTaxID.Text,txtBranchID.Text,email,namePath[namePath.Length-1]));
            }
            else
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("โปรแกรมกำลังทำงาน ถ้าต้องการยกเลิกให้กดปุ่ม Cancel!!!",
                "Warning",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Information, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                                                                                                      //MessageBox.Show("โปรแกรมกำลังทำงาน ถ้าต้องการยกเลิกให้กดปุ่ม cancel!!!");
            }

        }
        private bool check_copies(string copies)
        {
            if (copies == "0" || copies == "00")
            {
                return false;
            }
            else
            {
                Regex regex = new Regex(@"^[0-9]{1,2}");
                if (regex.IsMatch(copies))
                {
                    return true;
                    //true
                }
            }


            return false;
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {

                export();
                

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
                GC.SuppressFinalize(this);
            }

        }
        public static class PrinterSettings
        {
            [DllImport("winspool.drv",
              CharSet = CharSet.Auto,
              SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern Boolean SetDefaultPrinter(String name);
        }
        public void RunProcessFormM()
        {
            try
            {
                string strPathIO = string.Empty;
                string[] arrPathIO = null;
                ValueReturnForm valueReturn = new ValueReturnForm();
                bool firstStatusRunning = true;
                if (rdbManual.Checked == true)
                {
                    pfIO = new PathFilesIO();
                    pfIO.PathInput = txtInput.Text;
                    pfIO.TypeDoc = TypeDoc;
                    pfIO.PathOutput = txtOutput.Text;
                    if (chkPreview.Checked)
                    {
                        pfIO.TypePrintPreview = "M";
                    }
                    else
                    {
                        pfIO.TypePrintPreview = "A";
                    }
                    if (chkprintauto.Checked)
                    {
                        pfIO.TypePrinting = "A";
                        pfIO.Printer = cboPrinter.SelectedItem.ToString();
                    }
                    else
                    {
                        pfIO.TypePrinting = "M";
                    }

                    if (txtConfixExcel.Text.Equals(""))
                    {
                        string pathDefault = AppDomain.CurrentDomain.BaseDirectory + "ConfigExcel\\default.csv";
                        pfIO.PathConfigExcel = pathDefault;

                    }
                    else
                    {
                        pfIO.PathConfigExcel = txtConfixExcel.Text;
                    }
                    pfIO.TypeRunning = "M";
                    chkServiceURL();
                    strDateFileName = DateTime.Now.ToString("dd-M-yy'T'HH-mm-ss");
                    procress = new ProcressETAX();
                    valueReturn = procress.RunProcess(pfIO, strDateFileName, txtSellerTaxID.Text, txtBranchID.Text, APIKEY, txtUserCode.Text, txtAccessKey.Text, cboServiceCode.SelectedItem.ToString(), txtAmountFile.Text, this.____url, firstStatusRunning, valueReturn, this);

                    if (valueReturn.StatusFindPDF == false)
                    {
                        return;
                    }
                    else
                    {
                        string filename = Path.GetFileName(pfIO.PathInput);
                        if (chkprintauto.Checked && check___copies.Checked == false)
                        {
                            //PrinterSettings.SetDefaultPrinter(cboPrinter.SelectedItem.ToString());
                            //ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                            //{
                            //    UseShellExecute = true,
                            //    Verb = "print",
                            //    CreateNoWindow = true,
                            //    FileName = valueReturn.pathPrint,
                            //    //Arguments = printDialog1.PrinterSettings.PrinterName.ToString(),
                            //    WindowStyle = ProcessWindowStyle.Hidden
                            //};
                            try
                            {
                                etaxOneth_Printer.Class1 _printer = new etaxOneth_Printer.Class1();                                
                                _printer.PrintMethod(valueReturn.pathPrint, cboPrinter.SelectedItem.ToString(), short.Parse(input_copies.Text));                                
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ไม่พบตัวอ่านไฟล์ของคุณ");
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
                        }else if(chkprintauto.Checked && check___copies.Checked == true)
                        {
                            //if(filename.Split('.')[filename.Split('.').Length - 1] == "txt")
                            //{
                            //    MessageBox.Show("check == true");
                            //}
                            //MessageBox.Show("check == true");
                        }
                        txtStatusRunning.Text = "Program Complete!";
                        txtStatusRunning.BackColor = Color.PaleGreen;
                        txtStatusRunning.Refresh();
                        //txtStatus.Clear();
                        pgbLoad.Value = 0;
                        lbPercent.Text = "Export Data: 0%";
                        GC.SuppressFinalize(this);
                        procress.Dispose();
                    }

                }

            }
            catch (Exception ex)
            {

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
        public void RunProcessForm(DoWorkEventArgs ex)
        {
        Found:
            //Console.WriteLine(ex);
            procress.Dispose();
            if (bgworker.CancellationPending == true)
            {
                ex.Cancel = true;
                return;
            }
            try
            {

                string strPathIO = string.Empty;
                string[] arrPathIO = null;
                ValueReturnForm valueReturn = new ValueReturnForm();
                bool firstStatusRunning = true;

                //strDateFileName = DateTime.Now.ToString("dd-M-yy_HH-mm-ss");
                try
                {
                    txtStatusRunning.Text = "Program Running!";
                    txtStatusRunning.BackColor = Color.Salmon;
                    txtStatusRunning.Refresh();
                    if (statusWorker != "cancel")
                    {
                        Microsoft.Win32.RegistryKey rkey;
                        rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
                        myTestKey = (string)rkey.GetValue("PathConfigETAX");
                        string pathfile = myTestKey + "\\" + inputnameFile.Text;
                        StreamReader txtw = new StreamReader(pathfile);
                        String line = txtw.ReadToEnd();
                        byte[] data = Convert.FromBase64String(line);
                        using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                        {
                            byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                            using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                            {
                                ICryptoTransform transform = tripleDes.CreateDecryptor();
                                byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                                //Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                                line = UTF8Encoding.UTF8.GetString(results);
                            }
                        }
                        arrPathIO = line.Split(',');
                        //strPathIO = System.IO.File.ReadAllText(myTestKey + "\\" + inputnameFile.Text);
                        //arrPathIO = strPathIO.Split(',');

                        txtInput.Text = arrPathIO[0];
                        txtOutput.Text = arrPathIO[1];
                        //txtInput.Refresh();
                        //txtOutput.Refresh();
                        txtw.Close();
                    }

                }
                catch
                {
                    InputBox.SetLanguage(InputBox.Language.English);
                    InputBox.ShowDialog("Please Select Path Input and Output to Default Before Run Auto Option!",
                    "Warning",   //Text message (mandatory), Title (optional)
                    InputBox.Icon.Information, //Set icon type (default info)
                    InputBox.Buttons.Ok, //Set buttons (default ok)
                    InputBox.Type.Nothing, //Set type (default nothing)
                    null, //String field as ComboBox items (default null)
                    true, //Set visible in taskbar (default false)
                    new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                                                                                                          //MessageBox.Show("Please Select Path Input and Output to Default Before Run Auto Option!");
                    return;
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
                if (statusWorker == "Do")
                {
                    bool checkfilename;
                    string[] arrAllFileCheck = System.IO.Directory.GetFiles(arrPathIO[0]);
                    List<string> listAllFile = new List<string>(arrAllFileCheck);
                    string[] arrFilespcfg = System.IO.Directory.GetFiles(arrPathIO[0], "*.pcfg");
                    if (check___copies.Checked == true)
                    {
                        for (var i = 0; i < arrFilespcfg.Count(); i++)
                        {
                        
                            try
                            {
                                File.Move(arrFilespcfg[i], arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder + "\\" + Path.GetFileName(arrFilespcfg[i]));
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }
                    }
                    for (var i = 0; i < listAllFile.Count(); i++)
                    {
                        string strFileName = Path.GetFileNameWithoutExtension(listAllFile[i]);
                        var pattern = @",";
                        checkfilename = Regex.IsMatch(strFileName, pattern);
                        Console.WriteLine(checkfilename);
                        if (checkfilename == false)
                        {
                            try
                            {
                                File.Move(listAllFile[i], arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder + "\\" + Path.GetFileName(listAllFile[i]));
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }
                        else
                        {
                            try
                            {

                                if (!Directory.Exists(arrPathIO[1] + "\\" + "Fail" + "\\" + "FileError"))
                                {
                                    Directory.CreateDirectory(arrPathIO[1] + "\\" + "Fail" + "\\" + "FileError");
                                }
                                File.Move(listAllFile[i], arrPathIO[1] + "\\" + "Fail" + "\\" + "FileError" + "\\" + Path.GetFileName(listAllFile[i]));
                                listAllFile.Remove(listAllFile[i]);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            finally
                            {
                                for (int p = 0; p <= GC.MaxGeneration; p++)
                                {
                                    int count = GC.CollectionCount(p);
                                    GC.Collect();
                                }
                                GC.SuppressFinalize(this);
                            }

                        }
                    }
                    string[] arrAllFile = listAllFile.ToArray();

                    string patternChkString = @"([a-zA-Zก-๙0-9/])";
                    bool chkSting = false;
                    foreach (var item in arrAllFile)
                    {

                        //MessageBox.Show(Path.GetFileNameWithoutExtension(item));
                        try
                        {
                            File.Move(item, arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder + "\\" + Path.GetFileName(item));
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                    string[] arrFilesXlsx = System.IO.Directory.GetFiles(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder, "*.xlsx");
                    string[] arrFilesXls = System.IO.Directory.GetFiles(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder, "*.xls");
                    string[] arrFilesTxt = System.IO.Directory.GetFiles(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder, "*.txt");
                    string[] arrFilesCsv = System.IO.Directory.GetFiles(arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder, "*.csv");
                    
                    for (var i = 0; i < arrFilesXls.Length; i++)
                    {
                        if (Path.GetExtension(arrFilesXls[i]) == ".xls")
                        {
                            arrFilesXls[i] = arrFilesXls[i];
                        }
                        else
                        {
                            arrFilesXls[i] = "";
                        }
                    }
                    pfIO = new PathFilesIO();
                    pfIO.TypeRunning = "A";
                    pfIO.PathInput = arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder;
                    pfIO.PathOutput = arrPathIO[1];
                    pfIO.PathTemp = arrPathIO[0] + "\\InputTemp\\Temp";
                    pfIO.PathLogFileRun = arrPathIO[0] + "\\InputTemp\\LogFileRun";
                    pfIO.PathFileRun = arrPathIO[0] + "\\InputTemp\\FileRun";
                    pfIO.LogTimeProcess = arrPathIO[0] + "\\InputTemp\\LogProcess";
                    
                    pfIO.TypeDoc = TypeDoc;
                    chkServiceURL();
                    if (chkPreview.Checked)
                    {
                        pfIO.TypePrintPreview = "M";
                    }
                    else
                    {
                        pfIO.TypePrintPreview = "A";
                    }
                    if (chkprintauto.Checked)
                    {
                        pfIO.TypePrinting = "A";
                        pfIO.Printer = cboPrinter.SelectedItem.ToString();

                    }
                    else
                    {
                        pfIO.TypePrinting = "M";
                        pfIO.Printer = "";
                    }
                    if (arrPathIO[12].Equals(""))
                    {
                        string pathDefault = AppDomain.CurrentDomain.BaseDirectory + "ConfigExcel\\default.csv";
                        pfIO.PathConfigExcel = pathDefault;

                    }
                    else
                    {
                        pfIO.PathConfigExcel = arrPathIO[12];
                    }
                    //pfIO.PathConfigExcel = arrPathIO[12];
                    System.IO.Directory.CreateDirectory(pfIO.PathTemp);
                    System.IO.Directory.CreateDirectory(pfIO.PathLogFileRun);
                    System.IO.Directory.CreateDirectory(pfIO.PathFileRun);
                    System.IO.Directory.CreateDirectory(pfIO.LogTimeProcess);
                    //CreateTextFile(pfIO.PathLogFileRun + "\\LogFileRunPerTime.txt", "");

                    if (arrFilesXlsx.Length > 0 || arrFilesTxt.Length > 0 || arrFilesXls.Length > 0 || arrFilesCsv.Length > 0)
                    {
                        FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                        string version = fvi.FileVersion;
                        strDateFileName = DateTime.Now.ToString("dd-M-yy'T'HH-mm-ss");
                        pfIO.DateTimeFolderName = strDateFileName;
                        //pfIO.PathErr = arrPathIO[1] + "\\" + strDateFileName + "\\Fail\\Error";
                        //pfIO.PathSource_F = arrPathIO[1] + "\\" + strDateFileName + "\\Fail\\Source";
                        //pfIO.PathSource_S = arrPathIO[1] + "\\" + strDateFileName + "\\Success\\Source";
                        //pfIO.PathSuccess_O = arrPathIO[1] + "\\" + strDateFileName + "\\Success\\Output";
                        //pfIO.PathLogTime = arrPathIO[1] + "\\" + strDateFileName + "\\LogTime";
                        pfIO.PathErr = arrPathIO[1] + "\\Fail";
                        pfIO.PathSource_F = arrPathIO[1] + "\\Fail\\Source";
                        pfIO.PathSource_S = arrPathIO[1] + "\\Success\\Source";
                        //pfIO.PathSource_S = arrPathIO[1] + "\\Success\\Source"; ถ้าใช้งานต้องไปเปิด ใน OnethProcess บรรทัดที่ 1001-1020 ด้วย
                        pfIO.PathSuccess_O = arrPathIO[1] + "\\Success";
                        pfIO.PathLogTime = arrPathIO[1] + "\\LogTime";
                        pfIO.BCP_Folder = arrPathIO[1] + "\\BCPFolder";
                        System.IO.Directory.CreateDirectory(pfIO.PathErr);
                        System.IO.Directory.CreateDirectory(pfIO.PathSource_F);
                        System.IO.Directory.CreateDirectory(pfIO.PathSource_S); //ถ้าใช้งานต้องไปเปิด บรรทัด377 ก่อน
                        System.IO.Directory.CreateDirectory(pfIO.PathSuccess_O);
                        System.IO.Directory.CreateDirectory(pfIO.PathLogTime);
                        System.IO.Directory.CreateDirectory(pfIO.BCP_Folder);
                        statusProgram = 1;
                        while (valueReturn.StatusRunning == false)
                        {
                            if (statusWorker == "Do")
                            {
                                procress = new ProcressETAX();
                                System.IO.FileInfo ficheck = null;
                                try
                                {
                                    ficheck = new System.IO.FileInfo(strDateFileName);
                                    if (ReferenceEquals(ficheck, null))
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        valueReturn = procress.RunProcess(pfIO, strDateFileName, txtSellerTaxID.Text, txtBranchID.Text, APIKEY, txtUserCode.Text, txtAccessKey.Text, cboServiceCode.SelectedItem.ToString(), txtAmountFile.Text, this.____url, firstStatusRunning, valueReturn, this);
                                    }
                                }
                                catch (ArgumentException)
                                {

                                }
                                catch (System.IO.PathTooLongException)
                                {

                                }
                                catch (NotSupportedException)
                                {

                                }
                                catch (Exception)
                                {

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

                                try
                                {
                                    if (pingeng == true)
                                    {
                                        _apimail.next_export();
                                    }

                                }
                                catch (NullReferenceException e)
                                {
                                    Console.WriteLine(e.Message);
                                }
                                catch (Exception)
                                {

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
                                    //getsocket.Dispose();

                                }
                                //procressing.Show();
                                //toProcess = new etaxOnethProcess();
                                //valueReturn = procressing.RunProcess(pfIO, strDateFileName, txtSellerTaxID.Text, txtBranchID.Text, txtAPIKey.Text, txtUserCode.Text, txtAccessKey.Text, cboServiceCode.SelectedItem.ToString(), txtAmountFile.Text, cboServiceURL.SelectedItem.ToString(), firstStatusRunning, valueReturn);
                                //frmProcess = new etaxOnethProcess();
                                //frmProcess.Show();
                                //valueReturn = frmProcess.RunProcess(pfIO, strDateFileName, txtSellerTaxID.Text, txtBranchID.Text, txtAPIKey.Text, txtUserCode.Text, txtAccessKey.Text, cboServiceCode.SelectedItem.ToString(), txtAmountFile.Text, cboServiceURL.SelectedItem.ToString(), firstStatusRunning, valueReturn);
                                //frmProcess.Close();
                                //frmProcess.Dispose();
                                firstStatusRunning = false;
                                strDateFileName = DateTime.Now.ToString("dd-M-yy'T'HH-mm-ss");
                            }
                            else
                            {
                                valueReturn.StatusRunning = false;
                                ex.Cancel = true;
                                return;
                            }
                        }
                    }
                    else
                    {
                        valueReturn.StatusRunning = true;
                        statusProgram = 0;

                        //txtStatusRunning.Text = "File Not Found Run Again in " + txtTimeRun.Text + " Second";
                    }
                    if (statusProgram == 1)
                    {
                        txtStatusRunning.Text = "Program Complete! Run Again in " + txtTimeRun.Text + " Second";
                        txtStatusRunning.BackColor = Color.PaleGreen;
                        txtStatusRunning.Refresh();
                        //txtStatus.Clear();
                        pgbLoad.Value = 0;
                        lbPercent.Text = "Export Data: 0%";
                    }
                    else
                    {
                        txtStatusRunning.Text = "File Not Found";
                        txtStatusRunning.BackColor = Color.Salmon;
                        txtStatusRunning.Refresh();
                        Thread.Sleep(1000);
                        txtStatusRunning.Text = "Run Again in " + txtTimeRun.Text + " Second";
                        txtStatusRunning.BackColor = Color.PaleGreen;
                        txtStatusRunning.Refresh();
                        GC.SuppressFinalize(this);
                        GC.Collect(1, GCCollectionMode.Forced);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();



                    }
                    //MessageBox.Show(valueReturn.StatusRunning + "");
                    if (valueReturn.StatusRunning == true && statusWorker == "Do")
                    {
                        try
                        {
                            if (txtTimeRun.Text != null)
                            {
                                string iSec = (txtTimeRun.Text);
                                InvokeMethod(iSec);
                                try
                                {
                                    if (pingeng == true)
                                    {
                                        _apimail.next_export();
                                    }
                                }
                                catch (NullReferenceException e)
                                {
                                    Console.WriteLine(e.Message);
                                }
                                catch (Exception)
                                {

                                }
                                finally
                                {
                                    for (int i = 0; i <= GC.MaxGeneration; i++)
                                    {
                                        int count = GC.CollectionCount(i);
                                        GC.Collect();
                                    }
                                    GC.SuppressFinalize(this);
                                    //getsocket.Dispose();
                                }


                                if (statusWorker == "Do")
                                {
                                    GC.SuppressFinalize(this);
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    procress.Dispose();
                                    //RunProcessForm(ex);

                                    goto Found;

                                }

                            }


                            //MessageBox.Show("File Not Found!!!");
                        }
                        catch (Exception e)
                        {

                        }
                    }
                }

            }
            catch (Exception e)
            {

            }

        }

        public void CreateTextFile(string strPath, string strData)
        {
            try
            {
                TextWriter txtw = new StreamWriter(strPath);
                txtw.Write(strData);
                txtw.Close();
            }
            catch (FieldAccessException e)
            {
                //MessageBox.Show(e.ToString());
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

        public void chkServiceURL()
        {
            if (cboServiceURL.SelectedIndex.ToString() == "0")
            {
                this.APIKEY = "AK2-3UY8R84Q6L5KAOLE781044GRX2LZTJ95WHJDZWOORWXVOR92XUJFDTJ02FHTFVYFLQXF6QAHRNGJHSUNXGFZT65EBMHNRUO7J0IARSEBFMI8BN6OHT1M0BJGBAF4PB0HMMRTALW8QD4LCNDRNSXJOFNR5KWQOEPS7EKBVZLTSBJBGZ3676YJXLP96GNZ36CFC00HIVZHGO3DHO14ZEIEZN8AY99VKU0WR1QEM3DW3YZQE11OZ2UINCGIQN5ILDAKJ";
            }
            else if (cboServiceURL.SelectedIndex.ToString() == "1")
            {
                this.APIKEY = "AK2-3UY8R84Q6GFXD8QXZMFYNPUJRFOKWOT614C5GAAC88OVAGLD8F1HWAPB9LW05QQSACKVSN0FBI1H8WPO0NWU29MWO9854BBMW5OJ6IZOBUJFZRHV0ZPD4CP02PXLT95YHD6QX3TH01CZX7TJ7X4NBLKZLGQP8NF3BKZIU6NSW463VL61A8LBBAKIJSQS7M1TL2S50E6AGWRE85ZSK6AIHYDYXY7C19LKLFRFWXW3FAOAB61O0E5TGPBNC3P4X72BU";
            }
        }

        static void InvokeMethod(string iSecond)
        {
            string strTime = iSecond + "000";
            Thread.Sleep(Int32.Parse(strTime));
        }

        public bool CheckAddDataManual()
        {
            if (txtInput.Text.Equals("") || txtOutput.Text.Equals("") || txtSellerTaxID.Text.Equals("") || txtBranchID.Text.Equals("") ||
                txtUserCode.Text.Equals("") || txtAccessKey.Text.Equals("") || cboServiceCode.Text.Equals("") || this.____url.Equals("") ||
                cboTypeDoc.Text.Equals(""))
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("Please fill full the blank!",
                "Warning",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Information, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                //MessageBox.Show("Please fill full the blank!");
                return false;
            }
            else
            {
                if (chkprintauto.Checked == true)
                {
                    if (cboPrinter.SelectedIndex.ToString() == "-1")
                    {
                        InputBox.SetLanguage(InputBox.Language.English);
                        InputBox.ShowDialog("Please fill full the blank!",
                        "Warning",   //Text message (mandatory), Title (optional)
                        InputBox.Icon.Information, //Set icon type (default info)
                        InputBox.Buttons.Ok, //Set buttons (default ok)
                        InputBox.Type.Nothing, //Set type (default nothing)
                        null, //String field as ComboBox items (default null)
                        true, //Set visible in taskbar (default false)
                        new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)

                        return false;
                    }
                }
            }
            return true;
        }

        public bool CheckAddDataAuto()
        {
            if (txtSellerTaxID.Text.Equals("") || txtBranchID.Text.Equals("") || txtUserCode.Text.Equals("") || txtAccessKey.Text.Equals("") ||
                cboServiceCode.Text.Equals("") || txtAmountFile.Text.Equals("") || this.____url.Equals("") || txtTimeRun.Text.Equals("") ||
                cboTypeDoc.Text.Equals(""))
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("กรุณาเพิ่มข้อมูลที่ Config",
                "Warning",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Information, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                //MessageBox.Show("Please fill full the blank!");
                return false;
            }
            else
            {
                if (chkprintauto.Checked == true)
                {
                    if (cboPrinter.SelectedIndex.ToString() == "-1")
                    {
                        InputBox.SetLanguage(InputBox.Language.English);
                        InputBox.ShowDialog("Please fill full the blank!",
                        "Warning",   //Text message (mandatory), Title (optional)
                        InputBox.Icon.Information, //Set icon type (default info)
                        InputBox.Buttons.Ok, //Set buttons (default ok)
                        InputBox.Type.Nothing, //Set type (default nothing)
                        null, //String field as ComboBox items (default null)
                        true, //Set visible in taskbar (default false)
                        new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)

                        return false;
                    }
                }
            }
            return true;
        }

        public void DisableForm()
        {
            try
            {
                inputnameFile.Enabled = false;
                pictureBox2.Enabled = false;
                pictureBox4.Enabled = false;
                //metroToggle1.Enabled = false;
                label19.Enabled = false;
                input_copies.Enabled = false;
                txtAccessKey.Enabled = false;
                //txtAPIKey.Enabled = false;
                txtBranchID.Enabled = false;
                txtInput.Enabled = false;
                txtOutput.Enabled = false;
                cboServiceCode.Enabled = false;
                txtSellerTaxID.Enabled = false;
                txtUserCode.Enabled = false;
                btnInput.Enabled = false;
                btnOutput.Enabled = false;
                btnExport.Enabled = false;
                rdbAuto.Enabled = false;
                rdbManual.Enabled = false;
                txtAmountFile.Enabled = false;
                txtTimeRun.Enabled = false;
                cboServiceURL.Enabled = false;
                btnHome.Enabled = false;
                emailtxt.Enabled = false;
                cboPrinter.Enabled = false;
                cboTypeDoc.Enabled = false;
                chkPreview.Enabled = false;
                chkprintauto.Enabled = false;
                chkprintauto.Refresh();
                check___copies.Enabled = false;
                check___copies.Refresh();
                txtAccessKey.Refresh();
                //txtAPIKey.Refresh();
                txtBranchID.Refresh();
                txtInput.Refresh();
                txtOutput.Refresh();
                cboServiceCode.Refresh();
                txtSellerTaxID.Refresh();
                txtUserCode.Refresh();
                btnInput.Refresh();
                btnOutput.Refresh();
                btnExport.Refresh();
                rdbAuto.Refresh();
                rdbManual.Refresh();
                txtAmountFile.Refresh();
                txtTimeRun.Refresh();
                cboServiceURL.Refresh();
                btnHome.Refresh();
                emailtxt.Refresh();
                cboPrinter.Refresh();
                chkPreview.Refresh();
            }
            catch (Exception ea)
            {

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
        public void clearform()
        {
            txtAccessKey.Clear();
            txtSellerTaxID.Clear();
            txtBranchID.Clear();
            //txtAPIKey.Clear();
            txtUserCode.Clear();
            txtInput.Clear();
            txtOutput.Clear();
            //txtPathFile.Clear();
            txtStatusRunning.Clear();
            emailtxt.Clear();
        }
        public void EnableForm()
        {
            txtAccessKey.Enabled = true;
            //txtAPIKey.Enabled = true;
            txtBranchID.Enabled = true;
            cboServiceCode.Enabled = true;
            txtSellerTaxID.Enabled = true;
            txtUserCode.Enabled = true;
            btnExport.Enabled = true;
            rdbAuto.Enabled = true;
            rdbManual.Enabled = true;
            txtAmountFile.Enabled = true;
            txtTimeRun.Enabled = true;
            cboServiceURL.Enabled = true;
            btnHome.Enabled = true;
            emailtxt.Enabled = true;

            if (rdbAuto.Checked == true)
            {
                txtInput.Enabled = false;
                txtOutput.Enabled = false;
                btnInput.Enabled = false;
                btnOutput.Enabled = false;
                emailtxt.Enabled = false;
            }
            else
            {
                txtInput.Enabled = true;
                txtOutput.Enabled = true;
                btnInput.Enabled = true;
                btnOutput.Enabled = true;
                txtAmountFile.Enabled = false;
                txtTimeRun.Enabled = false;

            }
        }

        private void pnHead_MouseDown(object sender, MouseEventArgs e)
        {
            TogMove = 1;
            MValX = e.X;
            MValY = e.Y;

        }

        private void pnHead_MouseMove(object sender, MouseEventArgs e)
        {
            if (TogMove == 1)
            {
                this.SetDesktopLocation(MousePosition.X - MValX, MousePosition.Y - MValY);
            }

        }

        private void pnHead_MouseUp(object sender, MouseEventArgs e)
        {
            TogMove = 0;

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            RunProcessForm(e);
            //if(backgroundWorker1.CancellationPending == true)
            //{
            //    e.Cancel = true;
            //}

        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (!bgworker.IsBusy)
            {
                try
                {


                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine(ex.Message);
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
                Config CF = new Config();
                this.Close();
                CF.Show();
            }
            else
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("กรุณายกเลิกการทำงานก่อน",
                "Warning",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Information, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                //MessageBox.Show("กรุณายกเลิกการทำงานก่อน");
            }

        }

        private void btnInputAuto_Click(object sender, EventArgs e)
        {

        }

        private void textFilenameChange(object sender, EventArgs e)
        {
            try
            {
                cboServiceCode.Items.Clear();
                cboServiceURL.Items.Clear();
                cboServiceCode.Items.Add("S03");
                cboServiceCode.Items.Add("S06");
                cboServiceCode.Items.Add("S06(Excel Only)");
                cboServiceCode.Items.Add("S03(Excel Only)");
                cboServiceCode.Items.Add("S06(Excel Only & List Item)");
                cboServiceCode.Items.Add("BCP Service");
                cboServiceURL.Items.Add("ทดสอบระบบ (UAT)");
                cboServiceURL.Items.Add("Production");
                //cboServiceURL.Items.Add("https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument");
                //cboServiceURL.Items.Add("https://etaxsp.one.th/etaxdocumentws/etaxsigndocument");
                Microsoft.Win32.RegistryKey rkey;
                rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
                myTestKey = (string)rkey.GetValue("PathConfigETAX");
                rkey.Close();
                if (inputnameFile.Text.Equals(""))
                {
                    //InputBox.SetLanguage(InputBox.Language.English);
                    //InputBox.ShowDialog("ข้อมูลไม่ถูกต้อง..",
                    //"Warning",   //Text message (mandatory), Title (optional)
                    //InputBox.Icon.Information, //Set icon type (default info)
                    //InputBox.Buttons.Ok, //Set buttons (default ok)
                    //InputBox.Type.Nothing, //Set type (default nothing)
                    //null, //String field as ComboBox items (default null)
                    //true, //Set visible in taskbar (default false)
                    //new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular)); //Set font (default by system)
                }
                else
                {
                    StreamReader sr = new StreamReader(myTestKey + "\\" + inputnameFile.Text);
                    String line = sr.ReadToEnd();
                    byte[] data = Convert.FromBase64String(line);
                    using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                    {
                        byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                        using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                        {
                            ICryptoTransform transform = tripleDes.CreateDecryptor();
                            byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                            //Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                            line = UTF8Encoding.UTF8.GetString(results);
                        }
                    }
                    int len = line.Split(',').Length;
                    txtSellerTaxID.Text = line.Split(',')[2];
                    txtBranchID.Text = line.Split(',')[3];
                    txtAPIKey.Text = APIKEY;
                    txtUserCode.Text = line.Split(',')[5];
                    txtAccessKey.Text = line.Split(',')[6];
                    cboServiceCode.SelectedItem = line.Split(',')[7];
                    cboServiceURL.SelectedItem = line.Split(',')[8];

                    if(line.Split(',')[8] == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
                    {
                        cboServiceURL.SelectedIndex = 0;
                        this.____url = "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument";
                    }
                    else if(line.Split(',')[8] == "https://etaxsp.one.th/etaxdocumentws/etaxsigndocument")
                    {
                        cboServiceURL.SelectedIndex = 1;
                        this.____url = "https://etaxsp.one.th/etaxdocumentws/etaxsigndocument";
                    }
                    txtAmountFile.Text = line.Split(',')[9];
                    txtTimeRun.Text = line.Split(',')[10];
                    emailtxt.Text = line.Split(',')[11];
                    txtConfixExcel.Text = line.Split(',')[12];
                    try
                    {
                        ___copise = line.Split(',')[17];
                    }
                    catch (Exception ex)
                    {
                        ___copise = "1";
                    }
                    try
                    {
                        set__copise_file = line.Split(',')[16];
                    }
                    catch (Exception ex)
                    {
                        set__copise_file = "false";
                    }
                    
                    bool checkAutoPrint = false;
                    try
                    {
                        checkAutoPrint = bool.Parse(line.Split(',')[13]);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        checkAutoPrint = false;
                    }
                    //if (checkAutoPrint)
                    //{

                    chkprintauto.Checked = checkAutoPrint;
                    cboPrinter.SelectedItem = line.Split(',')[14];
                    //}
                    sr.Close();
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
        public void Outputmessage(string message)
        {
            this.txtStatus.Text += message + System.Environment.NewLine;
        }
        public void OutputPrc(int val, string txtval)
        {
            pgbLoad.Value = val;
            lbPercent.Text = txtval;
            lbPercent.Refresh();
        }


        private void btnPathFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFileDialog = new OpenFileDialog();
            string strPathIO = string.Empty;
            string[] arrPathIO = null;

            if (oFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //txtPathFile.Text = oFileDialog.FileName;
                //strPathIO = System.IO.File.ReadAllText(txtPathFile.Text);
                arrPathIO = strPathIO.Split(',');
                path = arrPathIO[0];
                txtSellerTaxID.Text = arrPathIO[2];
                txtBranchID.Text = arrPathIO[3];
                txtAPIKey.Text = APIKEY;

                //email = arrPathIO[5];
                txtSellerTaxID.Enabled = false;
                txtBranchID.Enabled = false;
                //txtAPIKey.Enabled = false;
                //txtPathFile.Enabled = false;
            }
        }
        //public void uploadText(string text)
        //{
        //    texttest = text;
        //    MessageBox.Show(texttest);
        //}


        private void Scroll_txt(object sender, EventArgs e)
        {
            txtStatus.SelectionStart = txtStatus.TextLength;
            txtStatus.ScrollToCaret();
        }

        private void setValue(object sender, EventArgs e)
        {

        }

        private void showEtaxOne(object sender, EventArgs e)
        {
            _apimail.form = this;
            _apimail._checkversion_program();

            string[] arrPathIO = null;
            string pathfile = myTestKey + "\\" + inputnameFile.Text;
            if (inputnameFile.Text.Length != 0)
            {
                StreamReader txtw = new StreamReader(pathfile);
                if (txtw != null)
                {
                    String line = txtw.ReadToEnd();
                    byte[] data = Convert.FromBase64String(line);
                    using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                    {
                        byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                        using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                        {
                            ICryptoTransform transform = tripleDes.CreateDecryptor();
                            byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                            //Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                            line = UTF8Encoding.UTF8.GetString(results);
                        }
                    }

                    txtw.Close();
                    arrPathIO = line.Split(',');
                    int lenOutput = arrPathIO[1].Split('\\').Length;
                    _apimail.email = emailtxt.Text;
                    _apimail.taxseller = txtSellerTaxID.Text;
                    _apimail.branch = txtBranchID.Text;
                    _apimail.path = arrPathIO[1].Split('\\')[lenOutput - 2];
                    _apimail.input = arrPathIO[0].Split('\\')[lenOutput - 2];
                    _apimail.timeuser = txtTimeRun.Text;
                    _apimail.typesoft = typeurl;
                    if (txtSellerTaxID.Text.Length != 0)
                    {
                        try
                        {
                            _apimail.start_prog();
                            Thread.Sleep(2000);
                            _apimail.first_export();
                        }
                        catch (NullReferenceException ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        Thread.Sleep(500);
                    }
                }

            }
            try
            {
                //getsocket.checkSocket();
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine(ex.Message);
            }


        }




        public void closeSocket()
        {
            try
            {
                //getsocket.EndSocketAuto("CloseApplication", this);
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //string content = File.ReadAllText(txtInput.Text, Encoding.GetEncoding(874));
            //MessageBox.Show(content);
        }

        private void btnConfixExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFileDialog = new OpenFileDialog();
            if (oFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtConfixExcel.Text = oFileDialog.FileName;
            }
        }



        private void cboServiceURL_TextChanged(object sender, EventArgs e)
        {

        }

        private void cboServiceURL_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboServiceURL.SelectedIndex == 0)
            {
                lbUrl.Text = "ทดสอบระบบ (UAT)";
                typeurl = "UAT";
                this.____url = "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument";
            }
            else if (cboServiceURL.SelectedIndex == 1)
            {
                lbUrl.Text = "Production";
                typeurl = "PROD";
                this.____url = "https://etaxsp.one.th/etaxdocumentws/etaxsigndocument";
            }
            else if (cboServiceURL.SelectedIndex == -1)
            {
                lbUrl.Text = "";
                typeurl = "";
            }
        }

        public void percText(int text)
        {

        }

        //สำหรับสร้างชื่อโฟลเดอร์แบบแรนด้อม
        public string RandomNumberAndPassword()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(RandomString(4, true));
            builder.Append(RandomNumber(1000, 9999));
            builder.Append(RandomString(2, false));
            return builder.ToString();
        }
        public int RandomNumber(int min, int max)
        {
            Random random = new Random();
            return random.Next(min, max);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            //.Trim(' ', '\r', '\n')
            
            string namefile = "update\\etax-setup" + txtversioncurrent + ".exe";
            string myWebUrlFile = "https://devinet-etax.one.th/apiprog/api/downloadfile/" + txtversioncurrent;

            //
            //progressDlg.ProgressValue = 20;
            //progressDlg.ShowDialog();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;

            DialogResult result = MessageBox.Show("คุณต้องการอัพเดทเวอร์ชั่นโปรแกรม ETAX ONE TH เป็นเวอร์ชั่นล่าสุดหรือไม่?" + Environment.NewLine + "เวอร์ชั่นปัจจุบันของคุณ : " + version + " เวอร์ชั่นล่าสุด : " + txtversioncurrent + "", "สถานะการอัพเดท", MessageBoxButtons.OKCancel);

            if (result == DialogResult.OK)
            {
                Console.WriteLine(myWebUrlFile);
                sw.Start();

                if (!System.IO.Directory.Exists("update"))
                {
                    System.IO.Directory.CreateDirectory("update");
                }
                progressDlg = new ProgressDlg();
                //WebClient client = new WebClient();
                using (WebClient client = new WebClient())
                {
                    try
                    {
                        client.DownloadProgressChanged += wc_DownloadProgressChanged;
                        client.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
                        client.DownloadFileAsync(new Uri(myWebUrlFile), namefile);
                        // successful...
                    }
                    catch (WebException ex)
                    {
                        // failed...
                        Console.WriteLine(ex.Message);
                    }
                }
                   


                progressDlg.ShowDialog();
            }


        }
        void wc_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            //progressDlg.label1.Text = "สถานะการดาวน์โหลดไฟล์ : " + e.ProgressPercentage + " %";    
            progressDlg.ProgressValue = e.ProgressPercentage;
            //Console.WriteLine(e.ProgressPercentage);
            progressDlg.label2.Text = string.Format("{0} MB's / {1} MB's", (e.BytesReceived / 1024d / 1024d).ToString("0.00"), (e.TotalBytesToReceive / 1024d / 1024d).ToString("0.00"));
            progressDlg.label1.Text = string.Format("ความเร็วอินเทอร์เน็ต {0} kb/s", (e.BytesReceived / 1024d / sw.Elapsed.TotalSeconds).ToString("0.00"));
        }
        void Completed(object sender, AsyncCompletedEventArgs e)
        {
            try
            {
                Process.Start("update\\etax-setup" + txtversioncurrent + ".exe");

            }
            catch (Win32Exception aaaa)
            {
                Console.WriteLine(aaaa.Message);
            }
            Application.Exit();

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            releasenote note = new releasenote();
            note.ShowDialog();
        }

        private void chkprintauto_CheckedChanged(object sender, EventArgs e)
        {
            if (chkprintauto.Checked == true)
            {
                check___copies.Visible = true;
                input_copies.Visible = true;
                label19.Visible = true;
                //label24.Visible = true;
                if(rdbManual.Checked == true)
                {
                    this.___copise = "1";
                    check___copies.Visible = false;
                }
                input_copies.Text = this.___copise;
                input_copies.Refresh();
                cboPrinter.Visible = true;
                try
                {
                    if (bool.Parse(set__copise_file) == true)
                    {
                        check___copies.Checked = true;
                    }
                    else if (bool.Parse(set__copise_file) == false)
                    {
                        check___copies.Checked = false;
                    }
                }
                catch (Exception ex)
                {
                    if (rdbManual.Checked)
                    {
                        check___copies.Checked = false;
                        input_copies.Text = "1";
                    }
                }
                
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    cboPrinter.Items.Add(printer);
                }
            }
            else
            {
                try
                {
                    if (bool.Parse(set__copise_file) == true)
                    {
                        check___copies.Checked = true;
                    }
                    else if (bool.Parse(set__copise_file) == false)
                    {
                        check___copies.Checked = false;
                    }
                }
                catch (Exception ex)
                {
                    if (rdbManual.Checked)
                    {
                        check___copies.Checked = false;
                        input_copies.Text = "1";
                    }                    
                }
                input_copies.Text = "0";
                input_copies.Refresh();
                check___copies.Visible = false;
                input_copies.Visible = false;
                label19.Visible = false;
                //label24.Visible = false;
                cboPrinter.Visible = false;
                cboPrinter.Items.Clear();
            }
        }

        private void etaxOneth_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // modify the drag drop effects to Move
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                // no need for any drag drop effect
                e.Effect = DragDropEffects.None;
            }
        }

        private void etaxOneth_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files.Length == 1 && rdbManual.Checked == true)
                {
                    string typefile = Path.GetExtension(files.First());
                    if (files != null && files.Any() && typefile == ".txt" || typefile == ".xls" || typefile == ".xlsx" || typefile == ".pdf" || typefile == ".csv" )
                    {
                        txtInput.Text = files.First();
                    }
                }
                else if (rdbAuto.Checked == true)
                {
                    string inputpath = txtInput.Text;
                    Console.WriteLine(inputpath);
                    if (files != null && files.Any())
                    {
                        for (var i = 0; i < files.Length; i++)
                        {
                            string typefile = Path.GetExtension(files[i]);
                            try
                            {
                                if (typefile == ".txt" || typefile == ".xls" || typefile == ".xlsx" || typefile == ".pdf" || typefile == ".csv" || typefile == ".pcfg")
                                {
                                    string namefile = files[i];
                                    Console.WriteLine(inputpath + "\\" + Path.GetFileName(namefile));
                                    File.Copy(namefile, inputpath + "\\" + Path.GetFileName(namefile));
                                }
                            }
                            catch (IOException iox)
                            {
                                Console.WriteLine(iox.Message);
                            }

                        }

                    }
                    //File.Copy(files, @"otherDirectory\someFile.txt");
                }


                // Your desired code goes here to process the file(s) being dropped
            }
        }

        private void txtStatus_MouseClick(object sender, MouseEventArgs e)
        {
            if (txtStatus.TextLength != 0)
            {

                string outputpath = txtOutput.Text + "\\LogTime";
                var directory = new DirectoryInfo(outputpath);
                var myFile = (from f in directory.GetFiles()
                              orderby f.LastWriteTime descending
                              select f).First();
                Process.Start(outputpath + "\\" + myFile);
            }
        }

        private void chkPreview_CheckedChanged(object sender, EventArgs e)
        {
            if (chkPreview.Checked == true)
            {
                RegistryKey adobe = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Adobe");
                if (adobe == null)
                {
                    var policies = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Policies");
                    if (null == policies)
                        return;
                    adobe = policies.OpenSubKey("Adobe");
                }
                if (adobe != null)
                {
                    RegistryKey acroRead = adobe.OpenSubKey("Acrobat Reader");
                    if (acroRead != null)
                    {
                        string[] acroReadVersions = acroRead.GetSubKeyNames();
                        if (acroReadVersions.Length == 1)
                        {

                        }
                        else
                        {
                            chkPreview.Checked = false;
                        }

                    }
                }
                else
                {
                    chkPreview.Checked = false;
                }
            }
        }

        private void cboTypeDoc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboTypeDoc.SelectedIndex == 0)
            {
                TypeDoc = "";
            }
            else if (cboTypeDoc.SelectedIndex == 1)
            {
                TypeDoc = "388";
            }
            else if (cboTypeDoc.SelectedIndex == 2)
            {
                TypeDoc = "380";
            }
            else if (cboTypeDoc.SelectedIndex == 3)
            {
                TypeDoc = "T01";
            }
            else if (cboTypeDoc.SelectedIndex == 4)
            {
                TypeDoc = "80";
            }
            else if (cboTypeDoc.SelectedIndex == 1)
            {
                TypeDoc = "81";
            }
            else if (cboTypeDoc.SelectedIndex == -1)
            {
                lbUrl.Text = "";
                typeurl = "";
            }
        }




        private void input_copies_Leave(object sender, EventArgs e)
        {
            bool chet_copies = check_copies(input_copies.Text);
            Console.WriteLine(chet_copies);
            if (chet_copies == false)
            {
                MessageBox.Show("กรุณากรอก Copies ใหม่อีกครั้ง", "ข้อความแจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                input_copies.Focus();
            }
        }

        private void cboServiceCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboServiceCode.SelectedIndex == 2)
            {
                chkPreview.Visible = true;
            }
            else
            {
                chkPreview.Checked = false;
                chkPreview.Visible = false;
            }
        }


        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (txtAccessKey.PasswordChar.ToString() == "*")
            {
                txtAccessKey.PasswordChar = '\0';
                pictureBox4.Visible = true;
                pictureBox2.Visible = false;
            }
            else
            {
                txtAccessKey.PasswordChar = '*';
                pictureBox4.Visible = false;
                pictureBox2.Visible = true;
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            if (txtAccessKey.PasswordChar.ToString() == "*")
            {
                txtAccessKey.PasswordChar = '\0';
                pictureBox4.Visible = true;
                pictureBox2.Visible = false;
            }
            else
            {
                txtAccessKey.PasswordChar = '*';
                pictureBox4.Visible = false;
                pictureBox2.Visible = true;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void txtAmountFile_TextChanged(object sender, EventArgs e)
        {

        }
        public void meth__openfolder()
        {
            string[] arrPathIO = null;
            string pathfile = myTestKey + "\\" + inputnameFile.Text;
            if (inputnameFile.Text.Length != 0)
            {
                StreamReader txtw = new StreamReader(pathfile);
                if (txtw != null)
                {
                    String line = txtw.ReadToEnd();
                    byte[] data = Convert.FromBase64String(line);
                    using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                    {
                        byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                        using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                        {
                            ICryptoTransform transform = tripleDes.CreateDecryptor();
                            byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                            //Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                            line = UTF8Encoding.UTF8.GetString(results);
                        }
                    }

                    txtw.Close();
                    arrPathIO = line.Split(',');
                    int lenOutput = arrPathIO[1].Split('\\').Length;
                    string folderPath = (arrPathIO[1]);
                    Process.Start("explorer.exe", folderPath);


                }
            }
        }
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            meth__openfolder();
        }

        private void label21_Click(object sender, EventArgs e)
        {
            meth__openfolder();
        }

        private void Menu_Click(object sender, EventArgs e)
        {
            etaxOneth __etaxone = new etaxOneth();
            __etaxone.Show();
        }

        private void input_copies_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void input_copies_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(input_copies.Text, "  ^ [0-9]"))
            {
                input_copies.Text = "";
            }
        }

        private void check___copies_CheckedChanged(object sender, EventArgs e)
        {
            if (check___copies.Checked)
            {
                input_copies.Visible = false;
                label19.Visible = false;
            }
            else
            {
                input_copies.Visible = true;
                label19.Visible = true;
            }
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            etaxOneth __newform = new etaxOneth();
            __newform.Show();
        }

        private void eTAXONETHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://etax.one.th");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                cancelProcress("CloseAndCancel");
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

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            this.selectConfigFile();
        }

        private void label22_Click(object sender, EventArgs e)
        {
            this.selectConfigFile();
        }

        public void selectConfigFile()
        {
            Microsoft.Win32.RegistryKey rkey;
            rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
            myTestKey = (string)rkey.GetValue("PathConfigETAX");

            int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
            if (fileCount == 0)
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("Not Found FileConfig..",
                "Select Config",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Error, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
            }
            else
            {
                string[] textfile = Directory.GetFiles(myTestKey, "*.cfg").Select(Path.GetFileName).ToArray();
                InputBox.SetLanguage(InputBox.Language.English);

                DialogResult res = InputBox.ShowDialog("Select FileConfig :",
                "Select Config",   //Text message (mandatory), Title (optional)
                    InputBox.Icon.Question, //Set icon type (default info)
                    InputBox.Buttons.YesNo, //Set buttons (default ok)
                    InputBox.Type.ComboBox, //Set type (default nothing)
                    textfile, //String field as ComboBox items (default null)
                    true, //Set visible in taskbar (default false)
                    new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                if (res == System.Windows.Forms.DialogResult.Yes)
                {
                    StreamReader sr = new StreamReader(myTestKey + "\\" + InputBox.ResultValue);
                    String line = sr.ReadToEnd();
                    byte[] data = Convert.FromBase64String(line);
                    using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                    {
                        byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                        using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                        {
                            ICryptoTransform transform = tripleDes.CreateDecryptor();
                            byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                            Console.WriteLine(UTF8Encoding.UTF8.GetString(results));
                            line = UTF8Encoding.UTF8.GetString(results);
                        }
                    }
                    Console.WriteLine(line);
                    int len = line.ToString().Split(',').Length;
                    txtSellerTaxID.Text = line.Split(',')[2];
                    txtBranchID.Text = line.Split(',')[3];
                    txtAPIKey.Text = APIKEY;
                    txtUserCode.Text = line.Split(',')[5];
                    txtAccessKey.Text = line.Split(',')[6];
                    cboServiceCode.SelectedItem = line.Split(',')[7];
                    //cboServiceURL.SelectedItem = line.Split(',')[8];
                    if (line.Split(',')[8] == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
                    {
                        cboServiceURL.SelectedIndex = 0;
                    }
                    else if (line.Split(',')[8] == "https://etaxsp.one.th/etaxdocumentws/etaxsigndocument")
                    {
                        cboServiceURL.SelectedIndex = 1;
                    }

                    sr.Close();
                }

            }
            rkey.Close();
        }

        public string RandomString(int size, bool lowerCase)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }
            if (lowerCase)
                return builder.ToString().ToLower();
            return builder.ToString();
        }
        
    }
    public class NetConfig_{
        public bool internet_status_()
        {
            bool con = NetworkInterface.GetIsNetworkAvailable();
            if (con == true)
            {
                return true;
            }
            else
            {
                return false;
                
            }
        }
    }
}
