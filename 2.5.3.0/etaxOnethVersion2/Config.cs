using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ookii.Dialogs;
using System.Windows.Forms;
using System.IO;
using MsgBox;
using System.Text.RegularExpressions;
using System.Security.Cryptography;



namespace etaxOnethVersion2
{
    public partial class Config : Form
    {
        int TogMove;
        int MValX;
        int MValY;
        string APIKEY = "AK2-3UY8R84Q6GFXD8QXZMFYNPUJRFOKWOT614C5GAAC88OVAGLD8F1HWAPB9LW05QQSACKVSN0FBI1H8WPO0NWU29MWO9854BBMW5OJ6IZOBUJFZRHV0ZPD4CP02PXLT95YHD6QX3TH01CZX7TJ7X4NBLKZLGQP8NF3BKZIU6NSW463VL61A8LBBAKIJSQS7M1TL2S50E6AGWRE85ZSK6AIHYDYXY7C19LKLFRFWXW3FAOAB61O0E5TGPBNC3P4X72BU";
        string Kind_Work;
        public string url____ { get; set; }
        public string strPathConfig = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName + "\\FolderConfig";
        public string namefileconfig;
        public string myTestKey;
        public string[] arrnameFiletoCombo;
        string checkemail;
        string hash = "0105561072420_00000_Etax_One_th";
        public string nameFiletext;
        public Config()
        {
            InitializeComponent();
        }
        private void pbClose_Click(object sender, EventArgs e)
        {
            this.Close();
            //Application.Exit();
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

        private void btnHome_Click(object sender, EventArgs e)
        {
            etaxOneth etaxForm = new etaxOneth();
            this.Close();
            etaxForm.Show();
        }

        private void pbRestoreDown_Click_1(object sender, EventArgs e)
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

        private void pbMinimize_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pbPowerOff_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            txtNameFile.Clear();
            txtInput.Clear();
            txtOutput.Clear();
            txtSellerTaxID.Clear();
            txtBranchID.Clear();
            //txtAPIKey.Clear();
            txtEmail.Clear();
            txtUserCode.Clear();
            txtConfixExcel.Clear();
            txtAccessKey.Clear();
            cboServiceCode.SelectedIndex = -1;
            cboServiceURL.SelectedIndex = -1;
            checkbtnset.Checked = false;
            Kind_Work = "New";
            txtAmountFile.Clear();
            txtTimeRun.Clear();
            InputBox.SetLanguage(InputBox.Language.English);
            InputBox.ShowDialog("Can Fill to Config!",
            "New Config",   //Text message (mandatory), Title (optional)
            InputBox.Icon.Information, //Set icon type (default info)
            InputBox.Buttons.Ok, //Set buttons (default ok)
            InputBox.Type.Nothing, //Set type (default nothing)
            null, //String field as ComboBox items (default null)
            true, //Set visible in taskbar (default false)
            new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
            //VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog();

            //if (dlg.ShowDialog() == DialogResult.OK)
            //{
            //    txtNameFile.Text = dlg.SelectedPath;
            //    Kind_Work = "New";
            //    //string strDataPath = Path.GetDirectoryName(txtInput.Text) + "," + txtOutput.Text;
            //    //File.WriteAllText(strPathInOut, String.Empty);
            //    //CreateTextFile(strPathInOut, strDataPath);
            //}
        }

        private void btnInput_Click(object sender, EventArgs e)
        {
            
            VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtInput.Text = dlg.SelectedPath;

                //string strDataPath = Path.GetDirectoryName(txtInput.Text) + "," + txtOutput.Text;
                //File.WriteAllText(strPathInOut, String.Empty);
                //CreateTextFile(strPathInOut, strDataPath);
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtOutput.Text = dlg.SelectedPath;

                //string strDataPath = Path.GetDirectoryName(txtInput.Text) + "," + txtOutput.Text;
                //File.WriteAllText(strPathInOut, String.Empty);
                //CreateTextFile(strPathInOut, strDataPath);
            }
        }

        public void chkServiceURL()
        {
            if (cboServiceURL.SelectedIndex.ToString() == "0")
            {
                this.APIKEY = "AK2-3UY8R84Q6L5KAOLE781044GRX2LZTJ95WHJDZWOORWXVOR92XUJFDTJ02FHTFVYFLQXF6QAHRNGJHSUNXGFZT65EBMHNRUO7J0IARSEBFMI8BN6OHT1M0BJGBAF4PB0HMMRTALW8QD4LCNDRNSXJOFNR5KWQOEPS7EKBVZLTSBJBGZ3676YJXLP96GNZ36CFC00HIVZHGO3DHO14ZEIEZN8AY99VKU0WR1QEM3DW3YZQE11OZ2UINCGIQN5ILDAKJ";
            }
            else if(cboServiceURL.SelectedIndex.ToString() == "1")
            {
                this.APIKEY = "AK2-3UY8R84Q6GFXD8QXZMFYNPUJRFOKWOT614C5GAAC88OVAGLD8F1HWAPB9LW05QQSACKVSN0FBI1H8WPO0NWU29MWO9854BBMW5OJ6IZOBUJFZRHV0ZPD4CP02PXLT95YHD6QX3TH01CZX7TJ7X4NBLKZLGQP8NF3BKZIU6NSW463VL61A8LBBAKIJSQS7M1TL2S50E6AGWRE85ZSK6AIHYDYXY7C19LKLFRFWXW3FAOAB61O0E5TGPBNC3P4X72BU";
            }
        }
        
        private void btnSave_Click(object sender, EventArgs e)
        {
            bool checkfill = CheckAddDataAuto();
            string strDataPath, strPathInOut;
            byte[] results;
            if (checkfill == true)
            {
                switch (Kind_Work)
                {
                    case "New":
                        chkServiceURL();
                        strDataPath = txtInput.Text + "," + txtOutput.Text + "," + txtSellerTaxID.Text + "," + txtBranchID.Text + "," + APIKEY + "," + txtUserCode.Text + "," + txtAccessKey.Text + "," + cboServiceCode.SelectedItem + "," + this.url____ + "," + txtAmountFile.Text + "," + txtTimeRun.Text + "," + txtEmail.Text + "," + txtConfixExcel.Text + "," + checkAutoPrint.Checked + "," + cboPrinter.SelectedItem + "," + checkbtnset.Checked + "," + chk__copiesfile.Checked + "," + txt_copies.Text;
                        byte[] data = UTF8Encoding.UTF8.GetBytes(strDataPath);
                        using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                        {
                            byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                            using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                            {
                                ICryptoTransform transform = tripleDes.CreateEncryptor();
                                results = transform.TransformFinalBlock(data, 0, data.Length);

                            }
                        }
                        strPathInOut = strPathConfig + "\\" + txtNameFile.Text + "_Config_" + DateTime.Now.ToString("dd-M-yy_HH-mm-ss") + ".cfg";
                        CreateTextFile(strPathInOut, Convert.ToBase64String(results, 0, results.Length), "New");
                        break;
                    case "Edit":
                        //MessageBox.Show(strPathConfig + "\\" + txtNameFile.Text);
                        chkServiceURL();
                        strDataPath = txtInput.Text + "," + txtOutput.Text + "," + txtSellerTaxID.Text + "," + txtBranchID.Text + "," + APIKEY + "," + txtUserCode.Text + "," + txtAccessKey.Text + "," + cboServiceCode.SelectedItem + "," + this.url____  + "," + txtAmountFile.Text + "," + txtTimeRun.Text + "," + txtEmail.Text + "," + txtConfixExcel.Text + "," + checkAutoPrint.Checked + "," + cboPrinter.SelectedItem + "," + checkbtnset.Checked + "," + chk__copiesfile.Checked + "," + txt_copies.Text;
                        byte[] data2 = UTF8Encoding.UTF8.GetBytes(strDataPath);
                        using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                        {
                            byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                            using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                            {
                                ICryptoTransform transform = tripleDes.CreateEncryptor();
                                results = transform.TransformFinalBlock(data2, 0, data2.Length);
                            }
                        }
                        strPathInOut = strPathConfig;
                        CreateTextFile(strPathInOut, Convert.ToBase64String(results, 0, results.Length), "Edit");
                        break;
                }
            }
        }
        public bool CheckAddDataAuto()
        {
            if (txtSellerTaxID.Text.Equals("") || txtBranchID.Text.Equals("") || txtNameFile.Text.Equals("") || txtInput.Text.Equals("") ||
                txtOutput.Text.Equals("") || txtUserCode.Text.Equals("") || txtAccessKey.Text.Equals("") || cboServiceCode.Equals("") || this.url____.Equals("") || txtAmountFile.Text.Equals("") || txtTimeRun.Text.Equals(""))
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
            else
            {
                if (checkAutoPrint.Checked == true )
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
                    else if (txt_copies.Text == "0" || txt_copies.Text == "00" || txt_copies.Text.Length == 0)
                    {
                        InputBox.SetLanguage(InputBox.Language.English);
                        InputBox.ShowDialog("กรุณาตรวจสอบ Copies หากต้องการ Auto Print",
                        "Warning",   //Text message (mandatory), Title (optional)
                        InputBox.Icon.Information, //Set icon type (default info)
                        InputBox.Buttons.Ok, //Set buttons (default ok)
                        InputBox.Type.Nothing, //Set type (default nothing)
                        null, //String field as ComboBox items (default null)
                        true, //Set visible in taskbar (default false)
                        new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                        txt_copies.Focus();
                        return false;
                    }
                }
                
            }
            return true;

        }
        public void CreateTextFile(string strPath, string strData,string substatus)
        {
            if(substatus == "New")
            {
                try
                {
                    
                    TextWriter txtw = new StreamWriter(strPath);
                    txtw.Write(strData);
                    txtw.Close();
                    InputBox.SetLanguage(InputBox.Language.English);
                    InputBox.ShowDialog("Save Config Success!",
                    "New Config",   //Text message (mandatory), Title (optional)
                    InputBox.Icon.Information, //Set icon type (default info)
                    InputBox.Buttons.Ok, //Set buttons (default ok)
                    InputBox.Type.Nothing, //Set type (default nothing)
                    null, //String field as ComboBox items (default null)
                    true, //Set visible in taskbar (default false)
                    new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                    refield();
                }
                catch (Exception e)
                {
                    
                }
            }
            else if(substatus == "Edit")
            {
                try
                {
                    File.Delete(strPath + "\\" + nameFiletext);
                }
                catch (Exception e)
                {

                }
                    
                    TextWriter txtw = new StreamWriter(strPath + "\\" + txtNameFile.Text + "_Config_" + DateTime.Now.ToString("dd-M-yy_HH-mm-ss") + ".cfg");
                    txtw.Write(strData);
                    txtw.Close();
                    InputBox.SetLanguage(InputBox.Language.English);
                    InputBox.ShowDialog("Save Config Success!",
                    "Edit Config",
                    InputBox.Icon.Information,
                    InputBox.Buttons.Ok,
                    InputBox.Type.Nothing,
                    null,
                    true,
                    new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold));
                refield();


            }
           
        }

        public void refield()
        {
            txtAPIKey.Text = APIKEY;
            txtBranchID.Text = "";
            txtInput.Text = "";
            txtOutput.Text = "";
            txtNameFile.Text = "";
            txtSellerTaxID.Text = "";
            txtEmail.Text = "";
            txtUserCode.Clear();
            txtAccessKey.Clear();
            cboServiceCode.SelectedIndex = -1;
            cboServiceURL.SelectedIndex = -1;
            txtAmountFile.Clear();
            txtConfixExcel.Text = "";
            txtTimeRun.Clear();
            checkbtnset.Checked = false;
            //txtAPIKey.Refresh();
            txtBranchID.Refresh();
            txtInput.Refresh();
            txtOutput.Refresh();
            txtNameFile.Refresh();
            txtSellerTaxID.Refresh();
            txtEmail.Refresh();
            txtUserCode.Refresh();
            txtAccessKey.Refresh();
            cboServiceCode.Refresh();
            cboServiceURL.Refresh();
            checkbtnset.Refresh();
            txtAmountFile.Refresh();
            txtTimeRun.Refresh();
            txtNameFile.Enabled = true;
            checkAutoPrint.Checked = false;
            cboPrinter.Items.Clear();
        }
        private void btnEdit_Click(object sender, EventArgs e)
        {
            Microsoft.Win32.RegistryKey rkey;
            rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
            myTestKey = (string)rkey.GetValue("PathConfigETAX");

            int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
            if(fileCount == 0)
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("Not Found FileConfig..",
                "Edit Config",   //Text message (mandatory), Title (optional)
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
                "Edit Config",   //Text message (mandatory), Title (optional)
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
                    int len = line.ToString().Split(',').Length;
                    txtInput.Text = line.Split(',')[0];
                    txtOutput.Text = line.Split(',')[1];
                    txtSellerTaxID.Text = line.Split(',')[2];
                    txtBranchID.Text = line.Split(',')[3];
                    txtAPIKey.Text = APIKEY;
                    txtUserCode.Text = line.Split(',')[5];
                    txtAccessKey.Text = line.Split(',')[6];
                    cboServiceCode.SelectedItem = line.Split(',')[7];
                    //cboServiceURL.SelectedItem = line.Split(',')[8];
                    if(line.Split(',')[8] == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
                    {
                        cboServiceURL.SelectedIndex = 0;
                    }else if(line.Split(',')[8] == "https://etaxsp.one.th/etaxdocumentws/etaxsigndocument")
                    {
                        cboServiceURL.SelectedIndex = 1;
                    }
                    txtAmountFile.Text = line.Split(',')[9];
                    txtTimeRun.Text = line.Split(',')[10];
                    txtEmail.Text = line.Split(',')[11];
                    txtConfixExcel.Text = line.Split(',')[12];
                    try
                    {
                        checkAutoPrint.Checked = bool.Parse(line.Split(',')[13]);
                        cboPrinter.SelectedItem = line.Split(',')[14];
                    }
                    catch (IndexOutOfRangeException ex)
                    {
                        checkAutoPrint.Checked = false;
                        cboPrinter.SelectedIndex = -1;
                    }
                    try
                    {
                        chk__copiesfile.Checked = bool.Parse(line.Split(',')[16]);
                    }
                    catch (Exception ea)
                    {
                        chk__copiesfile.Checked = false;
                    }
                    try
                    {
                        txt_copies.Text = line.Split(',')[17];
                    }
                    catch(Exception ex)
                    {
                        txt_copies.Text = "1";
                    }
                    checkbtnset.Checked = bool.Parse(line.Split(',')[15]);
                    
                    txtNameFile.Text = InputBox.ResultValue.Split('_')[0];
                    nameFiletext = InputBox.ResultValue;
                    Kind_Work = "Edit";
                    sr.Close();
                }
                
            }
            rkey.Close();
            //Set buttons language Czech/ English / German / Slovakian / Spanish(default English)
            //OpenFileDialog oFileDialog = new OpenFileDialog();
            //string strPathIO = string.Empty;
            //string[] arrPathIO = null;
            //if (oFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    Kind_Work = "Edit";
            //    txtNameFile.Text = oFileDialog.FileName;
            //    txtNameFile.Enabled = false;
            //    strPathIO = System.IO.File.ReadAllText(txtNameFile.Text);
            //    arrPathIO = strPathIO.Split(',');
            //    txtInput.Text = arrPathIO[0];
            //    txtOutput.Text = arrPathIO[1];
            //    txtSellerTaxID.Text = arrPathIO[2];
            //    txtBranchID.Text = arrPathIO[3];
            //    txtAPIKey.Text = arrPathIO[4];
            //}
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            Microsoft.Win32.RegistryKey rkey;
            rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
            myTestKey = (string)rkey.GetValue("PathConfigETAX");

            int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
            if (fileCount == 0)
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("Not Found FileConfig..",
                "Delete Config",   //Text message (mandatory), Title (optional)
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

                DialogResult resdel = InputBox.ShowDialog("Select FileConfig :",
                "Delete Config",   //Text message (mandatory), Title (optional)
                    InputBox.Icon.Information, //Set icon type (default info)
                    InputBox.Buttons.OkCancel, //Set buttons (default ok)
                    InputBox.Type.ComboBox, //Set type (default nothing)
                    textfile, //String field as ComboBox items (default null)
                    true, //Set visible in taskbar (default false)
                    new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                if (resdel == System.Windows.Forms.DialogResult.OK || resdel == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        txtNameFile.Clear();
                        txtInput.Clear();
                        txtOutput.Clear();
                        txtSellerTaxID.Clear();
                        txtBranchID.Clear();
                        //txtAPIKey.Clear();
                        txtUserCode.Clear();
                        txtAccessKey.Clear();
                        cboServiceCode.SelectedIndex = -1;
                        cboServiceURL.SelectedIndex = -1;
                        txtEmail.Clear();
                        txtAmountFile.Clear();
                        txtTimeRun.Clear();
                        txtConfixExcel.Clear();
                        checkbtnset.Checked = false;
                        checkAutoPrint.Checked = false;
                        cboPrinter.SelectedIndex = -1;
                        System.IO.File.Delete(myTestKey + "\\" + InputBox.ResultValue);
                        InputBox.SetLanguage(InputBox.Language.English);
                        InputBox.ShowDialog("Delete Config Success!",
                        "Delete Config",
                        InputBox.Icon.Information,
                        InputBox.Buttons.Ok,
                        InputBox.Type.Nothing,
                        null,
                        true,
                        new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold));
                    }
                    catch (Exception eaa)
                    {
                        InputBox.SetLanguage(InputBox.Language.English);
                        InputBox.ShowDialog("Not Found Config..",
                        "Delete Config",
                        InputBox.Icon.Error,
                        InputBox.Buttons.Ok,
                        InputBox.Type.Nothing,
                        null,
                        true,
                        new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold));
                    }
                    

                }
            }


            rkey.Close();
            //OpenFileDialog oFileDialog = new OpenFileDialog();

            //if (oFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    System.IO.File.Delete(oFileDialog.FileName);
            //    MessageBox.Show("Delete Success!!!");
            //}
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Config configform = new Config();
            this.Close();
            configform.Show();
        }

        private void Config_Load(object sender, EventArgs e)
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
            Kind_Work = "New";
        }

        private void txtMail_Leave(object sender, EventArgs e)
        {
            string pattern = "^([0-9a-zA-Z]([-\\.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-z-A-Z]{2,9})$";
            if (Regex.IsMatch(txtEmail.Text, pattern))
            {
                checkemail = "EMAILOK";
            }
            else
            {
                checkemail = "FAILMAIL";
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("โปรดตรวจสอบ Email ให้ถูกต้อง..",
                "Warning",
                InputBox.Icon.Error,
                InputBox.Buttons.Ok,
                InputBox.Type.Nothing,
                null,
                true,
                new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular));
            }
        }

        private void checknumber(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFileDialog = new OpenFileDialog();
            if (oFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtConfixExcel.Text = oFileDialog.FileName;
            }
        }

        private void checkbtnset_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkAutoPrint_CheckedChanged(object sender, EventArgs e)
        {
            cboPrinter.Items.Clear();
            if (checkAutoPrint.Checked)
            {
                labelAutoPrint.Visible = true;
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    cboPrinter.Items.Add(printer);
                }
                cboPrinter.Visible = true;
                //label20.Visible = true;
                //txt_copies.Visible = true;
                chk__copiesfile.Visible = true;
                chk__copiesfile.Checked = false;
                if (chk__copiesfile.Checked)
                {
                    label20.Visible = false;
                    txt_copies.Visible = false;
                }
                else
                {
                    label20.Visible = true;
                    txt_copies.Visible = true;
                    txt_copies.Text = "1";
                }
            }
            else
            {
                chk__copiesfile.Checked = false;
                if (chk__copiesfile.Checked)
                {
                    label20.Visible = false;
                    txt_copies.Visible = false;
                }
                else
                {
                    label20.Visible = false;
                    txt_copies.Visible = false;
                }
                labelAutoPrint.Visible = false;
                cboPrinter.Visible = false;
                cboPrinter.SelectedItem = "";                
                chk__copiesfile.Visible = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txt_copies.Text, "  ^ [0-9]"))
            {
                txt_copies.Text = "";
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void cboServiceURL_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cboServiceURL.SelectedIndex == 0)
            {
                this.url____ = "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument";
            }else if(cboServiceURL.SelectedIndex == 1)
            {
                this.url____ = "https://etaxsp.one.th/etaxdocumentws/etaxsigndocument";
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

        private void chk__copiesfile_CheckedChanged(object sender, EventArgs e)
        {
            if (chk__copiesfile.Checked)
            {
                label20.Visible = false;
                txt_copies.Visible = false;
            }
            else
            {
                label20.Visible = true;
                txt_copies.Visible = true;
            }
        }
        
        private void label12_ClientSizeChanged(object sender, EventArgs e)
        {

        }
    }
}
