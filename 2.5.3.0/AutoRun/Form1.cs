using etaxOnethVersion2;
using Ookii.Dialogs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Collections;
using MsgBox;
using Microsoft.Win32.TaskScheduler;
using System.Security.Cryptography;
using System.Security.Principal;
using Timer = System.Windows.Forms.Timer;
using ETAXStartup.Function;
using System.Diagnostics;

namespace ETAXStartup
{
    public partial class StartUp : Form
    {
        public string myTestKey;
        etaxOneth etax = new etaxOneth();
        string hash = "0105561072420_00000_Etax_One_th";
        public string strPathConfig = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName  + "\\FolderConfig";
        ArrayList listnameFile = new ArrayList();
        public ArrayList listnameWorker = new ArrayList();
        private System.Windows.Forms.Timer timer1;

        public StartUp()
        {            
            InitializeComponent();
            this.InitTimer();
        }        
        public void InitTimer()
        {
            timer1 = new Timer();
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Interval = 1000; // in miliseconds
            timer1.Start();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i <= GC.MaxGeneration; i++)
            {
                int count = GC.CollectionCount(i);
                GC.Collect();
            }
            GC.WaitForPendingFinalizers();
            GC.SuppressFinalize(this);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                etaxOneth etax = new etaxOneth();
                etax.Show();
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

        private void pbClose_Click(object sender, EventArgs e)
        {
            
            notifyIcon1.Dispose();
            Application.Exit();

        }

        private void pbMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Close();
                notifyIcon1.Visible = true;
                
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                etaxOneth etax = new etaxOneth();
                etax.Show();
                notifyIcon1.Visible = true;
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

        private void pbRestoreDown_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            pbRestoreDown.Visible = false;
            pbMaximize.Visible = true;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int countform1 = Application.OpenForms.OfType<etaxOneth>().Count();
            if(countform1 > 0)
            {
                MessageBox.Show("กรุณาปิดทีละหน้าต่าง !","คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                Application.Exit();
            }
            //etax.closeSocket();
            
        }

        private void pbPowerOff_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void setstartUp()
        {

            if (!Directory.Exists(strPathConfig))
            {
                try
                {
                    Directory.CreateDirectory(strPathConfig);

                    Microsoft.Win32.RegistryKey key;
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);

                    string app_path = System.Reflection.Assembly.GetEntryAssembly().Location;
                    string app_name = System.Reflection.Assembly.GetEntryAssembly().ManifestModule.Name;
                    if (!key.GetValueNames().Contains(app_name))
                    {
                        key.SetValue(app_name, app_path);
                        key.Close();
                    }
                    else
                    {
                        key.DeleteValue(app_name);
                        key.SetValue(app_name, app_path);
                        key.Close();
                    }


                    Microsoft.Win32.RegistryKey rkey;
                    rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
                    if (rkey == null)
                    {
                        rkey.SetValue("PathConfigETAX", strPathConfig);

                    }
                    else
                    {
                        rkey.SetValue("PathConfigETAX", strPathConfig);

                    }

                    myTestKey = (string)rkey.GetValue("PathConfigETAX");
                    this.WindowState = FormWindowState.Minimized;
                    this.ShowInTaskbar = false;
                    int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
                    if (fileCount == 0)
                    {
                        etax = new etaxOneth();
                        etax.Show();
                        notifyIcon1.ShowBalloonTip(100, "แจ้งเตือน", "โปรแกรมพร้อมทำงาน..", ToolTipIcon.Info);
                    }
                    else
                    {
                        string[] textfile = Directory.GetFiles(myTestKey, "*.cfg").Select(Path.GetFileName).ToArray();


                        for (int i = 0; i < fileCount; i++)
                        {
                            StreamReader sr = new StreamReader(myTestKey + "\\" + textfile[i]);
                            String line = sr.ReadToEnd();
                            byte[] data = Convert.FromBase64String(line);
                            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                            {
                                byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                                using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                                {
                                    ICryptoTransform transform = tripleDes.CreateDecryptor();
                                    byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                                    line = UTF8Encoding.UTF8.GetString(results);
                                }
                            }
                            etax = new etaxOneth();
                            int len = line.ToString().Split(',').Length;
                            bool checkst = bool.Parse(line.ToString().Split(',')[15]);
                            if (checkst == true)
                            {
                                etax.linkpath = (myTestKey + "\\" + textfile[i]).ToString();
                                etax.nameFile = textfile[i].ToString();
                                etax.textinfile = line.ToString();
                                etax.Show();
                                sr.Close();
                                notifyIcon1.ShowBalloonTip(100, "แจ้งเตือน", "โปรแกรมเริ่มการทำงาน..", ToolTipIcon.Info);
                            }
                            else
                            {


                                sr.Close();

                            }

                        }
                        
                        etax.Show();
                        
                       
                    }
                    int countform = Application.OpenForms.OfType<etaxOneth>().Count();
                    rkey.Close();
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
                }

                

        }
            else
            {

                try
                {
                    Microsoft.Win32.RegistryKey key;
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);

                    string app_path = System.Reflection.Assembly.GetEntryAssembly().Location;
                    string app_name = System.Reflection.Assembly.GetEntryAssembly().ManifestModule.Name;

                    if (!key.GetValueNames().Contains(app_name))
                    {
                        key.SetValue(app_name, app_path);
                        key.Close();
                    }
                    else
                    {
                        key.DeleteValue(app_name);
                        key.SetValue(app_name, app_path);
                        key.Close();
                    }



                    Microsoft.Win32.RegistryKey rkey;
                    rkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\ETAX\\Run");
                    if (rkey == null)
                    {
                        rkey.SetValue("PathConfigETAX", strPathConfig);

                    }
                    else
                    {
                        rkey.SetValue("PathConfigETAX", strPathConfig);

                    }

                    myTestKey = (string)rkey.GetValue("PathConfigETAX");
                    this.WindowState = FormWindowState.Minimized;
                    this.ShowInTaskbar = false;
                    int fileCount = Directory.GetFiles(myTestKey, "*.cfg", SearchOption.AllDirectories).Length;
                    if (fileCount == 0)
                    {
                        etax = new etaxOneth();
                        try
                        {
                            etax.Show();
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
                        notifyIcon1.ShowBalloonTip(100, "แจ้งเตือน", "โปรแกรมพร้อมทำงาน..", ToolTipIcon.Info);
                    }
                    else
                    {
                        string[] textfile = Directory.GetFiles(myTestKey, "*.cfg").Select(Path.GetFileName).ToArray();
                        ManageProgram Manage = new ManageProgram();
                        bool Openprogram = true;
                        Process[] process = Manage.Callprocess("ETAX-One Electronic Billing");
                        if(Manage.CheckProcess(process))
                        {
                            DialogResult dialogResult = MessageBox.Show("โปรแกรมกำลังทำงานอยู่ในส่วนของ Background Process คุณต้องการเปิดเป็น Foreground Process หรือไม่ (อาจมีผลกระทบต่อการทำงาน โปรดตวรจสอบว่าระบบไม่ได้กำลังสร้างเอกสารอยู่)", "คำเตือน",MessageBoxButtons.YesNo);
                            //MessageBox.Show("โปรแกรมกำลังทำงานอยู่ในส่วนของ Background Process คุณต้องการเปิดเป็น Foreground Process หรือไม่ (อาจมีผลกระทบต่อการทำงาน)");
                            if (dialogResult == DialogResult.Yes)
                            {
                                Manage.KillProcess(process);
                                Openprogram = true;
                            }
                            else if (dialogResult == DialogResult.No)
                            {
                                Openprogram = false;
                                Application.Exit();
                            }
                        }
                        else
                        {
                            Openprogram = true;
                        }
                        if (Openprogram)
                        {
                            for (int i = 0; i < fileCount; i++)
                            {
                                StreamReader sr = new StreamReader(myTestKey + "\\" + textfile[i]);
                                String line = sr.ReadToEnd();
                                byte[] data = Convert.FromBase64String(line);
                                using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
                                {
                                    byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                                    using (TripleDESCryptoServiceProvider tripleDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                                    {
                                        ICryptoTransform transform = tripleDes.CreateDecryptor();
                                        byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                                        line = UTF8Encoding.UTF8.GetString(results);
                                    }
                                }
                                etax = new etaxOneth();
                                int len = line.ToString().Split(',').Length;
                                bool checkst = bool.Parse(line.ToString().Split(',')[15]);
                                if (checkst == true)
                                {

                                    etax.linkpath = (myTestKey + "\\" + textfile[i]).ToString();
                                    etax.nameFile = textfile[i].ToString();
                                    etax.textinfile = line.ToString();
                                    etax.Show();
                                    sr.Close();
                                    notifyIcon1.ShowBalloonTip(100, "แจ้งเตือน", "โปรแกรมเริ่มการทำงาน..", ToolTipIcon.Info);
                                }
                                else
                                {

                                    sr.Close();

                                }

                            }
                            etax.Show();
                        }
                        
                       

                    }
                    int countform = Application.OpenForms.OfType<etaxOneth>().Count();
                    etax.countform = countform;
                    rkey.Close();
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
                }
                

            }
           

    

}

        private void Form1_Load(object sender, EventArgs e)
        {
            setstartUp();
            //this.InitTimer();
        }

        

        private void helptoIcon(object sender, FormClosingEventArgs e)
        {
            notifyIcon1.Icon = null;
            notifyIcon1.Dispose();
        }

       

        
    }
    class time_agian
    {
        public void start()
        {
            for (int i = 0; i <= GC.MaxGeneration; i++)
            {
                int count = GC.CollectionCount(i);
                Console.WriteLine(count);
                GC.Collect();
            }
            GC.SuppressFinalize(this);
        }
        
    }
}
