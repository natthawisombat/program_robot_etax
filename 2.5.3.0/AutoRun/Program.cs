using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using MsgBox;
using Microsoft.Win32.TaskScheduler;


namespace ETAXStartup
{
    static class Program
    {
        static Mutex m;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool first = false;
            m = new Mutex(true, Application.ProductName.ToString(), out first);
            
            if ((first))
            {
                
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                Application.Run(new StartUp());
                m.ReleaseMutex();
              
            }
            else
            {
                InputBox.SetLanguage(InputBox.Language.English);
                InputBox.ShowDialog("Application" + " " + Application.ProductName.ToString() + " " + "already running",
                "Open ETAX-One Electronic Billing",   //Text message (mandatory), Title (optional)
                InputBox.Icon.Information, //Set icon type (default info)
                InputBox.Buttons.Ok, //Set buttons (default ok)
                InputBox.Type.Nothing, //Set type (default nothing)
                null, //String field as ComboBox items (default null)
                true, //Set visible in taskbar (default false)
                new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold)); //Set font (default by system)
                //MessageBox.Show("Application" + " " + Application.ProductName.ToString() + " " + "already running");
            }

        }
      


    }
}
