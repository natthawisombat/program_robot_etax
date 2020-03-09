
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETAXStartup.Function
{
    class ManageProgram
    {
        Process[] processlist;

        public Process[] Callprocess(String name)
        {
            processlist = Process.GetProcessesByName(name);
            foreach (Process i in processlist)
            {
                Console.WriteLine("data => " + i.Id);
            }
            return processlist;
        }
        public bool CheckProcess(Process[] process)
        {
            if (process.Length == 2)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool KillProcess(Process[] process)
        {
            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                string pid = currentProcess.Id.ToString();
                foreach(Process id in process)
                {
                    if (id.Id.ToString() != pid)
                    {
                        id.Kill();
                    }
                }
                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error ==> " + ex.Message);
                return false;
            }
            
        }
    }
}
