using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace etaxOnethVersion2
{
    public partial class ProgressDlg : Form
    {
        public ProgressDlg()
        {
            InitializeComponent();
        }
        public int ProgressValue
        {
            set { this.progressBar1.Value = value; }
        }
    }
}
