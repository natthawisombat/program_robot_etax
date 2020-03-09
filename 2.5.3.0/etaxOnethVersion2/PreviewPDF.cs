using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace etaxOnethVersion2
{
    public partial class PreviewPDF : Form
    {
        int TogMove;
        int MValX;
        int MValY;
        public string PathPreviewPdf;
        public PreviewPDF()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PreviewPDF_Load(object sender, EventArgs e)
        {
            axAcroPDF1.src = PathPreviewPdf;
            //this.pdfDocumentViewer1.LoadFromFile(PathPreviewPdf);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
           
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
           
        }

        private void pbClose_Click(object sender, EventArgs e)
        {
            this.Close();
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

        private void pbMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
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
    }
}
