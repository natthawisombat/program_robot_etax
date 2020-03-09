namespace etaxOneth_Process
{
    partial class etaxOnethProcess
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(etaxOnethProcess));
            this.pbRestoreDown = new System.Windows.Forms.PictureBox();
            this.pbMinimize = new System.Windows.Forms.PictureBox();
            this.pbMaximize = new System.Windows.Forms.PictureBox();
            this.pbClose = new System.Windows.Forms.PictureBox();
            this.pnHead = new System.Windows.Forms.Panel();
            this.lbName = new System.Windows.Forms.Label();
            this.pbLogo = new System.Windows.Forms.PictureBox();
            this.lbPercent = new System.Windows.Forms.Label();
            this.pgbLoad = new System.Windows.Forms.ProgressBar();
            this.lbHeadName = new System.Windows.Forms.Label();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.pnDisplay = new System.Windows.Forms.Panel();
            this.tt = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pbRestoreDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbMinimize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbMaximize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbClose)).BeginInit();
            this.pnHead.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbLogo)).BeginInit();
            this.pnDisplay.SuspendLayout();
            this.SuspendLayout();
            // 
            // pbRestoreDown
            // 
            this.pbRestoreDown.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pbRestoreDown.BackColor = System.Drawing.Color.Transparent;
            this.pbRestoreDown.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbRestoreDown.Image = ((System.Drawing.Image)(resources.GetObject("pbRestoreDown.Image")));
            this.pbRestoreDown.Location = new System.Drawing.Point(604, 1);
            this.pbRestoreDown.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pbRestoreDown.Name = "pbRestoreDown";
            this.pbRestoreDown.Size = new System.Drawing.Size(44, 41);
            this.pbRestoreDown.TabIndex = 7;
            this.pbRestoreDown.TabStop = false;
            this.tt.SetToolTip(this.pbRestoreDown, "Restore Down");
            this.pbRestoreDown.Visible = false;
            this.pbRestoreDown.Click += new System.EventHandler(this.pbRestoreDown_Click);
            // 
            // pbMinimize
            // 
            this.pbMinimize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pbMinimize.BackColor = System.Drawing.Color.Transparent;
            this.pbMinimize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbMinimize.Image = ((System.Drawing.Image)(resources.GetObject("pbMinimize.Image")));
            this.pbMinimize.Location = new System.Drawing.Point(552, 1);
            this.pbMinimize.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pbMinimize.Name = "pbMinimize";
            this.pbMinimize.Size = new System.Drawing.Size(44, 41);
            this.pbMinimize.TabIndex = 6;
            this.pbMinimize.TabStop = false;
            this.tt.SetToolTip(this.pbMinimize, "Minimize");
            this.pbMinimize.Click += new System.EventHandler(this.pbMinimize_Click);
            // 
            // pbMaximize
            // 
            this.pbMaximize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pbMaximize.BackColor = System.Drawing.Color.Transparent;
            this.pbMaximize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbMaximize.Image = ((System.Drawing.Image)(resources.GetObject("pbMaximize.Image")));
            this.pbMaximize.Location = new System.Drawing.Point(604, 1);
            this.pbMaximize.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pbMaximize.Name = "pbMaximize";
            this.pbMaximize.Size = new System.Drawing.Size(44, 41);
            this.pbMaximize.TabIndex = 5;
            this.pbMaximize.TabStop = false;
            this.tt.SetToolTip(this.pbMaximize, "Maximize");
            this.pbMaximize.Click += new System.EventHandler(this.pbMaximize_Click);
            // 
            // pbClose
            // 
            this.pbClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pbClose.BackColor = System.Drawing.Color.Transparent;
            this.pbClose.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbClose.Image = ((System.Drawing.Image)(resources.GetObject("pbClose.Image")));
            this.pbClose.Location = new System.Drawing.Point(656, 1);
            this.pbClose.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pbClose.Name = "pbClose";
            this.pbClose.Size = new System.Drawing.Size(41, 42);
            this.pbClose.TabIndex = 4;
            this.pbClose.TabStop = false;
            this.tt.SetToolTip(this.pbClose, "Close");
            this.pbClose.Click += new System.EventHandler(this.pbClose_Click);
            // 
            // pnHead
            // 
            this.pnHead.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnHead.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.pnHead.Controls.Add(this.lbName);
            this.pnHead.Controls.Add(this.pbLogo);
            this.pnHead.Controls.Add(this.pbRestoreDown);
            this.pnHead.Controls.Add(this.pbMinimize);
            this.pnHead.Controls.Add(this.pbMaximize);
            this.pnHead.Controls.Add(this.pbClose);
            this.pnHead.Location = new System.Drawing.Point(0, 0);
            this.pnHead.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pnHead.Name = "pnHead";
            this.pnHead.Size = new System.Drawing.Size(699, 44);
            this.pnHead.TabIndex = 3;
            this.pnHead.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pnHead_MouseDown);
            this.pnHead.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pnHead_MouseMove);
            this.pnHead.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pnHead_MouseUp);
            // 
            // lbName
            // 
            this.lbName.AutoSize = true;
            this.lbName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbName.ForeColor = System.Drawing.Color.White;
            this.lbName.Location = new System.Drawing.Point(67, 12);
            this.lbName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbName.Name = "lbName";
            this.lbName.Size = new System.Drawing.Size(218, 20);
            this.lbName.TabIndex = 5;
            this.lbName.Text = "etaxOne.th-Process V2.3";
            // 
            // pbLogo
            // 
            this.pbLogo.Image = ((System.Drawing.Image)(resources.GetObject("pbLogo.Image")));
            this.pbLogo.Location = new System.Drawing.Point(9, 9);
            this.pbLogo.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pbLogo.Name = "pbLogo";
            this.pbLogo.Size = new System.Drawing.Size(49, 27);
            this.pbLogo.TabIndex = 5;
            this.pbLogo.TabStop = false;
            // 
            // lbPercent
            // 
            this.lbPercent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lbPercent.AutoSize = true;
            this.lbPercent.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbPercent.Location = new System.Drawing.Point(37, 379);
            this.lbPercent.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbPercent.Name = "lbPercent";
            this.lbPercent.Size = new System.Drawing.Size(80, 20);
            this.lbPercent.TabIndex = 9;
            this.lbPercent.Text = "Percent:";
            this.tt.SetToolTip(this.lbPercent, "Percent Running");
            // 
            // pgbLoad
            // 
            this.pgbLoad.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pgbLoad.Location = new System.Drawing.Point(41, 347);
            this.pgbLoad.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pgbLoad.Name = "pgbLoad";
            this.pgbLoad.Size = new System.Drawing.Size(616, 28);
            this.pgbLoad.TabIndex = 8;
            this.tt.SetToolTip(this.pgbLoad, "Loading");
            // 
            // lbHeadName
            // 
            this.lbHeadName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.lbHeadName.AutoSize = true;
            this.lbHeadName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbHeadName.Location = new System.Drawing.Point(36, 22);
            this.lbHeadName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbHeadName.Name = "lbHeadName";
            this.lbHeadName.Size = new System.Drawing.Size(74, 25);
            this.lbHeadName.TabIndex = 7;
            this.lbHeadName.Text = "Status";
            // 
            // txtStatus
            // 
            this.txtStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtStatus.Location = new System.Drawing.Point(41, 62);
            this.txtStatus.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ReadOnly = true;
            this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtStatus.Size = new System.Drawing.Size(615, 266);
            this.txtStatus.TabIndex = 6;
            this.tt.SetToolTip(this.txtStatus, "Status Running");
            // 
            // pnDisplay
            // 
            this.pnDisplay.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnDisplay.BackColor = System.Drawing.Color.White;
            this.pnDisplay.Controls.Add(this.lbPercent);
            this.pnDisplay.Controls.Add(this.pgbLoad);
            this.pnDisplay.Controls.Add(this.lbHeadName);
            this.pnDisplay.Controls.Add(this.txtStatus);
            this.pnDisplay.Location = new System.Drawing.Point(0, 43);
            this.pnDisplay.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pnDisplay.Name = "pnDisplay";
            this.pnDisplay.Size = new System.Drawing.Size(699, 422);
            this.pnDisplay.TabIndex = 4;
            // 
            // etaxOnethProcess
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(699, 465);
            this.Controls.Add(this.pnDisplay);
            this.Controls.Add(this.pnHead);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "etaxOnethProcess";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            ((System.ComponentModel.ISupportInitialize)(this.pbRestoreDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbMinimize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbMaximize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbClose)).EndInit();
            this.pnHead.ResumeLayout(false);
            this.pnHead.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbLogo)).EndInit();
            this.pnDisplay.ResumeLayout(false);
            this.pnDisplay.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pbRestoreDown;
        private System.Windows.Forms.PictureBox pbMinimize;
        private System.Windows.Forms.PictureBox pbMaximize;
        private System.Windows.Forms.PictureBox pbClose;
        private System.Windows.Forms.Panel pnHead;
        private System.Windows.Forms.Label lbName;
        private System.Windows.Forms.PictureBox pbLogo;
        private System.Windows.Forms.Label lbPercent;
        private System.Windows.Forms.ProgressBar pgbLoad;
        private System.Windows.Forms.Label lbHeadName;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Panel pnDisplay;
        private System.Windows.Forms.ToolTip tt;
    }
}

