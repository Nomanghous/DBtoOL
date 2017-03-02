namespace AccesstoOutlook
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbAccessDBPath = new System.Windows.Forms.TextBox();
            this.btnDbPath = new System.Windows.Forms.Button();
            this.tbTableName = new System.Windows.Forms.TextBox();
            this.labelTable = new System.Windows.Forms.Label();
            this.btnContactsFolder = new System.Windows.Forms.Button();
            this.tbContactsFolder = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnTransfer = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.ProgressBarHandler_Timer = new System.Windows.Forms.Timer(this.components);
            this.label2 = new System.Windows.Forms.Label();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.LBWait = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(120, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Access Database Path";
            // 
            // tbAccessDBPath
            // 
            this.tbAccessDBPath.Location = new System.Drawing.Point(366, 81);
            this.tbAccessDBPath.Name = "tbAccessDBPath";
            this.tbAccessDBPath.ReadOnly = true;
            this.tbAccessDBPath.Size = new System.Drawing.Size(194, 20);
            this.tbAccessDBPath.TabIndex = 1;
            // 
            // btnDbPath
            // 
            this.btnDbPath.Location = new System.Drawing.Point(566, 79);
            this.btnDbPath.Name = "btnDbPath";
            this.btnDbPath.Size = new System.Drawing.Size(37, 23);
            this.btnDbPath.TabIndex = 2;
            this.btnDbPath.Text = "...";
            this.btnDbPath.UseVisualStyleBackColor = true;
            this.btnDbPath.Click += new System.EventHandler(this.btnDbPath_Click);
            // 
            // tbTableName
            // 
            this.tbTableName.Location = new System.Drawing.Point(366, 139);
            this.tbTableName.Name = "tbTableName";
            this.tbTableName.Size = new System.Drawing.Size(194, 20);
            this.tbTableName.TabIndex = 4;
            // 
            // labelTable
            // 
            this.labelTable.AutoSize = true;
            this.labelTable.Location = new System.Drawing.Point(120, 142);
            this.labelTable.Name = "labelTable";
            this.labelTable.Size = new System.Drawing.Size(158, 13);
            this.labelTable.TabIndex = 3;
            this.labelTable.Text = "Access Database Table Name :";
            // 
            // btnContactsFolder
            // 
            this.btnContactsFolder.Enabled = false;
            this.btnContactsFolder.Location = new System.Drawing.Point(566, 108);
            this.btnContactsFolder.Name = "btnContactsFolder";
            this.btnContactsFolder.Size = new System.Drawing.Size(37, 23);
            this.btnContactsFolder.TabIndex = 8;
            this.btnContactsFolder.Text = "...";
            this.btnContactsFolder.UseVisualStyleBackColor = true;
            this.btnContactsFolder.Click += new System.EventHandler(this.btnContactsFolder_Click);
            // 
            // tbContactsFolder
            // 
            this.tbContactsFolder.Location = new System.Drawing.Point(366, 110);
            this.tbContactsFolder.Name = "tbContactsFolder";
            this.tbContactsFolder.Size = new System.Drawing.Size(194, 20);
            this.tbContactsFolder.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(120, 113);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(121, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Outlook Contacts Folder";
            // 
            // btnTransfer
            // 
            this.btnTransfer.Location = new System.Drawing.Point(248, 278);
            this.btnTransfer.Name = "btnTransfer";
            this.btnTransfer.Size = new System.Drawing.Size(227, 23);
            this.btnTransfer.TabIndex = 9;
            this.btnTransfer.Text = "Transfer All";
            this.btnTransfer.UseVisualStyleBackColor = true;
            this.btnTransfer.Click += new System.EventHandler(this.btnTransfer_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(248, 307);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(227, 23);
            this.btnUpdate.TabIndex = 9;
            this.btnUpdate.Text = "Transfer Updates Only";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(248, 336);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(227, 23);
            this.btnDelete.TabIndex = 9;
            this.btnDelete.Text = "Delete Contacts From Outlook";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(80, 423);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(480, 23);
            this.progressBar1.TabIndex = 11;
            // 
            // ProgressBarHandler_Timer
            // 
            this.ProgressBarHandler_Timer.Enabled = true;
            this.ProgressBarHandler_Timer.Interval = 1000;
            this.ProgressBarHandler_Timer.Tick += new System.EventHandler(this.ProgressBarHandler_Timer_Tick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(613, 451);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Powered by : Logixcess";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.DarkSlateGray;
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.btnContactsFolder);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.tbAccessDBPath);
            this.panel1.Controls.Add(this.btnDelete);
            this.panel1.Controls.Add(this.btnDbPath);
            this.panel1.Controls.Add(this.btnUpdate);
            this.panel1.Controls.Add(this.labelTable);
            this.panel1.Controls.Add(this.btnTransfer);
            this.panel1.Controls.Add(this.tbTableName);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.tbContactsFolder);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(747, 417);
            this.panel1.TabIndex = 13;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(248, 366);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(227, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Open App Directory";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(120, 181);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(101, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Select Logging Flag";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Select Flag",
            "Level 0",
            "Level 1",
            "Level 2"});
            this.comboBox1.Location = new System.Drawing.Point(366, 178);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(194, 21);
            this.comboBox1.TabIndex = 10;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // LBWait
            // 
            this.LBWait.AutoSize = true;
            this.LBWait.Location = new System.Drawing.Point(566, 428);
            this.LBWait.Name = "LBWait";
            this.LBWait.Size = new System.Drawing.Size(70, 13);
            this.LBWait.TabIndex = 14;
            this.LBWait.Text = "Please wait...";
            this.LBWait.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(747, 472);
            this.Controls.Add(this.LBWait);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.progressBar1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(763, 510);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbAccessDBPath;
        private System.Windows.Forms.Button btnDbPath;
        private System.Windows.Forms.TextBox tbTableName;
        private System.Windows.Forms.Label labelTable;
        private System.Windows.Forms.Button btnContactsFolder;
        private System.Windows.Forms.TextBox tbContactsFolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnTransfer;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Timer ProgressBarHandler_Timer;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label LBWait;
    }
}

