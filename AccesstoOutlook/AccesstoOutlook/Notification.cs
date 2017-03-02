using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccesstoOutlook
{
    public partial class Notification : Form
    {
        public Notification()
        {
            InitializeComponent();
        }

        private void Notification_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            lbTotal.Location = new Point(panel1.Width / 2, panel1.Height / 2);
            lbTotal.Text =  Passvalues.totalRecords;
            lbSuccess.Text = Passvalues.message;
            Timer.Start();

            int x = Screen.PrimaryScreen.WorkingArea.Width - this.Width;
            int y = Screen.PrimaryScreen.WorkingArea.Height - this.Height;
            this.Location = new Point(x, y);
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void Notification_FormClosed(object sender, FormClosedEventArgs e)
        {
            Timer.Stop();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Timer.Stop();
            this.Dispose();
        }
    }
}
