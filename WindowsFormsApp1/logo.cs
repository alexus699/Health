using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class logo : Form
    {
        public logo()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.Opacity < 1)
            {
                this.Opacity += 0.05;
            }
            else
            {
                timer1.Stop();
                System.Threading.Thread.Sleep(1500); // ждем 
                this.Close();
            }
        }

        private void logo_Load(object sender, EventArgs e)
        {
            this.Opacity = 0;
            timer1.Start();
        }
    }
}
