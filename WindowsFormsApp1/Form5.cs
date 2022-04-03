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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
            Form1.form5Cancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
            textBox1.Text = "";
        }

        // --------------------------------------------------------------------------------------
        // ДАЛЕЕ
        private void button14_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Поле не заполнено.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Form1.form5param1 = textBox1.Text;
                Form1.form5Cancel = false;
                this.Close();
            }
        }

        // --------------------------------------------------------------------------------------
        // ОТМЕНА
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // --------------------------------------------------------------------------------------
        // нажата Enter
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13) // клавиша Enter
                if (String.IsNullOrEmpty(textBox1.Text))
                {
                    MessageBox.Show("Поле не заполнено.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Form1.form5param1 = textBox1.Text;
                    Form1.form5Cancel = false;
                    this.Close();
                }
        }
    }
}
