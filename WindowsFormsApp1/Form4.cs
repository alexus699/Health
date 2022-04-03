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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
            Form1.form4Cancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
            Form1.formGraf = 0; // пока ничего не выбрано
        }

        // ========================================================================================
        // ДАЛЕЕ
        private void button1_Click(object sender, EventArgs e)
        {
            if (Form1.formGraf == 0)
                MessageBox.Show("Не выбран вид графика", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                Form1.form4Cancel = false;
                this.Close();
            }
        }

        // ========================================================================================
        // ОТМЕНА
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ========================================================================================
        // Выделение нажатого графика
        private void graf_button1_Click(object sender, EventArgs e)
        {
            button1.BackColor = SystemColors.ActiveCaption;
            button2.BackColor = SystemColors.Control;
            button3.BackColor = SystemColors.Control;
            button4.BackColor = SystemColors.Control;
            Form1.formGraf = 1;
        }

        private void graf_button2_Click(object sender, EventArgs e)
        {
            button1.BackColor = SystemColors.Control;
            button2.BackColor = SystemColors.ActiveCaption;
            button3.BackColor = SystemColors.Control;
            button4.BackColor = SystemColors.Control;
            Form1.formGraf = 2;
        }

        private void graf_button3_Click(object sender, EventArgs e)
        {
            button1.BackColor = SystemColors.Control;
            button2.BackColor = SystemColors.Control;
            button3.BackColor = SystemColors.ActiveCaption;
            button4.BackColor = SystemColors.Control;
            Form1.formGraf = 3;
        }

        private void graf_button4_Click(object sender, EventArgs e)
        {
            button1.BackColor = SystemColors.Control;
            button2.BackColor = SystemColors.Control;
            button3.BackColor = SystemColors.Control;
            button4.BackColor = SystemColors.ActiveCaption;
            Form1.formGraf = 4;
        }

    // -----------------------------------
    }
} // of form
