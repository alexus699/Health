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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            Form1.form3Cancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
            textBox04.Text = "10";
        }

        // --------------------------------------------------------------------------------------
        // ДАЛЕЕ
        private void button1_Click(object sender, EventArgs e)
        {
            bool itsok = true; // флаг, что все поля заполнены
            if (String.IsNullOrEmpty(textBox01.Text))
            {
                MessageBox.Show("Не заполнено поле 1", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }
            else
            if (String.IsNullOrEmpty(textBox02.Text))
            {
                MessageBox.Show("Не заполнено поле 2", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }
            else
            if (String.IsNullOrEmpty(textBox03.Text))
            {
                MessageBox.Show("Не заполнено поле 3", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }
            else
            if (String.IsNullOrEmpty(textBox04.Text))
            {
                MessageBox.Show("Не заполнено поле 4", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }
            if (itsok)
            {
                // проверка максоценку на 0 
                if (Double.Parse(textBox01.Text) <= 0)
                {
                    MessageBox.Show("Максимальная оценка должна быть больше 0.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    itsok = false;
                }
                // проверка на дробный шаг
                double xmax = Double.Parse(textBox01.Text);
                double dotkol = Double.Parse(textBox04.Text); // точек будет столько +1 крайняя
                double xstep = Math.Round(xmax / dotkol, 0); // рассчитаем шаг по оси х
                if ((xstep * dotkol) < xmax)
                {
                    MessageBox.Show(dotkol + " точек не делит " + xmax + " баллов на целые равные отрезки. Используйте кратные значения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    itsok = false;
                }
            }
            // проверки закончились
            if (itsok) // всё норм - сохраняем и закрываем
            {
                Form1.formXmax = Double.Parse(textBox01.Text);
                Form1.formYmin = Double.Parse(textBox02.Text);
                Form1.FormYmax = Double.Parse(textBox03.Text);
                Form1.formDot = Double.Parse(textBox04.Text);
                Form1.form3Cancel = false;
                this.Close();
            }
        }

        // --------------------------------------------------------------------------------------
        // ОТМЕНА
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // --------------------------------------------------------------------------------------
        // процедура режет все нажатые клавиши, кроме цифр и Backspace (8) и минус (45)
        private void textBox01_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 45) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }
        // --------------------------------------------------------------------------------------

        // -----------------------------------
    }
} // of form
