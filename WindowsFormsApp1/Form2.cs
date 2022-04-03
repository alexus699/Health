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
    public partial class Form2 : Form
    {
        // глобальные переменный
        // public static - доступные из других форм
        public static DataTable Izm1table, Izm2table;  // 2 таблицы для хранения измерений

        public Form2()
        {
            InitializeComponent();
            comboBox2.DataSource = Izm1table;
            comboBox2.DisplayMember = "name";
            comboBox2.ValueMember = "n";
            comboBox4.DataSource = Izm2table;
            comboBox4.DisplayMember = "name";
            comboBox4.ValueMember = "n";
            Form1.form2Cancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
            textBox3.Text = "1";
            textBox1.Focus(); // курсор на 1ый текстбокс
        }


        // =====================================================================================
        // ОТМЕНА
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // =====================================================================================
        // ДАЛЕЕ
        private void button1_Click(object sender, EventArgs e)
        {
            bool itsok = true; // флаг, что все поля заполнены
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Не заполнено поле 'название оценки'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }else
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Не заполнено поле 'единицы измерения'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }else
            if (String.IsNullOrEmpty(comboBox5.Text))
            {
                MessageBox.Show("Не заполнено поле 'пол'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }else
            if (String.IsNullOrEmpty(comboBox1.Text))
            {
                MessageBox.Show("Не заполнено поле 'способ расчета'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }else
            if ((String.IsNullOrEmpty(comboBox3.Text)) && (comboBox1.SelectedIndex == 1))
                {
                MessageBox.Show("Выберите знак расчета ( - + * / )", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                itsok = false;
            }

            if (itsok) // всё норм - сохраняем и закрываем
            {
                Form1.formName = textBox1.Text;
                Form1.formEd = textBox2.Text;
                Form1.formSposob = comboBox1.Text;
                Form1.formIzm1 = comboBox2.Text;
                Form1.formZnak = comboBox3.Text;
                Form1.formIzm2 = comboBox4.Text;
                Form1.formMult = textBox3.Text;
                Form1.formSex = comboBox5.Text.Substring(0,1); // 1ая буква от слова (м или ж)
                Form1.form2Cancel = false;
                this.Close();
            }
        } // of button1_Click

        // --------------------------------------------------------------------------------------
        // процедура режет все нажатые клавиши, кроме цифр 1,0  и Backspace (8)
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (number != 48 && number != 49 && number != 8) // цифры и клавиша BackSpace
                {
                    e.Handled = true;
                }
        }

        // =====================================================================================
        // если выбрано "копируется" - закрываем знак и второе измерение
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                comboBox3.Visible = false;
                comboBox4.Visible = false;
                label6.Visible = false;
                textBox3.Visible = false;
                textBox3.Text = "1";
            } else {
                comboBox3.Visible = true;
                comboBox4.Visible = true;
                label6.Visible = true;
                textBox3.Visible = true;
            }
        }

        // =====================================================================================
    }
} // of form
