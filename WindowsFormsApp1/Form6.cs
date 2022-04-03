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
    public partial class Form6 : Form
    {
        // глобальные переменный
        // public static - доступные из других форм
        // 4 таблицы для приема данных из гланой формы
        public static DataTable Izm1table, // измерения
            Izm2table, // оценки
            Izm3table, // пол
            Izm4table, Izm5table; // фильтры

        // эти таблицы - копии из основной формы
        public static DataTable MyFilters;
        public static DataTable MyFilterValues1, // 10 таблиц с одной колонкой, в которых могут храниться допустимые значения фильтров
                                MyFilterValues2,
                                MyFilterValues3,
                                MyFilterValues4,
                                MyFilterValues5,
                                MyFilterValues6,
                                MyFilterValues7,
                                MyFilterValues8,
                                MyFilterValues9,
                                MyFilterValues10;

        private void lab12_CheckedChanged(object sender, EventArgs e)
        {
            if (lab12.Checked)
                box12.Enabled = true;
            else
                box12.Enabled = false;
        }

        private void lab13_CheckedChanged(object sender, EventArgs e)
        {
            if (lab13.Checked)
                box13.Enabled = true;
            else
                box13.Enabled = false;
        }

        private void lab14_CheckedChanged(object sender, EventArgs e)
        {
            if (lab14.Checked)
                box14.Enabled = true;
            else
                box14.Enabled = false;
        }

        private void lab15_CheckedChanged(object sender, EventArgs e)
        {
            if (lab15.Checked)
                box15.Enabled = true;
            else
                box15.Enabled = false;
        }

        private void lab16_CheckedChanged(object sender, EventArgs e)
        {
            if (lab16.Checked)
                box16.Enabled = true;
            else
                box16.Enabled = false;
        }

        private void lab17_CheckedChanged(object sender, EventArgs e)
        {
            if (lab17.Checked)
                box17.Enabled = true;
            else
                box17.Enabled = false;
        }

        private void lab11_CheckedChanged(object sender, EventArgs e)
        {
            if (lab11.Checked)
                box11.Enabled = true;
            else
                box11.Enabled = false;
        }

        private void lab10_CheckedChanged(object sender, EventArgs e)
        {
            if (lab10.Checked)
                box10.Enabled = true;
            else
                box10.Enabled = false;
        }

        private void lab9_CheckedChanged(object sender, EventArgs e)
        {
            if (lab9.Checked)
                box9.Enabled = true;
            else
                box9.Enabled = false;
        }

        private void lab8_CheckedChanged(object sender, EventArgs e)
        {
            if (lab8.Checked)
                box8.Enabled = true;
            else
                box8.Enabled = false;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
                comboBox3.Enabled = true;
            else
                comboBox3.Enabled = false;
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked)
            {
                box6.Enabled = true;
                box7.Enabled = true;
            }
            else
            {
                box6.Enabled = false;
                box7.Enabled = false;
            }
        }

        private void detBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (detBox1.Checked)
            {
                comboBox9.Enabled = false;
                detBox2.Checked = false;
                detBox3.Checked = false;
                groupBox2.Visible = true;
            }
            else
            if (!(detBox1.Checked)&&(!detBox2.Checked)&&(!detBox3.Checked))
            {
                groupBox2.Visible = false;
                detBox4.Checked = false;
                detBox5.Checked = false;
                detBox6.Checked = false;
            }
        }

        private void detBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (detBox4.Checked)
            {
                comboBox10.Enabled = false;
                detBox5.Checked = false;
                detBox6.Checked = false;
            }
        }

        private void detBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (detBox2.Checked)
            {
                comboBox9.Enabled = false;
                detBox1.Checked = false;
                detBox3.Checked = false;
                groupBox2.Visible = true;
            }
            else
            if (!(detBox1.Checked) && (!detBox2.Checked) && (!detBox3.Checked))
            {
                groupBox2.Visible = false;
                detBox4.Checked = false;
                detBox5.Checked = false;
                detBox6.Checked = false;
            }
        }

        private void detBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (detBox5.Checked)
            {
                comboBox10.Enabled = false;
                detBox4.Checked = false;
                detBox6.Checked = false;
            }
        }

        private void detBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (detBox3.Checked)
            {
                comboBox9.Enabled = true;
                detBox1.Checked = false;
                detBox2.Checked = false;
                groupBox2.Visible = true;
            }
            else
            {
                comboBox9.Enabled = false;
                if (!(detBox1.Checked) && (!detBox2.Checked) && (!detBox3.Checked))
                {
                    groupBox2.Visible = false;
                    detBox4.Checked = false;
                    detBox5.Checked = false;
                    detBox6.Checked = false;
                }
            }
        }

        private void detBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (detBox6.Checked)
            {
                comboBox10.Enabled = true;
                detBox4.Checked = false;
                detBox5.Checked = false;
            }
            else comboBox10.Enabled = false;
        }

        public Form6()
        {
            InitializeComponent();
            // заполняем форму значениями
            comboBox1.DataSource = Izm1table;
            comboBox1.DisplayMember = "name";
            comboBox1.ValueMember = "n";
            comboBox2.DataSource = Izm2table;
            comboBox2.DisplayMember = "name";
            comboBox2.ValueMember = "n";
            // ФИЛЬТРЫ пол
            comboBox3.DataSource = Izm3table;
            comboBox3.DisplayMember = "name";
            comboBox3.ValueMember = "n";
            comboBox3.Enabled = false;
            // детализация ФИЛЬТРЫ
            comboBox9.DataSource = Izm4table;
            comboBox9.DisplayMember = "name";
            comboBox9.ValueMember = "n";
            comboBox10.DataSource = Izm5table;
            comboBox10.DisplayMember = "name";
            comboBox10.ValueMember = "n";
            // ФИЛЬТРЫ ----------
            if (MyFilters.Rows[0].ItemArray[1].Equals(true))
            {
                lab8.Visible = true;
                box8.Visible = true;
                lab8.Text = MyFilters.Rows[0].ItemArray[2].ToString();
                box8.DisplayMember = MyFilterValues1.Columns[0].ColumnName;
                box8.ValueMember = MyFilterValues1.Columns[0].ColumnName;
                box8.DataSource = MyFilterValues1;
                box8.Enabled = false;
            }
            else
            {
                lab8.Visible = false;
                box8.Visible = false;
            }
            // -------
            if (MyFilters.Rows[1].ItemArray[1].Equals(true))
            {
                lab9.Visible = true;
                box9.Visible = true;
                lab9.Text = MyFilters.Rows[1].ItemArray[2].ToString();
                box9.DisplayMember = MyFilterValues2.Columns[0].ColumnName;
                box9.ValueMember = MyFilterValues2.Columns[0].ColumnName;
                box9.DataSource = MyFilterValues2;
                box9.Enabled = false;
            }
            else
            {
                lab9.Visible = false;
                box9.Visible = false;
            }
            // -------
            if (MyFilters.Rows[2].ItemArray[1].Equals(true))
            {
                lab10.Visible = true;
                box10.Visible = true;
                lab10.Text = MyFilters.Rows[2].ItemArray[2].ToString();
                box10.DisplayMember = MyFilterValues3.Columns[0].ColumnName;
                box10.ValueMember = MyFilterValues3.Columns[0].ColumnName;
                box10.DataSource = MyFilterValues3;
                box10.Enabled = false;
            }
            else
            {
                lab10.Visible = false;
                box10.Visible = false;
            }
            // -------
            if (MyFilters.Rows[3].ItemArray[1].Equals(true))
            {
                lab11.Visible = true;
                box11.Visible = true;
                lab11.Text = MyFilters.Rows[3].ItemArray[2].ToString();
                box11.DisplayMember = MyFilterValues4.Columns[0].ColumnName;
                box11.ValueMember = MyFilterValues4.Columns[0].ColumnName;
                box11.DataSource = MyFilterValues4;
                box11.Enabled = false;
            }
            else
            {
                lab11.Visible = false;
                box11.Visible = false;
            }
            // -------
            if (MyFilters.Rows[4].ItemArray[1].Equals(true))
            {
                lab12.Visible = true;
                box12.Visible = true;
                lab12.Text = MyFilters.Rows[4].ItemArray[2].ToString();
                box12.DisplayMember = MyFilterValues5.Columns[0].ColumnName;
                box12.ValueMember = MyFilterValues5.Columns[0].ColumnName;
                box12.DataSource = MyFilterValues5;
                box12.Enabled = false;
            }
            else
            {
                lab12.Visible = false;
                box12.Visible = false;
            }
            // -------
            if (MyFilters.Rows[5].ItemArray[1].Equals(true))
            {
                lab13.Visible = true;
                box13.Visible = true;
                lab13.Text = MyFilters.Rows[5].ItemArray[2].ToString();
                box13.DisplayMember = MyFilterValues6.Columns[0].ColumnName;
                box13.ValueMember = MyFilterValues6.Columns[0].ColumnName;
                box13.DataSource = MyFilterValues6;
                box13.Enabled = false;
            }
            else
            {
                lab13.Visible = false;
                box13.Visible = false;
            }
            // -------
            if (MyFilters.Rows[6].ItemArray[1].Equals(true))
            {
                lab14.Visible = true;
                box14.Visible = true;
                lab14.Text = MyFilters.Rows[6].ItemArray[2].ToString();
                box14.DisplayMember = MyFilterValues7.Columns[0].ColumnName;
                box14.ValueMember = MyFilterValues7.Columns[0].ColumnName;
                box14.DataSource = MyFilterValues7;
                box14.Enabled = false;
            }
            else
            {
                lab14.Visible = false;
                box14.Visible = false;
            }
            // -------
            if (MyFilters.Rows[7].ItemArray[1].Equals(true))
            {
                lab15.Visible = true;
                box15.Visible = true;
                lab15.Text = MyFilters.Rows[7].ItemArray[2].ToString();
                box15.DisplayMember = MyFilterValues8.Columns[0].ColumnName;
                box15.ValueMember = MyFilterValues8.Columns[0].ColumnName;
                box15.DataSource = MyFilterValues8;
                box15.Enabled = false;
            }
            else
            {
                lab15.Visible = false;
                box15.Visible = false;
            }
            // -------
            if (MyFilters.Rows[8].ItemArray[1].Equals(true))
            {
                lab16.Visible = true;
                box16.Visible = true;
                lab16.Text = MyFilters.Rows[8].ItemArray[2].ToString();
                box16.DisplayMember = MyFilterValues9.Columns[0].ColumnName;
                box16.ValueMember = MyFilterValues9.Columns[0].ColumnName;
                box16.DataSource = MyFilterValues9;
                box16.Enabled = false;
            }
            else
            {
                lab16.Visible = false;
                box16.Visible = false;
            }
            // -------
            if (MyFilters.Rows[9].ItemArray[1].Equals(true))
            {
                lab17.Visible = true;
                box17.Visible = true;
                lab17.Text = MyFilters.Rows[9].ItemArray[2].ToString();
                box17.DisplayMember = MyFilterValues10.Columns[0].ColumnName;
                box17.ValueMember = MyFilterValues10.Columns[0].ColumnName;
                box17.DataSource = MyFilterValues10;
                box17.Enabled = false;
            }
            else
            {
                lab17.Visible = false;
                box17.Visible = false;
            }
            box6.Value = DateTime.Today;
            box6.Enabled = false;
            box7.Value = DateTime.Today;
            box7.Enabled = false;
            // по фильтрам

            Form1.form6Cancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
        }

        // --------------------------------------------------------------------------------------
        // ДАЛЕЕ
        private void button14_Click(object sender, EventArgs e)
        {
            bool itsok = true; // флаг, что все поля заполнены
            if ((checkBox2.Checked) && (!checkBox3.Checked))
                if (MessageBox.Show("Фильтр по полу не выбран. Всё равно продолжить?", "Внимание!", MessageBoxButtons.YesNo) == DialogResult.No)
                    itsok = false;
            if ((!detBox1.Checked) && (!detBox2.Checked) && (!detBox3.Checked))
            {
                MessageBox.Show("Деталиация по столбцам должна быть выбрана!", "Внимание!");
                itsok = false;
            }
            if ((detBox1.Checked) && (detBox4.Checked) || (detBox2.Checked) && (detBox5.Checked))
            {
                MessageBox.Show("Деталиация по строкам и столбцам не может быть одинаковая!", "Внимание!");
                itsok = false;
            }
            if ((detBox3.Checked) && (detBox6.Checked) && (comboBox9.Text == comboBox10.Text))
            {
                MessageBox.Show("Деталиация по строкам и столбцам не может быть одинаковая!", "Внимание!");
                itsok = false;
            }
            if (itsok) // всё норм - сохраняем и закрываем
            {
                if (checkBox4.Checked) Form1.form6param19 = true; else Form1.form6param19 = false;
                if (checkBox5.Checked) Form1.form6param20 = true; else Form1.form6param20 = false;
                if (checkBox1.Checked) Form1.form6param1 = comboBox1.Text; else Form1.form6param1 = ""; // измерение
                if (checkBox2.Checked) Form1.form6param2 = comboBox2.Text; else Form1.form6param2 = ""; // оценка
                if (checkBox3.Checked) Form1.form6param3 = comboBox3.Text; else Form1.form6param3 = ""; // пол
                if (checkBox14.Checked) // период
                {
                    Form1.form6param4 = true;
                    Form1.form6param5 = DateTime.Parse(box6.Text);
                    Form1.form6param6 = DateTime.Parse(box7.Text);
                }
                else
                {
                    Form1.form6param4 = false;
                    Form1.form6param5 = DateTime.Parse("01.01.2000");
                    Form1.form6param6 = DateTime.Parse("01.01.2000");
                }
                if (lab8.Checked)  Form1.form6param7 =  box8.Text;  else Form1.form6param7 = ""; // фильтр
                if (lab9.Checked)  Form1.form6param8 =  box9.Text;  else Form1.form6param8 = ""; // фильтр
                if (lab10.Checked) Form1.form6param9 =  box10.Text; else Form1.form6param9 = ""; // фильтр
                if (lab11.Checked) Form1.form6param10 = box11.Text; else Form1.form6param10 = ""; // фильтр
                if (lab12.Checked) Form1.form6param11 = box12.Text; else Form1.form6param11 = ""; // фильтр
                if (lab13.Checked) Form1.form6param12 = box13.Text; else Form1.form6param12 = ""; // фильтр
                if (lab14.Checked) Form1.form6param13 = box14.Text; else Form1.form6param13 = ""; // фильтр
                if (lab15.Checked) Form1.form6param14 = box15.Text; else Form1.form6param14 = ""; // фильтр
                if (lab16.Checked) Form1.form6param15 = box16.Text; else Form1.form6param15 = ""; // фильтр
                if (lab17.Checked) Form1.form6param16 = box17.Text; else Form1.form6param16 = ""; // фильтр
                // детализация1
                Form1.form6param17 = "";
                if (detBox1.Checked) Form1.form6param17 = "SEX";
                if (detBox2.Checked) Form1.form6param17 = "YEAR";
                if (detBox3.Checked) Form1.form6param17 = comboBox9.Text;
                // детализация2
                Form1.form6param18 = "";
                if (detBox4.Checked) Form1.form6param18 = "SEX";
                if (detBox5.Checked) Form1.form6param18 = "YEAR";
                if (detBox6.Checked) Form1.form6param18 = comboBox10.Text;
                Form1.form6Cancel = false;
                this.Close();
            }
        }

        // --------------------------------------------------------------------------------------
        // щелкнули по ПО ИЗМЕРЕНИЮ
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                comboBox1.Enabled = true;
                checkBox2.Checked = false;
                comboBox2.Enabled = false;
            }
            else
            {
                comboBox1.Enabled = false;
                checkBox2.Checked = true;
                comboBox2.Enabled = true;
            }
        }

        // --------------------------------------------------------------------------------------
        // щелкнули по ПО ОЦЕНКЕ
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                comboBox2.Enabled = true;
                checkBox1.Checked = false;
                comboBox1.Enabled = false;
            }
            else
            {
                comboBox2.Enabled = false;
                checkBox1.Checked = true;
                comboBox1.Enabled = true;
            }
        }

        // --------------------------------------------------------------------------------------
        // ОТМЕНА
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
