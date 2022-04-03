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
    public partial class Form7 : Form
    {
        // глобальные переменный
        // public static - доступные из других форм
        // 4 таблицы для приема данных из гланой формы
        public static DataTable Izm1table, // измерения
            Izm2table, // оценки
            Izm3table; // пол

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

        // --------------------------------------------------------------------------------------
        // ДАЛЕЕ
        private void button14_Click(object sender, EventArgs e)
        {
            if (checkBox3.Checked) Form1.form7param3 = comboBox3.Text; else Form1.form7param3 = ""; // пол
            if (checkBox14.Checked) // период
            {
                Form1.form7param4 = true;
                Form1.form7param5 = DateTime.Parse(box6.Text);
                Form1.form7param6 = DateTime.Parse(box7.Text);
            }
            else
            {
                Form1.form7param4 = false;
                Form1.form7param5 = DateTime.Parse("01.01.2000");
                Form1.form7param6 = DateTime.Parse("01.01.2000");
            }
            if (lab8.Checked) Form1.form7param7 = box8.Text; else Form1.form7param7 = ""; // фильтр
            if (lab9.Checked) Form1.form7param8 = box9.Text; else Form1.form7param8 = ""; // фильтр
            if (lab10.Checked) Form1.form7param9 = box10.Text; else Form1.form7param9 = ""; // фильтр
            if (lab11.Checked) Form1.form7param10 = box11.Text; else Form1.form7param10 = ""; // фильтр
            if (lab12.Checked) Form1.form7param11 = box12.Text; else Form1.form7param11 = ""; // фильтр
            if (lab13.Checked) Form1.form7param12 = box13.Text; else Form1.form7param12 = ""; // фильтр
            if (lab14.Checked) Form1.form7param13 = box14.Text; else Form1.form7param13 = ""; // фильтр
            if (lab15.Checked) Form1.form7param14 = box15.Text; else Form1.form7param14 = ""; // фильтр
            if (lab16.Checked) Form1.form7param15 = box16.Text; else Form1.form7param15 = ""; // фильтр
            if (lab17.Checked) Form1.form7param16 = box17.Text; else Form1.form7param16 = ""; // фильтр
            Form1.form7tab1 = Izm1table;
            Form1.form7tab2 = Izm2table;
            Form1.form7Cancel = false;
            this.Close();
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

        public Form7()
        {
            InitializeComponent();
            // для удобства в процедуре Report2 запомним в отдельной колонке названия полей (до удаления пустых)
            Izm1table.Columns.Add(new DataColumn("izm", typeof(String)));
            Izm1table.Rows[0]["izm"] = "izm01";
            Izm1table.Rows[1]["izm"] = "izm02";
            Izm1table.Rows[2]["izm"] = "izm03";
            Izm1table.Rows[3]["izm"] = "izm04";
            Izm1table.Rows[4]["izm"] = "izm05";
            Izm1table.Rows[5]["izm"] = "izm06";
            Izm1table.Rows[6]["izm"] = "izm07";
            Izm1table.Rows[7]["izm"] = "izm08";
            Izm1table.Rows[8]["izm"] = "izm09";
            Izm1table.Rows[9]["izm"] = "izm10";
            Izm1table.Rows[10]["izm"] = "izm11";
            Izm1table.Rows[11]["izm"] = "izm12";
            Izm1table.Rows[12]["izm"] = "izm13";
            Izm1table.Rows[13]["izm"] = "izm14";
            Izm1table.Rows[14]["izm"] = "izm15";
            Izm1table.Rows[15]["izm"] = "izm16";
            Izm1table.Rows[16]["izm"] = "izm17";
            Izm1table.Rows[17]["izm"] = "izm18";
            Izm1table.Rows[18]["izm"] = "izm19";
            Izm1table.Rows[19]["izm"] = "izm20";
            // заполняем измерения
            for (int i = Izm1table.Rows.Count-1; i > 0; i--) // уберем пустые
            {
                if (Izm1table.Rows[i]["вкл"].Equals(false)) // (галочка не стоит)
                    Izm1table.Rows[i].Delete();
            }
            dataGridView1.DataSource = Izm1table;
            dataGridView1.Columns[0].Visible = false; // некоторые колонки не видны
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // заполняем оценки
            dataGridView2.DataSource = Izm2table;
            dataGridView2.Columns[1].ReadOnly = true;
            dataGridView2.Columns[2].Visible = false; // некоторые колонки не видны
            dataGridView2.Columns[3].Visible = false;
            dataGridView2.Columns[4].Visible = false;
            dataGridView2.Columns[5].Visible = false;
            dataGridView2.Columns[6].Visible = false;
            dataGridView2.Columns[7].ReadOnly = true;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // запрещаем сортировки в таблицах DGView
            foreach (DataGridViewColumn column in dataGridView1.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (DataGridViewColumn column in dataGridView2.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            // ФИЛЬТРЫ пол
            comboBox3.DataSource = Izm3table;
            comboBox3.DisplayMember = "name";
            comboBox3.ValueMember = "n";
            comboBox3.Enabled = false;
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
            // }по фильтрам
            Form1.form7Cancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
