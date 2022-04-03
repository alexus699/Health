using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class SelectFile : Form
    {
        public SelectFile()
        {
            InitializeComponent();
            Form1.formCancel = true; // вдруг закроют крестиком или кнопкой ОТМЕНА
            Form1.dbPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            label1.Text = Form1.dbPath;
            tableRefresh();
        }

        // ******************************************************************************************
        private void tableRefresh()
        {
            listView1.Clear();
            // Заполнение Listview 
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.Columns.Add("Файл", 100);
            listView1.Columns.Add("Размер", 70, HorizontalAlignment.Right);
            listView1.Columns.Add("Дата изменения", 160, HorizontalAlignment.Right);
            string[] arr = new string[4];
            ListViewItem itm;
            // 1ая ссылка - новая база
            arr[0] = "Новая база";
            arr[1] = "-";
            arr[2] = "-";
            itm = new ListViewItem(arr);
            listView1.Items.Add(itm);
            // Читаем названия файлов в папке
            DirectoryInfo scanDir = new DirectoryInfo(Form1.dbPath);
            FileInfo[] scanFiles = scanDir.GetFiles("*.sqlite");
            foreach (FileInfo file in scanFiles)
            {
                //Добавляем в таблицу
                arr[0] = file.Name;
                arr[1] = FormatBytes(file.Length);
                arr[2] = file.LastWriteTime.ToString();
                itm = new ListViewItem(arr);
                listView1.Items.Add(itm);
            }
            // выбираем первую
            listView1.Focus();
            listView1.Items[0].Selected = true;
        }//tableRefresh()

        // ******************************************************************************************
        private void buttonChangeDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog DirDialog = new FolderBrowserDialog();
            //DirDialog.RootFolder = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            DirDialog.SelectedPath = Form1.dbPath;
            DirDialog.Description = "Выберите папку с базой";
            if (DirDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                Form1.dbPath = DirDialog.SelectedPath;
                label1.Text = Form1.dbPath;
                tableRefresh();
                }
        }//buttonChangeDir_Click

        // ******************************************************************************************

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Form1.formCancel = true;
            this.Close();
        }//buttonCancel_Click

        // ******************************************************************************************

        private void buttonOk_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {

                if (listView1.SelectedItems[0].Text == "Новая база")
                    Form1.dbFileName = "new";
                else
                    Form1.dbFileName = listView1.SelectedItems[0].Text;
                Form1.formCancel = false;
                this.Close();
            }
        }//buttonOk_Click

        // ******************************************************************************************
        private static string FormatBytes(long bytes)
        {
            string[] Suffix = { "B", "KB", "MB", "GB", "TB" };
            int i;
            double dblSByte = bytes;
            for (i = 0; i < Suffix.Length && bytes >= 1024; i++, bytes /= 1024)
            {
                dblSByte = bytes / 1024.0;
            }
            return String.Format("{0:0.##} {1}", dblSByte, Suffix[i]);
        } //FormatBytes
        // ******************************************************************************************
    }//form
}//formApp
