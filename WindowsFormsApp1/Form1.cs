using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;
using ZedGraph;
using System.Diagnostics;
using System.Drawing.Printing;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        // глобальные переменные-----------------------------------------------------
        DataSet MySettings; // ...набор, объединяющий все 12 таблиц
        System.Data.DataTable MyFilters, MyIzm; // две таблицы настроек
        System.Data.DataTable MyFilterValues1, // 10 таблиц с одной колонкой, в которых могут храниться допустимые значения фильтров
                MyFilterValues2,
                MyFilterValues3,
                MyFilterValues4,
                MyFilterValues5,
                MyFilterValues6,
                MyFilterValues7,
                MyFilterValues8,
                MyFilterValues9,
                MyFilterValues10;

        DataSet MyMarks; // ...набор, объединяющий главную таблицу и таблицы с точками оценок
        System.Data.DataTable MySpisok; // список оценок и ссылки на таблицы
        System.Data.DataTable MyGraphs; // точки графиков оценок

        // для Обработки перетаскивания в Listbox1
        static int listbox1offset = 15; //чтоб имя перетаскиваемого элемента рядом с мышкой было 
        Stopwatch sw = new Stopwatch(); //чтобы не реагировал на обычный клик мышкой
        bool listbox1moving = false;
        int listbox1index = -1;
        System.Data.DataTable TempFilterValues;

        // для постройки графика
        double dotkol = 10;    // по идее больше 10 точек не надо
        double xmin = 0;      // по идее не выбирается - всегда 0 - минимальная оценка
        double xstep;
        double xmax;
        double ystep;
        double ymin;
        double ymax;

        // глоб.переменные для обмена между окнами
        // public static - доступные из других форм
        public static string formName, formEd, formSposob, formIzm1, formZnak, formIzm2, formMult, formSex;
        public static double formXmax, formDot, formYmin, FormYmax, formGraf;
        public static bool formCancel, form2Cancel, form3Cancel, form4Cancel, form5Cancel, form6Cancel, form7Cancel; // для отслеживания ОТМЕНЫ, для каждой формы своя - иначе ошибка
        public static string form5param1;
        public static string form6param1, form6param2, form6param3, form6param7, form6param8, form6param9, form6param10, form6param11, form6param12, form6param13, form6param14, form6param15, form6param16, form6param17, form6param18;
        public static string form7param7, form7param8, form7param9, form7param10, form7param11, form7param12, form7param13, form7param14, form7param15, form7param16;
        public static string form7param3;
        public static bool form6param4, form6param19, form6param20;
        public static bool form7param4;
        public static DateTime form6param5, form6param6;
        public static DateTime form7param5, form7param6;
        public static DataTable form7tab1, form7tab2;
        public static String dbFileName;
        public static String dbName; // тоже имя, только без разширения .sqlite
        public static String dbPath;

        // Глобальные переменные MySQL
        private SQLiteConnection dbConn;
        private DataSet myDataSet;
        private DataSet myDataSet2;
        SQLiteDataAdapter myDataAdapter, myDataAdapter2; // адаптеры для отображения таблиц sqlite в datagridах
        bool FlagStud, FlagCheck, FlagMark, FlagIzm; // флаги для отслеживания несохраненных изменений
        // для вкладки Ввод
        private DataSet vvodDataSet;
        SQLiteDataAdapter vvodDataAdapter;
        String id_found = "";
        bool student_found;

        // для печати TextRichBox
        private int checkPrint;
        StringBuilder tableRtf;
        private int passport_wide = 3000; // ширина 1ой колонки в паспорте ФР
        private int passport_wide2 = 1500; // ширина колонок с медосмотрами в паспорте ФР
        private string marksXML         = "-marks.xml";      // название файла с оценками
        private string TemplateMARKS    = "Health01.new";   // название файла с шаблоном для файла оценок
        private string settingsXML      = "-settings.xml";   // название файла с настройками
        private string TemplateSETTINGS = "Health02.new";   // название файла с шаблоном для файла ностроек
        private string passportDOCX      = "-report01.docx";   // название файла с шаблоном паспорта
        private string TemplateRTF01    = "Health03.new";   // название файла с шаблоном для шаблона паспорта
        private string report01RTF      = "-report02.rtf";   // название файла с шаблоном отчета по средним
        private string TemplateRTF02    = "Health04.new";   // название файла с шаблоном для шаблона по средним


        public Form1()
        {
            InitializeComponent();
            // логотип
            // отключим на время отладки !!!
            Form formlogo = new logo();
            formlogo.ShowDialog();

            // для печати TextRichBox
            this.printDocument1.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.printDocument1_BeginPrint);
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            button10.Visible = false; // пока спрячем кнопку печати

            dbConn = new SQLiteConnection();
            // Выбираем файл базы
            Form SelectFileForm = new SelectFile();
            SelectFileForm.ShowDialog();
            if (formCancel == true) // вдруг закроют крестиком или кнопкой ОТМЕНА
            {
                MessageBox.Show("База не загружена.");
                //bdFileCreate(); // создаём пустую
                // закрываем приложение
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            else if (dbFileName == "new") bdFileCreate(); // создаём пустую
            bdFileConnect(); // открываем базу (имя файла указано или новое имя файла, в обоих случаях)

            //этот Label нужен для того чтобы при перетаскивании в Listboxe(для хранения Значений фильтров) держать название перетаскиваемого элемента рядом
            hidden_label.Visible = false;
            hidden_label.BackColor = Color.Transparent;
            hidden_label.BorderStyle = BorderStyle.FixedSingle;

            // ЗАГРУЗКА НАСТРОЕК ИЗ XML-файлов
            InitSettingsDataSet(false); // без сброса настроек на умолчанию
            LoadSettingsXML();
            dataGridView1.DataSource = MyFilters;// Загруженные XML данные подключаем к DGVюшкам
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.DataSource = MyIzm;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // оценки
            InitMarksDataSet(false);// без сброса настроек на умолчанию
            LoadMarksXML();
            dataGridView3.DataSource = MySpisok;// Загруженные XML данные подключаем к DGVюшкам
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // настройки zedGraph
            zedGraphControl1.EditButtons = MouseButtons.Left; // Разрешим перемещать точки
            zedGraphControl1.IsEnableHEdit = true; // Точки можно перемещать, как по горизонтали,...
            zedGraphControl1.IsEnableVEdit = true; // ... так и по вертикали.
            // Подпишемся на событие, вызываемое после перемещения точки
            zedGraphControl1.PointEditEvent += new ZedGraphControl.PointEditHandler(zedGraphControl1_PointEditEvent);
            // здесь уже можно рисовать DrawGraph();
            // Подписываем колонки в базе данных
            //dataGridViewStudent.Columns["id"].Visible = false;
            dataGridViewStudent.Columns["passport"].HeaderText = "документ";
            dataGridViewStudent.Columns["f"].HeaderText = "фамилия";
            dataGridViewStudent.Columns["i"].HeaderText = "имя";
            dataGridViewStudent.Columns["o"].HeaderText = "отчество";
            dataGridViewStudent.Columns["sex"].HeaderText = "пол";
            dataGridViewStudent.Columns["born"].HeaderText = "дата рождения";
            // служебные колонки в оценках никогда не видны
            dataGridView3.Columns[3].Visible = false;
            dataGridView3.Columns[4].Visible = false;
            dataGridView3.Columns[5].Visible = false;
            dataGridView3.Columns[6].Visible = false;
            // запрещаем сортировки в колонках оценки, фильтры и измерения
            foreach (DataGridViewColumn column in dataGridView1.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (DataGridViewColumn column in dataGridView2.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (DataGridViewColumn column in dataGridView3.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            // на 1ой вкладке отмечаем, что изменений пока нет
            FlagStud = false;
            FlagCheck = false;
            button11.Visible = false;
            button12.Visible = false;
            dataGridViewStudent.Enabled = true;
            medCheckDataGridView.Enabled = true;
        } // Form

        // **********************************************************************************************

        private void bdFileCreate()
        {
            int i;
            bool NoErrors = true;
            SQLiteCommand sqlCmd = new SQLiteCommand();
            dbFileName = "base";
            for (i = 0; i < 10000; i++)
            {
                if (!File.Exists(dbPath + "\\" + dbFileName + i.ToString() + ".sqlite"))
                    break;
            }
            dbFileName += i.ToString(); // добавим цифру к новой базе
            dbName = dbFileName; // сохраним без расширения
            dbFileName += ".sqlite"; // добавим расширение

            SQLiteConnection.CreateFile(dbPath + "\\" + dbFileName);
            try
            {
                dbConn = new SQLiteConnection("Data Source=" + dbPath + "\\" + dbFileName + ";Version=3;");
                dbConn.Open();
                sqlCmd.Connection = dbConn;
                sqlCmd.CommandText = "CREATE TABLE IF NOT EXISTS Student ([Id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, [passport] NCHAR(10) NULL, [f] NVARCHAR(20) NULL, [i] NVARCHAR(20) NULL, [o] NVARCHAR(20) NULL, [sex] NVARCHAR(1) NULL, [born] DATETIME NULL);";
                sqlCmd.ExecuteNonQuery();
                sqlCmd.CommandText = "INSERT INTO Student(Id,passport,f,i,o,sex,born) VALUES(1,'12345678','Иванов', 'Иван', 'Иванович', 'м', '2000-12-31');";
                sqlCmd.ExecuteNonQuery();
                sqlCmd.CommandText = "CREATE TABLE IF NOT EXISTS MedCheck ( [Id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, [idStudent] INT NULL, [data] DATETIME NULL, [fil01] NVARCHAR(20) NULL, [fil02] NVARCHAR(20) NULL, [fil03] NVARCHAR(20) NULL, [fil04] NVARCHAR(20) NULL, [fil05] NVARCHAR(20) NULL, [fil06] NVARCHAR(20) NULL, [fil07] NVARCHAR(20) NULL, [fil08] NVARCHAR(20) NULL, [fil09] NVARCHAR(20) NULL, [fil10] NVARCHAR(20) NULL, [izm01] SMALLINT NULL, [izm02] SMALLINT NULL, [izm03] SMALLINT NULL, [izm04] SMALLINT NULL, [izm05] SMALLINT NULL, [izm06] SMALLINT NULL, [izm07] SMALLINT NULL, [izm08] SMALLINT NULL, [izm09] SMALLINT NULL, [izm10] SMALLINT NULL, [izm11] SMALLINT NULL, [izm12] SMALLINT NULL, [izm13] SMALLINT NULL, [izm14] SMALLINT NULL, [izm15] SMALLINT NULL, [izm16] SMALLINT NULL, [izm17] SMALLINT NULL, [izm18] SMALLINT NULL, [izm19] SMALLINT NULL, [izm20] SMALLINT NULL, CONSTRAINT[FK_osmotr_ToStudent] FOREIGN KEY([idStudent]) REFERENCES Student ([Id]));";
                //sqlCmd.CommandText = "CREATE TABLE IF NOT EXISTS MedCheck ( [Id] INT IDENTITY(1, 1) NOT NULL, [idStudent] INT NULL, [data] DATETIME NULL, [fil01] NVARCHAR(20) NULL, [fil02] NVARCHAR(20) NULL, [fil03] NVARCHAR(20) NULL, [fil04] NVARCHAR(20) NULL, [fil05] NVARCHAR(20) NULL, [fil06] NVARCHAR(20) NULL, [fil07] NVARCHAR(20) NULL, [fil08] NVARCHAR(20) NULL, [fil09] NVARCHAR(20) NULL, [fil10] NVARCHAR(20) NULL, [izm01] SMALLINT NULL, [izm02] SMALLINT NULL, [izm03] SMALLINT NULL, [izm04] SMALLINT NULL, [izm05] SMALLINT NULL, [izm06] SMALLINT NULL, [izm07] SMALLINT NULL, [izm08] SMALLINT NULL, [izm09] SMALLINT NULL, [izm10] SMALLINT NULL, [izm11] SMALLINT NULL, [izm12] SMALLINT NULL, [izm13] SMALLINT NULL, [izm14] SMALLINT NULL, [izm15] SMALLINT NULL, [izm16] SMALLINT NULL, [izm17] SMALLINT NULL, [izm18] SMALLINT NULL, [izm19] SMALLINT NULL, [izm20] SMALLINT NULL, PRIMARY KEY ([Id] ASC), CONSTRAINT[FK_osmotr_ToStudent] FOREIGN KEY([idStudent]) REFERENCES Student ([Id]));";
                // sqlCmd.CommandText = "CREATE TABLE IF NOT EXISTS MedCheck ( [idStudent] INT NULL, [data] DATETIME NULL, [fil01] NVARCHAR(20) NULL, [fil02] NVARCHAR(20) NULL, [fil03] NVARCHAR(20) NULL, [fil04] NVARCHAR(20) NULL, [fil05] NVARCHAR(20) NULL, [fil06] NVARCHAR(20) NULL, [fil07] NVARCHAR(20) NULL, [fil08] NVARCHAR(20) NULL, [fil09] NVARCHAR(20) NULL, [fil10] NVARCHAR(20) NULL, [izm01] SMALLINT NULL, [izm02] SMALLINT NULL, [izm03] SMALLINT NULL, [izm04] SMALLINT NULL, [izm05] SMALLINT NULL, [izm06] SMALLINT NULL, [izm07] SMALLINT NULL, [izm08] SMALLINT NULL, [izm09] SMALLINT NULL, [izm10] SMALLINT NULL, [izm11] SMALLINT NULL, [izm12] SMALLINT NULL, [izm13] SMALLINT NULL, [izm14] SMALLINT NULL, [izm15] SMALLINT NULL, [izm16] SMALLINT NULL, [izm17] SMALLINT NULL, [izm18] SMALLINT NULL, [izm19] SMALLINT NULL, [izm20] SMALLINT NULL, CONSTRAINT[FK_osmotr_ToStudent] FOREIGN KEY([idStudent]) REFERENCES Student ([Id]));";
                sqlCmd.ExecuteNonQuery();
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
                NoErrors = false;
            }
            // копируем шаблоны оценок, настроек и отчетов для новой базы с новыми именами
            // ОЦЕНКИ
            if (File.Exists(dbPath + "\\" + dbName + marksXML))
                MessageBox.Show("Файл оценок " + dbPath + "\\" + dbName + marksXML + " уже существует в данной папке.");
            else
            {
                try
                {
                    File.Copy(Path.Combine(dbPath, TemplateMARKS), Path.Combine(dbPath, dbName + marksXML)); // Файл не будет перезаписан, если существует
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Ошибка. Файл оценок не может быть создан!" + Environment.NewLine + "(" + ex.Message + ")");
                    NoErrors = false;
                }
            }
            // НАСТРОЙКИ
            if (File.Exists(dbPath + "\\" + dbName + settingsXML))
                MessageBox.Show("Файл оценок " + dbPath + "\\" + dbName + settingsXML + " уже существует в данной папке.");
            else
            {
                try
                {
                    File.Copy(Path.Combine(dbPath, TemplateSETTINGS), Path.Combine(dbPath, dbName + settingsXML)); // Файл не будет перезаписан, если существует
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Ошибка. Файл настроек не может быть создан!" + Environment.NewLine + "(" + ex.Message + ")");
                    NoErrors = false;
                }
            }
            // Паспорт
            if (File.Exists(dbPath + "\\" + dbName + passportDOCX))
                MessageBox.Show("Файл отчета " + dbPath + "\\" + dbName + passportDOCX + " уже существует в данной папке.");
            else
            {
                try
                {
                    File.Copy(Path.Combine(dbPath, TemplateRTF01), Path.Combine(dbPath, dbName + passportDOCX)); // Файл не будет перезаписан, если существует
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Ошибка. Шаблон отчета не может быть создан!" + Environment.NewLine + "(" + ex.Message + ")");
                    NoErrors = false;
                }
            }
            // Отчет1
            if (File.Exists(dbPath + "\\" + dbName + report01RTF))
                MessageBox.Show("Файл отчета " + dbPath + "\\" + dbName + report01RTF + " уже существует в данной папке.");
            else
            {
                try
                {
                    File.Copy(Path.Combine(dbPath, TemplateRTF02), Path.Combine(dbPath, dbName + report01RTF)); // Файл не будет перезаписан, если существует
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Ошибка. Шаблон отчета не может быть создан!" + Environment.NewLine + "(" + ex.Message + ")");
                    NoErrors = false;
                }
            }
            if (NoErrors)
                MessageBox.Show("Успешно создана новая база: \n" +
                    "\n" +
                    "Создан файл " + dbPath + "\\" + dbFileName + "\n"+
                    "Создан файл " + dbPath + "\\" + dbName + marksXML + "\n"+
                    "Создан файл " + dbPath + "\\" + dbName + settingsXML + "\n"+
                    "Создан файл " + dbPath + "\\" + dbName + passportDOCX + "\n"+
                    "Создан файл " + dbPath + "\\" + dbName + report01RTF + "\n"+
                    "\n" +
                    "Нажмите ОК для продолжения.");
        }// bdFileCreate()

        // **********************************************************************************************
        private void bdFileConnect()
        {
            if (!File.Exists(dbFileName))
                MessageBox.Show("Файл " + dbFileName + "не существует.");
            try
            {
                dbConn = new SQLiteConnection("Data Source=" + dbPath + "\\" + dbFileName + ";Version=3;");
                dbName = dbFileName.Substring(0, dbFileName.Length-7); // запомним имя без разрешения .sqlite
                dbConn.Open();
                myDataSet = new DataSet();
                // адаптер
                myDataAdapter = new SQLiteDataAdapter("Select* From Student", dbConn);
                // читаем
                myDataAdapter.Fill(myDataSet);
                //для изменения
                myDataAdapter.UpdateCommand = new SQLiteCommand("UPDATE Student SET passport=@passport, f=@f, i=@i, o=@o, sex=@sex, born=@born WHERE Id=@Id", dbConn);
                myDataAdapter.UpdateCommand.Parameters.Add("@Id", DbType.Int32, 10, "Id");
                myDataAdapter.UpdateCommand.Parameters.Add("@passport", DbType.StringFixedLength, 10, "passport");
                myDataAdapter.UpdateCommand.Parameters.Add("@f", DbType.String, 20, "f");
                myDataAdapter.UpdateCommand.Parameters.Add("@i", DbType.String, 20, "i");
                myDataAdapter.UpdateCommand.Parameters.Add("@o", DbType.String, 20, "o");
                myDataAdapter.UpdateCommand.Parameters.Add("@sex", DbType.String, 1, "sex");
                myDataAdapter.UpdateCommand.Parameters.Add("@born", DbType.DateTime, 10, "born");
                //для добавления
                myDataAdapter.InsertCommand = new SQLiteCommand("insert into Student ([Id],[passport],[f],[i],[o],[sex],[born]) values(@Id,@passport,@f,@i,@o,@sex,@born)", dbConn);
                myDataAdapter.InsertCommand.Parameters.Add("@Id", DbType.Int32, 10, "Id");
                myDataAdapter.InsertCommand.Parameters.Add("@passport", DbType.StringFixedLength, 10, "passport");
                myDataAdapter.InsertCommand.Parameters.Add("@f", DbType.String, 20, "f");
                myDataAdapter.InsertCommand.Parameters.Add("@i", DbType.String, 20, "i");
                myDataAdapter.InsertCommand.Parameters.Add("@o", DbType.String, 20, "o");
                myDataAdapter.InsertCommand.Parameters.Add("@sex", DbType.String, 1, "sex");
                myDataAdapter.InsertCommand.Parameters.Add("@born", DbType.DateTime, 10, "born");
                //для удаления
                myDataAdapter.DeleteCommand = new SQLiteCommand("DELETE FROM Student WHERE Id=@Id", dbConn);
                myDataAdapter.DeleteCommand.Parameters.Add("@Id", DbType.Int32, 10, "Id");
                // вешаем адаптер
                dataGridViewStudent.DataSource = myDataSet.Tables[0].DefaultView;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        } //bdFileConnect()

        // **********************************************************************************************
        private void dataGridViewStudent_SelectionChanged(object sender, EventArgs e)
        {
            // может почемуто запуститься при создании базы, когда настройки еще не загружены и NamesUpdate() выдаст ошибку
            // по этому проверяем перед запуском
            if ((dataGridViewStudent.CurrentCell != null)&&(MyFilters != null))
            {
                try
                {
                    myDataSet2 = new DataSet();
                    string commandText = "SELECT * FROM MedCheck WHERE MedCheck.idStudent = @idStudent;";
                    SQLiteCommand cmd = new SQLiteCommand(commandText, dbConn);
                    SQLiteParameter Param = new SQLiteParameter("@idStudent", DbType.String);
                    Param.Value = dataGridViewStudent.CurrentRow.Cells[0].Value;
                    cmd.Parameters.Add(Param);
                    myDataAdapter2 = new SQLiteDataAdapter(cmd);
                    myDataAdapter2.Fill(myDataSet2);
                    //для изменения
                    myDataAdapter2.UpdateCommand = new SQLiteCommand("UPDATE MedCheck SET data=@data, fil01=@fil01, fil02=@fil02, fil03=@fil03, fil04=@fil04, fil05=@fil05, fil06=@fil06, fil07=@fil07, fil08=@fil08, fil09=@fil09, fil10=@fil10, izm01=@izm01, izm02=@izm02, izm03=@izm03, izm04=@izm04, izm05=@izm05, izm06=@izm06, izm07=@izm07, izm08=@izm08, izm09=@izm09, izm10=@izm10, izm11=@izm11, izm12=@izm12, izm13=@izm13, izm14=@izm14, izm15=@izm15, izm16=@izm16, izm17=@izm17, izm18=@izm18, izm19=@izm19, izm20=@izm20 WHERE Id=@Id", dbConn);
                    myDataAdapter2.UpdateCommand.Parameters.Add("@Id", DbType.Int32, 10, "Id");
                    //myDataAdapter2.UpdateCommand.Parameters.Add("@idStudent", DbType.Int32, 10, "idStudent");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@data", DbType.DateTime, 10, "data");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil01", DbType.String, 20, "fil01");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil02", DbType.String, 20, "fil02");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil03", DbType.String, 20, "fil03");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil04", DbType.String, 20, "fil04");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil05", DbType.String, 20, "fil05");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil06", DbType.String, 20, "fil06");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil07", DbType.String, 20, "fil07");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil08", DbType.String, 20, "fil08");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil09", DbType.String, 20, "fil09");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@fil10", DbType.String, 20, "fil10");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm01", DbType.Int16, 20, "izm01");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm02", DbType.Int16, 20, "izm02");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm03", DbType.Int16, 20, "izm03");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm04", DbType.Int16, 20, "izm04");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm05", DbType.Int16, 20, "izm05");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm06", DbType.Int16, 20, "izm06");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm07", DbType.Int16, 20, "izm07");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm08", DbType.Int16, 20, "izm08");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm09", DbType.Int16, 20, "izm09");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm10", DbType.Int16, 20, "izm10");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm11", DbType.Int16, 20, "izm11");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm12", DbType.Int16, 20, "izm12");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm13", DbType.Int16, 20, "izm13");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm14", DbType.Int16, 20, "izm14");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm15", DbType.Int16, 20, "izm15");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm16", DbType.Int16, 20, "izm16");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm17", DbType.Int16, 20, "izm17");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm18", DbType.Int16, 20, "izm18");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm19", DbType.Int16, 20, "izm19");
                    myDataAdapter2.UpdateCommand.Parameters.Add("@izm20", DbType.Int16, 20, "izm20");
                    //для удаления
                    myDataAdapter2.DeleteCommand = new SQLiteCommand("DELETE FROM MedCheck WHERE Id=@Id", dbConn);
                    myDataAdapter2.DeleteCommand.Parameters.Add("@Id", DbType.Int32, 10, "Id");
                    // вешаем адаптер
                    medCheckDataGridView.DataSource = myDataSet2.Tables[0].DefaultView;
                    NamesUpdate();
                }
                catch (SQLiteException ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
            } // check for null

        } //StudentChanged()

        //===============================================================================================================
        // Сохранить таблицу медосмотров
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                medCheckDataGridView.EndEdit();
                myDataAdapter2.Update(myDataSet2.Tables[0]);
                myDataSet2.AcceptChanges();
                FlagCheck = false;
                button11.Visible = false;
                dataGridViewStudent.Enabled = true;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }

        }

        //===============================================================================================================
        // Сохранить таблицу обследуемых
        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridViewStudent.EndEdit();
                myDataAdapter.Update(myDataSet.Tables[0]);
                myDataSet.AcceptChanges();
                FlagStud = false;
                button12.Visible = false;
                medCheckDataGridView.Enabled = true;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        //===============================================================================================================
        // При смене строки в настройках ФИЛЬТРЫ ------------------------------------------------------------
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell != null)
            {
                groupBox1.Visible = false; // запретим редактирование
                int SelectedFilter = -1;
                //listBox1.Items.Clear();
                SelectedFilter = dataGridView1.CurrentCell.RowIndex;
                switch (SelectedFilter)
                {
                    case 0:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues1.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues1.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues1;
                        break;
                    case 1:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues2.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues2.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues2;
                        break;
                    case 2:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues3.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues3.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues3;
                        break;
                    case 3:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues4.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues4.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues4;
                        break;
                    case 4:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues5.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues5.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues5;
                        break;
                    case 5:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues6.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues6.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues6;
                        break;
                    case 6:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues7.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues7.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues7;
                        break;
                    case 7:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues8.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues8.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues8;
                        break;
                    case 8:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues9.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues9.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues9;
                        break;
                    case 9:
                        listBox1.Visible = true;
                        listBox1.DisplayMember = MyFilterValues10.Columns[0].ColumnName;
                        listBox1.ValueMember = MyFilterValues10.Columns[0].ColumnName;
                        listBox1.DataSource = MyFilterValues10;
                        break;
                    default:
                        listBox1.Visible = false;
                        break;
                }
                listBox1.Refresh();
            }  // check for null
        } //of SelectionChanged

        //===============================================================================================================
        // При смене строки в таблице ОЦЕНКИ ------------------------------------------------------------

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            int SelectedFilter = -1;
            String zznak, name_izm1, name_izm2, name_sex;
            int nomer_izm1, nomer_izm2, multiplayer;
            if (dataGridView3.CurrentCell != null)
            {
                SelectedFilter = dataGridView3.CurrentCell.RowIndex;
                if (MySpisok.Rows[SelectedFilter]["вкл"].Equals(true))
                {
                    groupBox2.Visible = true; // откроем правую панель
                    sposob_label.Text = "Способ расчета: ";
                    zznak = MySpisok.Rows[SelectedFilter]["znak"].ToString();
                    nomer_izm1 = Convert.ToInt32(MySpisok.Rows[SelectedFilter]["izm1"].ToString());
                    name_izm1 = MyIzm.Rows[nomer_izm1]["название"].ToString();
                    nomer_izm2 = Convert.ToInt32(MySpisok.Rows[SelectedFilter]["izm2"].ToString());
                    name_izm2 = MyIzm.Rows[nomer_izm2]["название"].ToString();
                    multiplayer = Convert.ToInt32(MySpisok.Rows[SelectedFilter]["mult"].ToString());
                    name_sex = MySpisok.Rows[SelectedFilter]["пол"].ToString();
                    switch (zznak)
                    {
                        case "0": //копируется
                            sposob_label.Text += "копируется из " + name_izm1;
                            break;
                        case "1":// вычисляется +
                            sposob_label.Text += "вычисляется из " + name_izm1 + " + " + name_izm2;
                            break;
                        case "2":// вычисляется -
                            sposob_label.Text += "вычисляется из " + name_izm1 + " - " + name_izm2;
                            break;
                        case "3":// вычисляется *
                            sposob_label.Text += "вычисляется из " + name_izm1 + " * " + name_izm2;
                            break;
                        case "4":// вычисляется /
                            sposob_label.Text += "вычисляется из " + name_izm1 + " / " + name_izm2;
                            break;
                        default:
                            // неверное значение
                            break;
                    }
                    // умножитель
                    if (multiplayer > 1)
                        sposob_label.Text += " x " + multiplayer;
                    // пол
                    switch (name_sex)
                    {
                        case "М": //копируется
                            sposob_label.Text += " (мужчины)";
                            break;
                        case "м":// вычисляется +
                            sposob_label.Text += " (мужчины)";
                            break;
                        case "Ж":// вычисляется -
                            sposob_label.Text += " (женщины)";
                            break;
                        case "ж":// вычисляется *
                            sposob_label.Text += " (женщины)";
                            break;
                        default:
                            // неверное значение
                            sposob_label.Text += ", в оценке неверно задан пол (Ж или М)";
                            break;
                    }
                    // рисуем график
                    string name_y_osi = MySpisok.Rows[SelectedFilter]["название"].ToString() + " (" + MySpisok.Rows[SelectedFilter]["ед.изм."].ToString() + ")";
                    string x_csv = MyGraphs.Rows[SelectedFilter]["x"].ToString();
                    string y_csv = MyGraphs.Rows[SelectedFilter]["y"].ToString();
                    string[] xx = x_csv.Split(',');
                    string[] yy = y_csv.Split(',');
                    DrawGraph(name_y_osi, xx, yy);
                }
                else
                    groupBox2.Visible = false; // закроем правую панель
            } // check for null
            TotalMark_label.Text = "Максимальные общие оценки: Муж=" + TotalBall("М") + " Жен=" + TotalBall("Ж");
            dot_label.Text = "";
        } //of SelectionChanged

        //===============================================================================================================
        // Здесь инициируются две таблицы настроек и объединяющий их набор-----------------------------------------
        public void InitSettingsDataSet(bool DefaultFlag)
        {
            MySettings = new DataSet();
            // Настройки фильтров
            MyFilters = new System.Data.DataTable();
            MyFilters.Columns.Add(new DataColumn("№", typeof(Int32)));
            MyFilters.Columns.Add(new DataColumn("вкл", typeof(bool)));
            MyFilters.Columns.Add(new DataColumn("название", typeof(String)));
            if (DefaultFlag)
            {
                MyFilters.Rows.Add(1, true, "Курс");
                MyFilters.Rows.Add(2, true, "Специализация");
                MyFilters.Rows.Add(3, true, "Институт");
                MyFilters.Rows.Add(4, true, "Группа");
                MyFilters.Rows.Add(5, false, "");
                MyFilters.Rows.Add(6, false, "");
                MyFilters.Rows.Add(7, false, "");
                MyFilters.Rows.Add(8, false, "");
                MyFilters.Rows.Add(9, false, "");
                MyFilters.Rows.Add(10, false, "");
            }
            MySettings.Tables.Add(MyFilters);
            MySettings.Tables[0].TableName = "Filters"; // XML key

            // Настройки измерений
            MyIzm = new System.Data.DataTable();
            MyIzm.Columns.Add(new DataColumn("№", typeof(Int32)));
            MyIzm.Columns.Add(new DataColumn("вкл", typeof(bool)));
            MyIzm.Columns.Add(new DataColumn("название", typeof(String)));
            MyIzm.Columns.Add(new DataColumn("ед.изм.", typeof(String)));
            if (DefaultFlag)
            {
                MyIzm.Rows.Add(1, true, "Рост", "см");
                MyIzm.Rows.Add(2, true, "Вес", "кг");
                MyIzm.Rows.Add(3, true, "АД верхнее", "мм рт.ст.");
                MyIzm.Rows.Add(4, true, "АД нижнее", "мм рт.ст.");
                MyIzm.Rows.Add(5, true, "Спирометрия", "куб.см");
                MyIzm.Rows.Add(6, true, "Динамометрия", "кг");
                MyIzm.Rows.Add(7, true, "Пульс до", "уд/мин");
                MyIzm.Rows.Add(8, true, "Пульс после", "уд/мин");
                MyIzm.Rows.Add(9, true, "Время восст.", "сек");
                MyIzm.Rows.Add(10, true, "Гибкость", "см");
                MyIzm.Rows.Add(11, false, "", "");
                MyIzm.Rows.Add(12, false, "", "");
                MyIzm.Rows.Add(13, false, "", "");
                MyIzm.Rows.Add(14, false, "", "");
                MyIzm.Rows.Add(15, false, "", "");
                MyIzm.Rows.Add(16, false, "", "");
                MyIzm.Rows.Add(17, false, "", "");
                MyIzm.Rows.Add(18, false, "", "");
                MyIzm.Rows.Add(19, false, "", "");
                MyIzm.Rows.Add(20, false, "", "");
            }
            MySettings.Tables.Add(MyIzm);
            MySettings.Tables[1].TableName = "Izm"; // XML key

            // Настройки значений фильтра1 - курс
            MyFilterValues1 = new System.Data.DataTable();
            MyFilterValues1.Columns.Add(new DataColumn("value", typeof(String)));
            if (DefaultFlag)
            {
                MyFilterValues1.Rows.Add("1");
                MyFilterValues1.Rows.Add("2");
                MyFilterValues1.Rows.Add("3");
                MyFilterValues1.Rows.Add("4");
                MyFilterValues1.Rows.Add("5");
            }
            MySettings.Tables.Add(MyFilterValues1);
            MySettings.Tables[2].TableName = "FilterValues1"; // XML key

            // Настройки значений фильтра2 - группа
            MyFilterValues2 = new System.Data.DataTable();
            MyFilterValues2.Columns.Add(new DataColumn("value", typeof(String)));
            // тут пусто по умолчанию, групп сильно много
            MySettings.Tables.Add(MyFilterValues2);
            MySettings.Tables[3].TableName = "FilterValues2"; // XML key

            // Настройки значений фильтра3 - факультет
            MyFilterValues3 = new System.Data.DataTable();
            MyFilterValues3.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues3);
            MySettings.Tables[4].TableName = "FilterValues3"; // XML key

            // Настройки значений фильтра4 - тренер
            MyFilterValues4 = new System.Data.DataTable();
            MyFilterValues4.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues4);
            MySettings.Tables[5].TableName = "FilterValues4"; // XML key

            // Настройки значений фильтра5 - факультет
            MyFilterValues5 = new System.Data.DataTable();
            MyFilterValues5.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues5);
            MySettings.Tables[6].TableName = "FilterValues5"; // XML key

            // Настройки значений фильтра6
            MyFilterValues6 = new System.Data.DataTable();
            MyFilterValues6.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues6);
            MySettings.Tables[7].TableName = "FilterValues6"; // XML key

            // Настройки значений фильтра7
            MyFilterValues7 = new System.Data.DataTable();
            MyFilterValues7.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues7);
            MySettings.Tables[8].TableName = "FilterValues7"; // XML key

            // Настройки значений фильтра8
            MyFilterValues8 = new System.Data.DataTable();
            MyFilterValues8.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues8);
            MySettings.Tables[9].TableName = "FilterValues8"; // XML key

            // Настройки значений фильтра9
            MyFilterValues9 = new System.Data.DataTable();
            MyFilterValues9.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues9);
            MySettings.Tables[10].TableName = "FilterValues9"; // XML key

            // Настройки значений фильтра10
            MyFilterValues10 = new System.Data.DataTable();
            MyFilterValues10.Columns.Add(new DataColumn("value", typeof(String)));
            // пока тоже пусто
            MySettings.Tables.Add(MyFilterValues10);
            MySettings.Tables[11].TableName = "FilterValues10"; // XML key

        }

        // of InitSettingsDataSet---------------------------------------------------------------------------------------------

        //===============================================================================================================
        // Здесь инициируются две таблицы настроек и объединяющий их набор-----------------------------------------
        public void InitMarksDataSet(bool DefaultFlag)
        {
            MyMarks = new DataSet();
            // Список оценок - главная таблица
            MySpisok = new System.Data.DataTable();
            MySpisok.Columns.Add(new DataColumn("вкл", typeof(bool)));
            MySpisok.Columns.Add(new DataColumn("название", typeof(String)));
            MySpisok.Columns.Add(new DataColumn("ед.изм.", typeof(String)));
            MySpisok.Columns.Add(new DataColumn("izm1", typeof(Int32)));
            MySpisok.Columns.Add(new DataColumn("izm2", typeof(Int32)));
            MySpisok.Columns.Add(new DataColumn("znak", typeof(String)));
            MySpisok.Columns.Add(new DataColumn("mult", typeof(String)));
            MySpisok.Columns.Add(new DataColumn("пол", typeof(String)));
            if (DefaultFlag)
            {
                MySpisok.Rows.Add(true, "Рост", "см.", 0, 0, 0, 1, "м");
                MySpisok.Rows.Add(true, "Спирометрия", "см3", 4, 0, 0, 1, "м");
            }
            MyMarks.Tables.Add(MySpisok);
            MyMarks.Tables[0].TableName = "Spisok"; // XML key

            // графики - списки точек
            MyGraphs = new System.Data.DataTable();
            MyGraphs.Columns.Add(new DataColumn("x", typeof(String)));
            MyGraphs.Columns.Add(new DataColumn("y", typeof(String)));
            if (DefaultFlag)
            {
                MyGraphs.Rows.Add("0,4,8,12,16,20,16,12,8,4,0", "100,115,130,145,160,175,190,205,220,235,250");
                MyGraphs.Rows.Add("0,2,4,6,8,10,12,14,16,18,20", "1000,1900,2800,3700,4600,5500,6400,7300,8200,9100,10000");
            }
            MyMarks.Tables.Add(MyGraphs);
            MyMarks.Tables[1].TableName = "MyGraphs"; // XML key
            TotalMark_label.Text = "Максимальные общие оценки: Муж=" + TotalBall("М") + " Жен=" + TotalBall("Ж");
            dot_label.Text = "";
        } // of InitMarksDataSet

        //===============================================================================================================
        // Появились изменения - включаем флаги и кнопки сохранения
        private void dataGridViewStudent_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            FlagStud = true;
            button12.Visible = true;
            medCheckDataGridView.Enabled = false;
        }

        //===============================================================================================================
        // Появились изменения - включаем флаги и кнопки сохранения
        private void medCheckDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            FlagCheck = true;
            button11.Visible = true;
            dataGridViewStudent.Enabled = false;
        }

        //===============================================================================================================
        // Если есть изменения - спросим перед закрытием программы
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (FlagStud || FlagCheck || FlagMark || FlagIzm)
                if (MessageBox.Show("Есть несохраненные данные. Точно закрыть программу?", "Health", MessageBoxButtons.YesNo) == DialogResult.No)
                    e.Cancel = true;
        }

        //===============================================================================================================
        // Если есть изменения - спросим перед сменой вкладки
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (FlagStud || FlagCheck || FlagMark || FlagIzm)
                if (MessageBox.Show("Есть несохраненные данные. Точно закрыть вкладку?", "Health", MessageBoxButtons.YesNo) == DialogResult.No)
                    e.Cancel = true;
                else
                {
                    // на 1ой вкладке отмечаем, что изменений нет
                    FlagStud = false;
                    FlagCheck = false;
                    FlagMark = false;
                    FlagIzm = false;
                    button11.Visible = false;
                    button12.Visible = false;
                    dataGridViewStudent.Enabled = true;
                    medCheckDataGridView.Enabled = true;
                }
        }
        //===============================================================================================================
        // Отчет N 3
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            // заряжаем в таблицу названия измерений ( для формы7)
            Form7.Izm1table = MyIzm;
            Form7.Izm2table = MySpisok;
            // заряжаем в таблицу названия полов ( для формы7)
            System.Data.DataTable tempIzm3 = new System.Data.DataTable();
            tempIzm3.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm3.Columns.Add(new DataColumn("name", typeof(String)));
            tempIzm3.Rows.Add(1, "мужчины");
            tempIzm3.Rows.Add(2, "женщины");
            Form7.Izm3table = tempIzm3;
            // копируем фильтры для отображения в форме6
            Form7.MyFilters = MyFilters;
            Form7.MyFilterValues1 = MyFilterValues1;
            Form7.MyFilterValues2 = MyFilterValues2;
            Form7.MyFilterValues3 = MyFilterValues3;
            Form7.MyFilterValues4 = MyFilterValues4;
            Form7.MyFilterValues5 = MyFilterValues5;
            Form7.MyFilterValues6 = MyFilterValues6;
            Form7.MyFilterValues7 = MyFilterValues7;
            Form7.MyFilterValues8 = MyFilterValues8;
            Form7.MyFilterValues9 = MyFilterValues9;
            Form7.MyFilterValues10 = MyFilterValues10;

            // ГОТОВО, открываем форму
            Form ConstrForm7 = new Form7();
            ConstrForm7.ShowDialog();
            if (form7Cancel == false) // если не нажималась отмена, продолжаем конструктор
                Report2(form7param3, form7param4, form7param5, form7param6, form7param7, form7param8, form7param9, form7param10, form7param11, form7param12, form7param13, form7param14, form7param15, form7param16, form7tab1, form7tab2);
        }

        //===============================================================================================================
        // Появились изменения - включаем флаги и кнопки сохранения
        private void medCheckDataGridView_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            FlagCheck = true;
            button11.Visible = true;
            dataGridViewStudent.Enabled = false;
        }

        //===============================================================================================================
        // Появились изменения - включаем флаги и кнопки сохранения
        private void dataGridViewStudent_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            FlagStud = true;
            button12.Visible = true;
            medCheckDataGridView.Enabled = false;
        }

        //===============================================================================================================
        // Появились изменения - включаем флаги и кнопки сохранения
        private void dataGridViewStudent_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            FlagStud = true;
            button12.Visible = true;
            medCheckDataGridView.Enabled = false;
        }

        //===============================================================================================================
        // при нажатии Enter ищем обследуемого по зачетке (документу)
        private void box1_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) && (box1.Text.Length > 0)) // нажата Enter
            {
                vvodDataSet = new DataSet();
                string commandText = "SELECT * FROM Student WHERE Student.passport = @passport;";
                SQLiteCommand cmd = new SQLiteCommand(commandText, dbConn);
                SQLiteParameter Param = new SQLiteParameter("@passport", DbType.StringFixedLength);
                Param.Value = box1.Text;
                cmd.Parameters.Add(Param);
                vvodDataAdapter = new SQLiteDataAdapter(cmd);
                vvodDataAdapter.Fill(vvodDataSet);
                student_found = false;
                for (int i = 0; i < vvodDataSet.Tables[0].Rows.Count; i++)
                {
                    student_found = true;
                    ClearVvodForm();
                    id_found = vvodDataSet.Tables[0].Rows[i]["Id"].ToString(); // сохраним его Id (int->string)
                    box0.Text = vvodDataSet.Tables[0].Rows[i]["passport"].ToString();
                    box2.Text = vvodDataSet.Tables[0].Rows[i]["f"].ToString();
                    box3.Text = vvodDataSet.Tables[0].Rows[i]["i"].ToString();
                    box4.Text = vvodDataSet.Tables[0].Rows[i]["o"].ToString();
                    box5.Text = vvodDataSet.Tables[0].Rows[i]["sex"].ToString();
                    box6.Text = vvodDataSet.Tables[0].Rows[i]["born"].ToString();
                }
                if (student_found) // обследуемый НАЙДЕН
                {
                    label1.Text = "Обследуемый найден, введите дату и данные медосмотра...";
                    label1.Visible = true;
                    box1.Text = "";
                    //box0.Enabled = false;
                    box7.Visible = true;
                    label8.Visible = true;
                    groupBox3.Visible = true;
                    groupBox3.BackColor = Color.Transparent;
                    groupBox4.Visible = true;
                    groupBox5.Visible = true;
                    box7.Focus(); // курсор на дату осмотра
                    box7.Value = DateTime.Today;
                    button9.Visible = true;
                }
                else// обследуемый НЕ НАЙДЕН
                {
                    String vrem = box1.Text;
                    ClearVvodForm();
                    //box0.Enabled = true;
                    //box0.Enabled = false;
                    box0.Text = vrem;
                    label1.Text = "Обследуемый не найден, введите новые данные...";
                    label1.Visible = true;
                    box7.Visible = true;
                    label8.Visible = true;
                    groupBox3.Visible = true;
                    groupBox3.BackColor = SystemColors.AppWorkspace;
                    groupBox4.Visible = true;
                    groupBox5.Visible = true;
                    box2.Focus(); // курсор на фамилию
                    box6.Value = DateTime.Parse("01.01.2000");
                    box7.Value = DateTime.Today;
                    button9.Visible = true;
                }
            }
        }

        //===============================================================================================================
        public void SaveSettingsXML(bool DefaultFlag)
        {
            if (DefaultFlag)
            {
                InitSettingsDataSet(true); // с флагом сбросок настроек на умолчанию
                // Загруженные XML данные подключаем к DGVюшкам
                dataGridView1.DataSource = MyFilters;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView2.DataSource = MyIzm;
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            MySettings.WriteXml(dbPath + "\\" + dbName + settingsXML);
            MessageBox.Show("Настройки сохранены.");
        } //of SaveSettingsXML

        //===============================================================================================================

        public void SaveMarksXML(bool DefaultFlag)
        {
            if (DefaultFlag)
            {
                InitMarksDataSet(true); // с флагом сбросок настроек на умолчанию
                // Загруженные XML данные подключаем к DGVюшкам
                dataGridView3.DataSource = MySpisok;// Загруженные XML данные подключаем к DGVюшкам
                dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            MyMarks.WriteXml(dbPath + "\\" + dbName + marksXML);
            MessageBox.Show("Оценки сохранены.");
        } //of SaveMarksXML

        //===============================================================================================================
        // Сохранение данных из формы ввода в базу
        private void button9_Click(object sender, EventArgs e)
        {
            // Проверка на пустые поля
            bool empty = false;
            if (box2.Text.Length == 0) empty = true;
            if (box3.Text.Length == 0) empty = true;
            if (box4.Text.Length == 0) empty = true;
            if (box5.Text.Length == 0) empty = true;
            if ((box8.Visible) && (box8.Text.Length == 0)) empty = true;
            if ((box9.Visible) && (box9.Text.Length == 0)) empty = true;
            if ((box10.Visible) && (box10.Text.Length == 0)) empty = true;
            if ((box11.Visible) && (box11.Text.Length == 0)) empty = true;
            if ((box12.Visible) && (box12.Text.Length == 0)) empty = true;
            if ((box13.Visible) && (box13.Text.Length == 0)) empty = true;
            if ((box14.Visible) && (box14.Text.Length == 0)) empty = true;
            if ((box15.Visible) && (box15.Text.Length == 0)) empty = true;
            if ((box16.Visible) && (box16.Text.Length == 0)) empty = true;
            if ((box17.Visible) && (box17.Text.Length == 0)) empty = true;
            if ((box18.Visible) && (box18.Text.Length == 0)) empty = true;
            if ((box19.Visible) && (box19.Text.Length == 0)) empty = true;
            if ((box20.Visible) && (box20.Text.Length == 0)) empty = true;
            if ((box21.Visible) && (box21.Text.Length == 0)) empty = true;
            if ((box22.Visible) && (box22.Text.Length == 0)) empty = true;
            if ((box23.Visible) && (box23.Text.Length == 0)) empty = true;
            if ((box24.Visible) && (box24.Text.Length == 0)) empty = true;
            if ((box25.Visible) && (box25.Text.Length == 0)) empty = true;
            if ((box26.Visible) && (box26.Text.Length == 0)) empty = true;
            if ((box27.Visible) && (box27.Text.Length == 0)) empty = true;
            if ((box28.Visible) && (box28.Text.Length == 0)) empty = true;
            if ((box29.Visible) && (box29.Text.Length == 0)) empty = true;
            if ((box30.Visible) && (box30.Text.Length == 0)) empty = true;
            if ((box31.Visible) && (box31.Text.Length == 0)) empty = true;
            if ((box32.Visible) && (box32.Text.Length == 0)) empty = true;
            if ((box33.Visible) && (box33.Text.Length == 0)) empty = true;
            if ((box34.Visible) && (box34.Text.Length == 0)) empty = true;
            if ((box35.Visible) && (box35.Text.Length == 0)) empty = true;
            if ((box36.Visible) && (box36.Text.Length == 0)) empty = true;
            if ((box37.Visible) && (box37.Text.Length == 0)) empty = true;
            if (empty == true)
                if (MessageBox.Show("Некоторые поля не заполнены. Сохранить с пустыми значениями?", "Health", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    empty = false;
            if (empty == false)
            {
                // сохраняем
                string commandText2;
                SQLiteCommand cmd2;
                if (student_found) // обследуемый НАЙДЕН (определяется в box1_KeyDown)
                {
                    // обновим обследуемого...
                    try
                    {
                        commandText2 = "UPDATE Student SET f=@f,i=@i,o=@o,sex=@sex,born=@born WHERE passport=@passport";
                        cmd2 = new SQLiteCommand(commandText2, dbConn);
                        //cmd2.Parameters.AddWithValue("@Id", int.Parse(id_found));
                        cmd2.Parameters.AddWithValue("@passport", box0.Text);
                        cmd2.Parameters.AddWithValue("@f", box2.Text);
                        cmd2.Parameters.AddWithValue("@i", box3.Text);
                        cmd2.Parameters.AddWithValue("@o", box4.Text);
                        cmd2.Parameters.AddWithValue("@sex", box5.Text);
                        cmd2.Parameters.AddWithValue("@born", box6.Value);
                        cmd2.ExecuteNonQuery();
                        label1.Text = "Данные обследуемого обновлены";
                    }
                    catch (SQLiteException ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                }
                else// обследуемый НЕ НАЙДЕН
                {
                    // создадим обследуемого...
                    try
                    {
                        commandText2 = "insert into Student ([Id],[passport],[f],[i],[o],[sex],[born]) values(@Id,@passport,@f,@i,@o,@sex,@born)";
                        cmd2 = new SQLiteCommand(commandText2, dbConn);
                        cmd2.Parameters.AddWithValue("@Id", null); // autoincrement
                        cmd2.Parameters.AddWithValue("@passport", box0.Text);
                        cmd2.Parameters.AddWithValue("@f", box2.Text);
                        cmd2.Parameters.AddWithValue("@i", box3.Text);
                        cmd2.Parameters.AddWithValue("@o", box4.Text);
                        cmd2.Parameters.AddWithValue("@sex", box5.Text);
                        cmd2.Parameters.AddWithValue("@born", box6.Value);
                        cmd2.ExecuteNonQuery();
                        label1.Text = "Данные нового обследуемого добавлены";
                    }
                    catch (SQLiteException ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    // получим Id нового
                    vvodDataSet = new DataSet();
                    string commandText = "SELECT * FROM Student WHERE Student.passport = @passport;";
                    SQLiteCommand cmd = new SQLiteCommand(commandText, dbConn);
                    SQLiteParameter Param = new SQLiteParameter("@passport", DbType.StringFixedLength);
                    Param.Value = box0.Text;
                    cmd.Parameters.Add(Param);
                    vvodDataAdapter = new SQLiteDataAdapter(cmd);
                    vvodDataAdapter.Fill(vvodDataSet);
                    student_found = false;
                    for (int i = 0; i < vvodDataSet.Tables[0].Rows.Count; i++)
                    {
                        id_found = vvodDataSet.Tables[0].Rows[i]["Id"].ToString(); // сохраним его Id (int->string)
                    }
                } // if обследуемый найден или не найден

                // ... и добавим медосмотр
                try
                {
                    commandText2 = "insert into MedCheck ([Id],[idStudent],[data],[fil01],[fil02],[fil03],[fil04],[fil05],[fil06],[fil07],[fil08],[fil09],[fil10],[izm01],[izm02],[izm03],[izm04],[izm05],[izm06],[izm07],[izm08],[izm09],[izm10],[izm11],[izm12],[izm13],[izm14],[izm15],[izm16],[izm17],[izm18],[izm19],[izm20]) values(@Id,@idStudent,@data,@fil01,@fil02,@fil03,@fil04,@fil05,@fil06,@fil07,@fil08,@fil09,@fil10,@izm01,@izm02,@izm03,@izm04,@izm05,@izm06,@izm07,@izm08,@izm09,@izm10,@izm11,@izm12,@izm13,@izm14,@izm15,@izm16,@izm17,@izm18,@izm19,@izm20)";
                    cmd2 = new SQLiteCommand(commandText2, dbConn);
                    cmd2.Parameters.AddWithValue("@Id", null); // primary key auto_increment
                    cmd2.Parameters.AddWithValue("@idStudent", int.Parse(id_found));
                    cmd2.Parameters.AddWithValue("@data", box7.Value);
                    cmd2.Parameters.AddWithValue("@fil01", box8.Text);
                    cmd2.Parameters.AddWithValue("@fil02", box9.Text);
                    cmd2.Parameters.AddWithValue("@fil03", box10.Text);
                    cmd2.Parameters.AddWithValue("@fil04", box11.Text);
                    cmd2.Parameters.AddWithValue("@fil05", box12.Text);
                    cmd2.Parameters.AddWithValue("@fil06", box13.Text);
                    cmd2.Parameters.AddWithValue("@fil07", box14.Text);
                    cmd2.Parameters.AddWithValue("@fil08", box15.Text);
                    cmd2.Parameters.AddWithValue("@fil09", box16.Text);
                    cmd2.Parameters.AddWithValue("@fil10", box17.Text);
                    cmd2.Parameters.AddWithValue("@izm01", box18.Text);
                    cmd2.Parameters.AddWithValue("@izm02", box19.Text);
                    cmd2.Parameters.AddWithValue("@izm03", box20.Text);
                    cmd2.Parameters.AddWithValue("@izm04", box21.Text);
                    cmd2.Parameters.AddWithValue("@izm05", box22.Text);
                    cmd2.Parameters.AddWithValue("@izm06", box23.Text);
                    cmd2.Parameters.AddWithValue("@izm07", box24.Text);
                    cmd2.Parameters.AddWithValue("@izm08", box25.Text);
                    cmd2.Parameters.AddWithValue("@izm09", box26.Text);
                    cmd2.Parameters.AddWithValue("@izm10", box27.Text);
                    cmd2.Parameters.AddWithValue("@izm11", box28.Text);
                    cmd2.Parameters.AddWithValue("@izm12", box29.Text);
                    cmd2.Parameters.AddWithValue("@izm13", box30.Text);
                    cmd2.Parameters.AddWithValue("@izm14", box31.Text);
                    cmd2.Parameters.AddWithValue("@izm15", box32.Text);
                    cmd2.Parameters.AddWithValue("@izm16", box33.Text);
                    cmd2.Parameters.AddWithValue("@izm17", box34.Text);
                    cmd2.Parameters.AddWithValue("@izm18", box35.Text);
                    cmd2.Parameters.AddWithValue("@izm19", box36.Text);
                    cmd2.Parameters.AddWithValue("@izm20", box37.Text);
                    cmd2.ExecuteNonQuery();
                    label1.Text = label1.Text + " ,новый медосмотр добавлен.";
                    button9.Visible = false;
                    box7.Visible = false;
                    label8.Visible = false;
                    groupBox3.Visible = false;
                    groupBox4.Visible = false;
                    groupBox5.Visible = false;
                    box1.Focus(); // курсор на 1ый текстбокс
                    if (MessageBox.Show("Напечатать паспорт физического развития?", "Health", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    { 
                        PrintPassport(box0.Text);
                        tabControl1.SelectedTab = tabControl1.TabPages["tabPage3"]; // активировать вкладку "Отчеты"
                    }
                    ClearVvodForm();
                }
                catch (SQLiteException ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
            } // of dialog
        } // of save

        //===============================================================================================================
        //---------------------------------------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            SaveSettingsXML(true);
            FlagIzm = false;
            NamesUpdate();
        }

        //===============================================================================================================
        //---------------------------------------------------------------------------------------------------
        private void button2_Click(object sender, EventArgs e)
        {
            SaveSettingsXML(false);
            FlagIzm = false;
            NamesUpdate();
        }

        //===============================================================================================================
        //---------------------------------------------------------------------------------------------------
        private void button6_Click(object sender, EventArgs e)
        {
            SaveMarksXML(false);
            FlagMark = false;
            groupBox2.Visible = false; // закроем правую панель
        }

        //===============================================================================================================
        //---------------------------------------------------------------------------------------------------
        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView3.CurrentCell = dataGridView3[0, 0]; // уюерем фокус с удаляемых строк, иначе ошибка
            SaveMarksXML(true);
            FlagMark = false;
            groupBox2.Visible = false; // закроем правую панель
        }

        //===============================================================================================================
        //---------------------------------------------------------------------------------------------------
        public void LoadSettingsXML()
        {
            // читаем сразу весь набор и таблицы-----------------------------
            try
            {
                //open the file using a Stream
                using (Stream stream = new FileStream(dbPath + "\\" + dbName + settingsXML, FileMode.Open, FileAccess.Read))
                {
                    //use ReadXml to read the XML stream
                    MySettings.ReadXml(stream);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка. "+ dbPath + "\\" + dbName + settingsXML + " не загружен!" + Environment.NewLine + "(" + ex.Message + ")");
            }
        } //of LoadSettingsXML

        //===============================================================================================================
        // Появились изменения в Фильтрах - включаем флаг
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            FlagIzm = true;
        }

        //===============================================================================================================
        // Появились изменения в Измерениях - включаем флаг
        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            FlagIzm = true;
        }

        //===============================================================================================================
        // Меню ОТЧЕТЫ
        private void tabControl1_MouseClick(object sender, MouseEventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"]) // вкладка "Отчеты"
                if (e.Button == MouseButtons.Left)
                {
                    this.contextMenuStrip1.Show(this.tabControl1, e.Location);
                }
        }

        //===============================================================================================================
        // Форма отчета ПАСПОРТ ФР
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form ConstrForm5 = new Form5();
            ConstrForm5.ShowDialog();
            if (form5Cancel == false) // если не нажималась отмена, продолжаем конструктор
                PrintPassport(form5param1);
        }

        //===============================================================================================================
        //---------------------------------------------------------------------------------------------------
        public void LoadMarksXML()
        {
            // читаем сразу весь набор, все таблицы-----------------------------
            try
            {
                //open the file using a Stream
                using (Stream stream = new FileStream(dbPath + "\\" + dbName + marksXML, FileMode.Open, FileAccess.Read))
                {
                    //use ReadXml to read the XML stream
                    MyMarks.ReadXml(stream);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка. "+ dbPath + "\\" + dbName + marksXML + " не загружен!" + Environment.NewLine + "(" + ex.Message + ")");
            }
        } //of LoadMarksXML

        // НОВОЕ ЗНАЧЕНИЕ ИЗ СТРОКИ РЕДАКТИРОВАНИЯ В ТАБЛИЦУ LISTBOXА (ENTER)
        //===============================================================================================================
        private void button4_Click(object sender, EventArgs e)
        {
            listbox1index = listBox1.SelectedIndex; // сохраним
            if (listBox1.SelectedIndex >= 0)
            {
                DataRow newRow = TempFilterValues.NewRow();
                newRow[0] = filter_textBox.Text;
                TempFilterValues.Rows[listBox1.SelectedIndex].Delete();
                TempFilterValues.Rows.InsertAt(newRow, listbox1index);
                FlagIzm = true;
                listBox1.SelectedIndex = listbox1index; // восстановим
            }
        }

        // ОТОБРАЖАЕТ ЗНАЧЕНИЕ ТЕКУЩЕЙ ВЫБРАННОЙ ИЗ LISTBOXА СТРОКИ В ОКОШКЕ ДЛЯ РЕДАКТИРОВАНИЯ
        //===============================================================================================================
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 4 && listBox1.SelectedIndex > -1) // вызываем только при активной 4ой вкладке, иначе ошибка
            {
                // определим, какой из 10ти фильтров выбран
                groupBox1.Visible = true; // разрешим редактирование
                int SelectedFilter = -1;
                SelectedFilter = dataGridView1.CurrentCell.RowIndex;
                switch (SelectedFilter)
                {
                    case 0:
                        TempFilterValues = MyFilterValues1;
                        break;
                    case 1:
                        TempFilterValues = MyFilterValues2;
                        break;
                    case 2:
                        TempFilterValues = MyFilterValues3;
                        break;
                    case 3:
                        TempFilterValues = MyFilterValues4;
                        break;
                    case 4:
                        TempFilterValues = MyFilterValues5;
                        break;
                    case 5:
                        TempFilterValues = MyFilterValues6;
                        break;
                    case 6:
                        TempFilterValues = MyFilterValues7;
                        break;
                    case 7:
                        TempFilterValues = MyFilterValues8;
                        break;
                    case 8:
                        TempFilterValues = MyFilterValues9;
                        break;
                    case 9:
                        TempFilterValues = MyFilterValues10;
                        break;
                    default:
                        groupBox1.Visible = false;
                        listBox1.Visible = false;
                        break;
                }
                // textBox1.Text = TempFilterValues.Rows[listBox1.SelectedIndex].ItemArray[0].ToString();
            }
        }

        //===============================================================================================================
        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (listBox1.SelectedIndex > -1)
                {
                    // определим, какой из 10ти фильтров выбран
                    int SelectedFilter = -1;
                    SelectedFilter = dataGridView1.CurrentCell.RowIndex;
                    switch (SelectedFilter)
                    {
                        case 0:
                            TempFilterValues = MyFilterValues1;
                            break;
                        case 1:
                            TempFilterValues = MyFilterValues2;
                            break;
                        case 2:
                            TempFilterValues = MyFilterValues3;
                            break;
                        case 3:
                            TempFilterValues = MyFilterValues4;
                            break;
                        case 4:
                            TempFilterValues = MyFilterValues5;
                            break;
                        case 5:
                            TempFilterValues = MyFilterValues6;
                            break;
                        case 6:
                            TempFilterValues = MyFilterValues7;
                            break;
                        case 7:
                            TempFilterValues = MyFilterValues8;
                            break;
                        case 8:
                            TempFilterValues = MyFilterValues9;
                            break;
                        case 9:
                            TempFilterValues = MyFilterValues10;
                            break;
                        default:
                            listBox1.Visible = false;
                            break;
                    }
                    listbox1moving = true;
                    listbox1index = listBox1.SelectedIndex;
                    hidden_label.Visible = true;
                    hidden_label.Location = new System.Drawing.Point(listBox1.Location.X + e.X + listbox1offset, listBox1.Location.Y + e.Y + listbox1offset);
                    hidden_label.Text = TempFilterValues.Rows[listbox1index].ItemArray[0].ToString();                     // hidden_label.Text = listBox1.SelectedItem.ToString();
                    filter_textBox.Text = TempFilterValues.Rows[listbox1index].ItemArray[0].ToString();
                    sw.Start();
                }
            }
            catch (Exception) { }
        } //of listBox1_MouseDown

        //===============================================================================================================
        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
            //      try
            //    {
            sw.Stop();
            if (listbox1moving && listbox1index > -1 && sw.ElapsedMilliseconds > 100 && listBox1.SelectedIndex != listbox1index) //стопка условий чтоб наверняка лишнее не подвинуть
            {
                // сохраняем строку
                DataRow newRow = TempFilterValues.NewRow(); // string temp = listBox1.Items[listbox1index].ToString();
                newRow[0] = TempFilterValues.Rows[listbox1index].ItemArray[0]; // .ToString()
                                                                               // удаляем строку
                TempFilterValues.Rows[listbox1index].Delete(); //listBox1.Items.RemoveAt(listbox1index);
                if (listBox1.SelectedIndex < listbox1index)
                {
                    TempFilterValues.Rows.InsertAt(newRow, listBox1.SelectedIndex); //listBox1.Items.Insert(listBox1.SelectedIndex, temp);
                    listBox1.SelectedIndex = listBox1.SelectedIndex - 1;
                }
                else
                {
                    TempFilterValues.Rows.InsertAt(newRow, listBox1.SelectedIndex + 1); //listBox1.Items.Insert(listBox1.SelectedIndex + 1, temp);
                    listBox1.SelectedIndex = listBox1.SelectedIndex + 1;
                }
                // определим, какой из 10ти фильтров выбран
                int SelectedFilter = -1;
                SelectedFilter = dataGridView1.CurrentCell.RowIndex;
                switch (SelectedFilter)
                {
                    case 0:
                        MyFilterValues1 = TempFilterValues;
                        break;
                    case 1:
                        MyFilterValues2 = TempFilterValues;
                        break;
                    case 2:
                        MyFilterValues3 = TempFilterValues;
                        break;
                    case 3:
                        MyFilterValues4 = TempFilterValues;
                        break;
                    case 4:
                        MyFilterValues5 = TempFilterValues;
                        break;
                    case 5:
                        MyFilterValues6 = TempFilterValues;
                        break;
                    case 6:
                        MyFilterValues7 = TempFilterValues;
                        break;
                    case 7:
                        MyFilterValues8 = TempFilterValues;
                        break;
                    case 8:
                        MyFilterValues9 = TempFilterValues;
                        break;
                    case 9:
                        MyFilterValues10 = TempFilterValues;
                        break;
                    default:
                        listBox1.Visible = false;
                        break;
                }
            }
            //обнуляем все
            sw.Reset();
            listbox1index = -1;
            listbox1moving = false;
            hidden_label.Visible = false;

        } // of listBox1_MouseUp

        //===============================================================================================================
        private void listBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (hidden_label.Visible == true) hidden_label.Location = new System.Drawing.Point(listBox1.Location.X + e.X + listbox1offset, listBox1.Location.Y + e.Y + listbox1offset);
        } // of listBox1_MouseMove

        // ДОБАВИТЬ ПУСТОЕ ЗНАЧЕНИЕ ФИЛЬТРА В LISTBOX
        //===============================================================================================================
        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.Visible)
            {
                listbox1index = listBox1.SelectedIndex; // где курсор
                if (listbox1index < 0) listbox1index = 0; // если не нашли, ставим в начало
                // определим, какой из 10ти фильтров выбран
                int SelectedFilter = -1;
                SelectedFilter = dataGridView1.CurrentCell.RowIndex;
                switch (SelectedFilter)
                {
                    case 0:
                        MyFilterValues1.Rows.InsertAt(MyFilterValues1.NewRow(), listbox1index);
                        break;
                    case 1:
                        MyFilterValues2.Rows.InsertAt(MyFilterValues2.NewRow(), listbox1index);
                        break;
                    case 2:
                        MyFilterValues3.Rows.InsertAt(MyFilterValues3.NewRow(), listbox1index);
                        break;
                    case 3:
                        MyFilterValues4.Rows.InsertAt(MyFilterValues4.NewRow(), listbox1index);
                        break;
                    case 4:
                        MyFilterValues5.Rows.InsertAt(MyFilterValues5.NewRow(), listbox1index);
                        break;
                    case 5:
                        MyFilterValues6.Rows.InsertAt(MyFilterValues6.NewRow(), listbox1index);
                        break;
                    case 6:
                        MyFilterValues7.Rows.InsertAt(MyFilterValues7.NewRow(), listbox1index);
                        break;
                    case 7:
                        MyFilterValues8.Rows.InsertAt(MyFilterValues8.NewRow(), listbox1index);
                        break;
                    case 8:
                        MyFilterValues9.Rows.InsertAt(MyFilterValues9.NewRow(), listbox1index);
                        break;
                    case 9:
                        MyFilterValues10.Rows.InsertAt(MyFilterValues10.NewRow(), listbox1index);
                        break;
                    default:
                        listBox1.Visible = false;
                        break;
                }
                FlagIzm = true;
                // курсор назад на новую строку
                listBox1.SelectedIndex = listbox1index;

            }
        }

        //===============================================================================================================
        // УДАЛИТЬ ЗНАЧЕНИЕ ФИЛЬТРА---------------------------------------------------------------------------------------------
        private void button5_Click(object sender, EventArgs e)
        {
            listbox1index = listBox1.SelectedIndex; // где курсор
            if (listbox1index >= 0) // если не нашли - будет ошибка
            {
                // определим, какой из 10ти фильтров выбран
                int SelectedFilter = -1;
                SelectedFilter = dataGridView1.CurrentCell.RowIndex;
                switch (SelectedFilter)
                {
                    case 0:
                        MyFilterValues1.Rows[listbox1index].Delete();
                        break;
                    case 1:
                        MyFilterValues2.Rows[listbox1index].Delete();
                        break;
                    case 2:
                        MyFilterValues3.Rows[listbox1index].Delete();
                        break;
                    case 3:
                        MyFilterValues4.Rows[listbox1index].Delete();
                        break;
                    case 4:
                        MyFilterValues5.Rows[listbox1index].Delete();
                        break;
                    case 5:
                        MyFilterValues6.Rows[listbox1index].Delete();
                        break;
                    case 6:
                        MyFilterValues7.Rows[listbox1index].Delete();
                        break;
                    case 7:
                        MyFilterValues8.Rows[listbox1index].Delete();
                        break;
                    case 8:
                        MyFilterValues9.Rows[listbox1index].Delete();
                        break;
                    case 9:
                        MyFilterValues10.Rows[listbox1index].Delete();
                        break;
                    default:
                        listBox1.Visible = false;
                        break;
                }
                FlagIzm = true;
            }
            // курсор назад на новую строку
            //listBox1.SelectedIndex = listbox1index;

        }

        //===============================================================================================================
        // ОБНОВИТЬ ЗНАЧЕНИЯ ФИЛЬТРОВ И ИЗМЕРЕНИЙ ВСООТВЕТСТВИИ С НАСТРОЙКАМИ--------------------------------------------
        public void NamesUpdate()
        {
            dataGridViewStudent.Columns["id"].Visible = false;
            medCheckDataGridView.Columns["id"].Visible = false;
            medCheckDataGridView.Columns["idStudent"].Visible = false;
            medCheckDataGridView.Columns["data"].HeaderText = "дата обследования";
            // ФИЛЬТРЫ ----------
            if (MyFilters.Rows[0].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil01"].Visible = true;
                medCheckDataGridView.Columns["fil01"].HeaderText = MyFilters.Rows[0].ItemArray[2].ToString();
                lab8.Visible = true;
                box8.Visible = true;
                lab8.Text = MyFilters.Rows[0].ItemArray[2].ToString();
                box8.DisplayMember = MyFilterValues1.Columns[0].ColumnName;
                box8.ValueMember = MyFilterValues1.Columns[0].ColumnName;
                box8.DataSource = MyFilterValues1;
            }
            else
            {
                medCheckDataGridView.Columns["fil01"].Visible = false;
                lab8.Visible = false;
                box8.Visible = false;
            }
            // -------
            if (MyFilters.Rows[1].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil02"].Visible = true;
                medCheckDataGridView.Columns["fil02"].HeaderText = MyFilters.Rows[1].ItemArray[2].ToString();
                lab9.Visible = true;
                box9.Visible = true;
                lab9.Text = MyFilters.Rows[1].ItemArray[2].ToString();
                box9.DisplayMember = MyFilterValues2.Columns[0].ColumnName;
                box9.ValueMember = MyFilterValues2.Columns[0].ColumnName;
                box9.DataSource = MyFilterValues2;
            }
            else
            {
                medCheckDataGridView.Columns["fil02"].Visible = false;
                lab9.Visible = false;
                box9.Visible = false;
            }
            // -------
            if (MyFilters.Rows[2].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil03"].Visible = true;
                medCheckDataGridView.Columns["fil03"].HeaderText = MyFilters.Rows[2].ItemArray[2].ToString();
                lab10.Visible = true;
                box10.Visible = true;
                lab10.Text = MyFilters.Rows[2].ItemArray[2].ToString();
                box10.DisplayMember = MyFilterValues3.Columns[0].ColumnName;
                box10.ValueMember = MyFilterValues3.Columns[0].ColumnName;
                box10.DataSource = MyFilterValues3;
            }
            else
            {
                medCheckDataGridView.Columns["fil03"].Visible = false;
                lab10.Visible = false;
                box10.Visible = false;
            }
            // -------
            if (MyFilters.Rows[3].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil04"].Visible = true;
                medCheckDataGridView.Columns["fil04"].HeaderText = MyFilters.Rows[3].ItemArray[2].ToString();
                lab11.Visible = true;
                box11.Visible = true;
                lab11.Text = MyFilters.Rows[3].ItemArray[2].ToString();
                box11.DisplayMember = MyFilterValues4.Columns[0].ColumnName;
                box11.ValueMember = MyFilterValues4.Columns[0].ColumnName;
                box11.DataSource = MyFilterValues4;
            }
            else
            {
                medCheckDataGridView.Columns["fil04"].Visible = false;
                lab11.Visible = false;
                box11.Visible = false;
            }
            // -------
            if (MyFilters.Rows[4].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil05"].Visible = true;
                medCheckDataGridView.Columns["fil05"].HeaderText = MyFilters.Rows[4].ItemArray[2].ToString();
                lab12.Visible = true;
                box12.Visible = true;
                lab12.Text = MyFilters.Rows[4].ItemArray[2].ToString();
                box12.DisplayMember = MyFilterValues5.Columns[0].ColumnName;
                box12.ValueMember = MyFilterValues5.Columns[0].ColumnName;
                box12.DataSource = MyFilterValues5;
            }
            else
            {
                medCheckDataGridView.Columns["fil05"].Visible = false;
                lab12.Visible = false;
                box12.Visible = false;
            }
            // -------
            if (MyFilters.Rows[5].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil06"].Visible = true;
                medCheckDataGridView.Columns["fil06"].HeaderText = MyFilters.Rows[5].ItemArray[2].ToString();
                lab13.Visible = true;
                box13.Visible = true;
                lab13.Text = MyFilters.Rows[5].ItemArray[2].ToString();
                box13.DisplayMember = MyFilterValues6.Columns[0].ColumnName;
                box13.ValueMember = MyFilterValues6.Columns[0].ColumnName;
                box13.DataSource = MyFilterValues6;
            }
            else
            {
                medCheckDataGridView.Columns["fil06"].Visible = false;
                lab13.Visible = false;
                box13.Visible = false;
            }
            // -------
            if (MyFilters.Rows[6].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil07"].Visible = true;
                medCheckDataGridView.Columns["fil07"].HeaderText = MyFilters.Rows[6].ItemArray[2].ToString();
                lab14.Visible = true;
                box14.Visible = true;
                lab14.Text = MyFilters.Rows[6].ItemArray[2].ToString();
                box14.DisplayMember = MyFilterValues7.Columns[0].ColumnName;
                box14.ValueMember = MyFilterValues7.Columns[0].ColumnName;
                box14.DataSource = MyFilterValues7;
            }
            else
            {
                medCheckDataGridView.Columns["fil07"].Visible = false;
                lab14.Visible = false;
                box14.Visible = false;
            }
            // -------
            if (MyFilters.Rows[7].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil08"].Visible = true;
                medCheckDataGridView.Columns["fil08"].HeaderText = MyFilters.Rows[7].ItemArray[2].ToString();
                lab15.Visible = true;
                box15.Visible = true;
                lab15.Text = MyFilters.Rows[7].ItemArray[2].ToString();
                box15.DisplayMember = MyFilterValues8.Columns[0].ColumnName;
                box15.ValueMember = MyFilterValues8.Columns[0].ColumnName;
                box15.DataSource = MyFilterValues8;
            }
            else
            {
                medCheckDataGridView.Columns["fil08"].Visible = false;
                lab15.Visible = false;
                box15.Visible = false;
            }
            // -------
            if (MyFilters.Rows[8].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil09"].Visible = true;
                medCheckDataGridView.Columns["fil09"].HeaderText = MyFilters.Rows[8].ItemArray[2].ToString();
                lab16.Visible = true;
                box16.Visible = true;
                lab16.Text = MyFilters.Rows[8].ItemArray[2].ToString();
                box16.DisplayMember = MyFilterValues9.Columns[0].ColumnName;
                box16.ValueMember = MyFilterValues9.Columns[0].ColumnName;
                box16.DataSource = MyFilterValues9;
            }
            else
            {
                medCheckDataGridView.Columns["fil09"].Visible = false;
                lab16.Visible = false;
                box16.Visible = false;
            }
            // -------
            if (MyFilters.Rows[9].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["fil10"].Visible = true;
                medCheckDataGridView.Columns["fil10"].HeaderText = MyFilters.Rows[9].ItemArray[2].ToString();
                lab17.Visible = true;
                box17.Visible = true;
                lab17.Text = MyFilters.Rows[9].ItemArray[2].ToString();
                box17.DisplayMember = MyFilterValues10.Columns[0].ColumnName;
                box17.ValueMember = MyFilterValues10.Columns[0].ColumnName;
                box17.DataSource = MyFilterValues10;
            }
            else
            {
                medCheckDataGridView.Columns["fil10"].Visible = false;
                lab17.Visible = false;
                box17.Visible = false;
            }
            // ИЗМЕРЕНИЯ ----------
            if (MyIzm.Rows[0].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm01"].Visible = true;
                medCheckDataGridView.Columns["izm01"].HeaderText = MyIzm.Rows[0].ItemArray[2].ToString();
                lab18.Visible = true;
                box18.Visible = true;
                lab18.Text = MyIzm.Rows[0].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm01"].Visible = false;
                lab18.Visible = false;
                box18.Visible = false;
            }
            // -------
            if (MyIzm.Rows[1].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm02"].Visible = true;
                medCheckDataGridView.Columns["izm02"].HeaderText = MyIzm.Rows[1].ItemArray[2].ToString();
                lab19.Visible = true;
                box19.Visible = true;
                lab19.Text = MyIzm.Rows[1].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm02"].Visible = false;
                lab19.Visible = false;
                box19.Visible = false;
            }
            // -------
            if (MyIzm.Rows[2].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm03"].Visible = true;
                medCheckDataGridView.Columns["izm03"].HeaderText = MyIzm.Rows[2].ItemArray[2].ToString();
                lab20.Visible = true;
                box20.Visible = true;
                lab20.Text = MyIzm.Rows[2].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm03"].Visible = false;
                lab20.Visible = false;
                box20.Visible = false;
            }
            // -------
            if (MyIzm.Rows[3].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm04"].Visible = true;
                medCheckDataGridView.Columns["izm04"].HeaderText = MyIzm.Rows[3].ItemArray[2].ToString();
                lab21.Visible = true;
                box21.Visible = true;
                lab21.Text = MyIzm.Rows[3].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm04"].Visible = false;
                lab21.Visible = false;
                box21.Visible = false;
            }
            // -------
            if (MyIzm.Rows[4].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm05"].Visible = true;
                medCheckDataGridView.Columns["izm05"].HeaderText = MyIzm.Rows[4].ItemArray[2].ToString();
                lab22.Visible = true;
                box22.Visible = true;
                lab22.Text = MyIzm.Rows[4].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm05"].Visible = false;
                lab22.Visible = false;
                box22.Visible = false;
            }
            // -------
            if (MyIzm.Rows[5].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm06"].Visible = true;
                medCheckDataGridView.Columns["izm06"].HeaderText = MyIzm.Rows[5].ItemArray[2].ToString();
                lab23.Visible = true;
                box23.Visible = true;
                lab23.Text = MyIzm.Rows[5].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm06"].Visible = false;
                lab23.Visible = false;
                box23.Visible = false;
            }
            // -------
            if (MyIzm.Rows[6].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm07"].Visible = true;
                medCheckDataGridView.Columns["izm07"].HeaderText = MyIzm.Rows[6].ItemArray[2].ToString();
                lab24.Visible = true;
                box24.Visible = true;
                lab24.Text = MyIzm.Rows[6].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm07"].Visible = false;
                lab24.Visible = false;
                box24.Visible = false;
            }
            // -------
            if (MyIzm.Rows[7].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm08"].Visible = true;
                medCheckDataGridView.Columns["izm08"].HeaderText = MyIzm.Rows[7].ItemArray[2].ToString();
                lab25.Visible = true;
                box25.Visible = true;
                lab25.Text = MyIzm.Rows[7].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm08"].Visible = false;
                lab25.Visible = false;
                box25.Visible = false;
            }
            // -------
            if (MyIzm.Rows[8].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm09"].Visible = true;
                medCheckDataGridView.Columns["izm09"].HeaderText = MyIzm.Rows[8].ItemArray[2].ToString();
                lab26.Visible = true;
                box26.Visible = true;
                lab26.Text = MyIzm.Rows[8].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm09"].Visible = false;
                lab26.Visible = false;
                box26.Visible = false;
            }
            // -------
            if (MyIzm.Rows[9].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm10"].Visible = true;
                medCheckDataGridView.Columns["izm10"].HeaderText = MyIzm.Rows[9].ItemArray[2].ToString();
                lab27.Visible = true;
                box27.Visible = true;
                lab27.Text = MyIzm.Rows[9].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm10"].Visible = false;
                lab27.Visible = false;
                box27.Visible = false;
            }
            // -------
            if (MyIzm.Rows[10].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm11"].Visible = true;
                medCheckDataGridView.Columns["izm11"].HeaderText = MyIzm.Rows[10].ItemArray[2].ToString();
                lab28.Visible = true;
                box28.Visible = true;
                lab28.Text = MyIzm.Rows[10].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm11"].Visible = false;
                lab28.Visible = false;
                box28.Visible = false;
            }
            // -------
            if (MyIzm.Rows[11].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm12"].Visible = true;
                medCheckDataGridView.Columns["izm12"].HeaderText = MyIzm.Rows[11].ItemArray[2].ToString();
                lab29.Visible = true;
                box29.Visible = true;
                lab29.Text = MyIzm.Rows[11].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm12"].Visible = false;
                lab29.Visible = false;
                box29.Visible = false;
            }
            // -------
            if (MyIzm.Rows[12].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm13"].Visible = true;
                medCheckDataGridView.Columns["izm13"].HeaderText = MyIzm.Rows[12].ItemArray[2].ToString();
                lab30.Visible = true;
                box30.Visible = true;
                lab30.Text = MyIzm.Rows[12].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm13"].Visible = false;
                lab30.Visible = false;
                box30.Visible = false;
            }
            // -------
            if (MyIzm.Rows[13].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm14"].Visible = true;
                medCheckDataGridView.Columns["izm14"].HeaderText = MyIzm.Rows[13].ItemArray[2].ToString();
                lab31.Visible = true;
                box31.Visible = true;
                lab31.Text = MyIzm.Rows[13].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm14"].Visible = false;
                lab31.Visible = false;
                box31.Visible = false;
            }
            // -------
            if (MyIzm.Rows[14].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm15"].Visible = true;
                medCheckDataGridView.Columns["izm15"].HeaderText = MyIzm.Rows[14].ItemArray[2].ToString();
                lab32.Visible = true;
                box32.Visible = true;
                lab32.Text = MyIzm.Rows[14].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm15"].Visible = false;
                lab32.Visible = false;
                box32.Visible = false;
            }
            // -------
            if (MyIzm.Rows[15].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm16"].Visible = true;
                medCheckDataGridView.Columns["izm16"].HeaderText = MyIzm.Rows[15].ItemArray[2].ToString();
                lab33.Visible = true;
                box33.Visible = true;
                lab33.Text = MyIzm.Rows[15].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm16"].Visible = false;
                lab33.Visible = false;
                box33.Visible = false;
            }
            // -------
            if (MyIzm.Rows[16].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm17"].Visible = true;
                medCheckDataGridView.Columns["izm17"].HeaderText = MyIzm.Rows[16].ItemArray[2].ToString();
                lab34.Visible = true;
                box34.Visible = true;
                lab34.Text = MyIzm.Rows[16].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm17"].Visible = false;
                lab34.Visible = false;
                box34.Visible = false;
            }
            // -------
            if (MyIzm.Rows[17].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm18"].Visible = true;
                medCheckDataGridView.Columns["izm18"].HeaderText = MyIzm.Rows[17].ItemArray[2].ToString();
                lab35.Visible = true;
                box35.Visible = true;
                lab35.Text = MyIzm.Rows[17].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm18"].Visible = false;
                lab35.Visible = false;
                box35.Visible = false;
            }
            // -------
            if (MyIzm.Rows[18].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm19"].Visible = true;
                medCheckDataGridView.Columns["izm19"].HeaderText = MyIzm.Rows[18].ItemArray[2].ToString();
                lab36.Visible = true;
                box36.Visible = true;
                lab36.Text = MyIzm.Rows[18].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm19"].Visible = false;
                lab36.Visible = false;
                box36.Visible = false;
            }
            // -------
            if (MyIzm.Rows[19].ItemArray[1].Equals(true))
            {
                medCheckDataGridView.Columns["izm20"].Visible = true;
                medCheckDataGridView.Columns["izm20"].HeaderText = MyIzm.Rows[19].ItemArray[2].ToString();
                lab37.Visible = true;
                box37.Visible = true;
                lab37.Text = MyIzm.Rows[19].ItemArray[2].ToString();
            }
            else
            {
                medCheckDataGridView.Columns["izm20"].Visible = false;
                lab37.Visible = false;
                box37.Visible = false;
            }

        } // of NamesUpdate

        //===============================================================================================================
        // Обработчик события перемещения точки ----------------------------------------------------------------
        /// <param name="sender">Компонент ZedGraph</param>
        /// <param name="pane">Панель с графиком</param>
        /// <param name="curve">Кривая, точку которой переместили</param>
        /// <param name="iPt">Номер точки</param>
        /// <returns>Метод должен возвращать строку</returns>
        string zedGraphControl1_PointEditEvent(ZedGraphControl sender, GraphPane pane, CurveItem curve, int iPt)
        {
            // округлим значения
            curve[iPt].X = Math.Round(curve[iPt].X, 0);
            curve[iPt].Y = Math.Round(curve[iPt].Y, 0);
            // получим строки
            string x_csv = MyGraphs.Rows[dataGridView3.CurrentCell.RowIndex]["x"].ToString();
            string y_csv = MyGraphs.Rows[dataGridView3.CurrentCell.RowIndex]["y"].ToString();
            // преобразуем строки в массив
            string[] xx = x_csv.Split(',');
            string[] yy = y_csv.Split(',');
            // заменим нужную точку в массиве
            xx[iPt] = curve[iPt].X.ToString();
            yy[iPt] = curve[iPt].Y.ToString();
            // снова преобразуем массивы в строки
            x_csv = xx[0].ToString();
            y_csv = yy[0].ToString();
            for (int i = 1; i < xx.Length; i++) // со второго значения
            {
                x_csv = x_csv + "," + xx[i];
                y_csv = y_csv + "," + yy[i];
            }
            //сохраним строки в таблицу
            MyGraphs.Rows[dataGridView3.CurrentCell.RowIndex]["x"] = x_csv;
            MyGraphs.Rows[dataGridView3.CurrentCell.RowIndex]["y"] = y_csv;
            FlagMark = true;
            // статус
            string title = string.Format("Точка: {0}. Новые координаты: ({1}; {2})", iPt, curve[iPt].X, curve[iPt].Y);
            dot_label.Text = title;
            zedGraphControl1.Invalidate();
            TotalMark_label.Text = "Максимальные общие оценки: Муж=" + TotalBall("М") + " Жен=" + TotalBall("Ж");
            return title;
        }

        //===============================================================================================================
        // Обновить график----------------------------------------------------------------
        private void DrawGraph(String name, string[] xx, string[] yy)
        // name - название оценки - для вывода в название оси У
        // массивы хх и уу - точки
        {
            // Получим панель для рисования
            GraphPane pane = zedGraphControl1.GraphPane;
            // Оформление
            pane.XAxis.Title.Text = "Оценка (баллы)"; // Изменим тест надписи по оси X
            //pane.XAxis.Title.FontSpec.IsBold = false;             // Изменим параметры шрифта для оси X
            //pane.XAxis.Title.FontSpec.FontColor = Color.Blue;
            pane.YAxis.Title.Text = name;             // Изменим текст по оси Y
            pane.Title.Text = "График";             // Изменим текст заголовка графика
            //pane.Title.FontSpec.Fill.Brush = new SolidBrush(Color.Red);             // В параметрах шрифта сделаем заливку красным цветом
            //pane.Title.FontSpec.Fill.IsVisible = true;
            // Сделаем шрифт не полужирным
            //pane.Title.FontSpec.IsBold = false;
            pane.CurveList.Clear(); // Очистим список кривых на тот случай, если до этого сигналы уже были нарисованы
            pane.Legend.IsVisible = false; // легенду спрячем
            // Создадим список точек
            PointPairList list = new PointPairList();
            // Заполняем список точек
            for (int dot = 0; dot < xx.Length; dot++)
            {
                // добавим в список точку
                double dxx = Double.Parse(xx[dot]);
                double dyy = Double.Parse(yy[dot]);
                list.Add(dxx, dyy);
            }
            LineItem myCurve = pane.AddCurve("Sinc", list, Color.Blue, SymbolType.Square); // Создадим кривую с названием "Sinc", 
            // У кривой линия будет невидимой
            //myCurve.Line.IsVisible = false;
            myCurve.Symbol.Fill.Color = Color.Blue;    // Цвет заполнения отметок (ромбиков) - голубой
            myCurve.Symbol.Fill.Type = FillType.Solid; // Тип заполнения - сплошная заливка
            myCurve.Symbol.Size = 7; // Размер ромбиков
            /*
            // Устанавливаем интересующий нас интервал по оси X
            pane.XAxis.Scale.Min = xmin;
            pane.XAxis.Scale.Max = xmax;
            // Устанавливаем интересующий нас интервал по оси Y
            pane.YAxis.Scale.Min = ymin;
            pane.YAxis.Scale.Max = ymax;
            */
            zedGraphControl1.AxisChange();// Вызываем метод AxisChange (), чтобы обновить данные об осях. В противном случае на рисунке будет показана только часть графика, которая умещается в интервалы по осям, установленные по умолчанию
            // Обновляем график
            zedGraphControl1.Invalidate();
        }

        //===============================================================================================================
        // ОТКРЫТЬ КОНСТРУКТОР ДЛЯ СОЗДАНИЯ НОВОЙ ОЦЕНКИ ----------------------------------------------------------------
        private void button8_Click(object sender, EventArgs e)
        {
            // заряжаем в таблицы названия измерений ( для отображения в комбобоксах формы2)
            System.Data.DataTable tempIzm1 = new System.Data.DataTable(); //врем.таблица
            tempIzm1.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm1.Columns.Add(new DataColumn("name", typeof(String)));
            System.Data.DataTable tempIzm2 = new System.Data.DataTable(); //нужно 2, хоть они и одинаковые, иначе синхронизируются
            tempIzm2.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm2.Columns.Add(new DataColumn("name", typeof(String)));
            for (int i = 0; i < MyIzm.Rows.Count; i++)
            {
                if (MyIzm.Rows[i]["вкл"].Equals(true)) // if == true (галочка стоит)
                {
                    tempIzm1.Rows.Add(i + 1, MyIzm.Rows[i]["название"]);
                    tempIzm2.Rows.Add(i + 1, MyIzm.Rows[i]["название"]);
                }
            }
            Form2.Izm1table = tempIzm1;
            Form2.Izm2table = tempIzm2;
            // открываем подряд 3 формы
            Form ConstrForm2 = new Form2();
            Form ConstrForm3 = new Form3();
            Form ConstrForm4 = new Form4();
            ConstrForm2.ShowDialog();
            if (form2Cancel == false) // если не нажималась отмена, продолжаем конструктор
                ConstrForm3.ShowDialog();
            if (form3Cancel == false) // если не нажималась отмена, продолжаем конструктор
                ConstrForm4.ShowDialog();
            // для постройки графика всё собрали:
            // string formName, formEd, formSposob(копируется,вычисляется), formIzm1(назв.изм), formZnak(-,+,*,/), formIzm2(назв.изм);
            // double formXmax, formDot, formYmin, FormYmax, formGraf(1,2,3,4);
            if (form4Cancel == false) // если не нажималась отмена, продолжаем конструктор
            {
                int vremznak = 0;
                int vremMult = 1;
                if (formSposob == "копируется")
                    vremznak = 0;
                else
                    switch (formZnak)
                    {
                        case "+":// вычисляется +
                            vremznak = 1;
                            break;
                        case "-":// вычисляется -
                            vremznak = 2;
                            break;
                        case "*":// вычисляется *
                            vremznak = 3;
                            break;
                        case "/":// вычисляется /
                            vremznak = 4;
                            break;
                        default:
                            vremznak = 0; // неверное значение
                            break;
                    }
                // переводим название изм в цифру
                int vremIzm1 = 0, vremIzm2 = 0;
                for (int i = 0; i < MyIzm.Rows.Count; i++)
                {
                    if (MyIzm.Rows[i].ItemArray[2].ToString() == formIzm1) vremIzm1 = i;
                    if (MyIzm.Rows[i].ItemArray[2].ToString() == formIzm2) vremIzm2 = i;
                }
                // мультиплаер
                vremMult = int.Parse(formMult);
                MySpisok.Rows.Add(true, formName, formEd, vremIzm1, vremIzm2, vremznak, vremMult, formSex);
                // Создадим список точек
                // используем глобальные dotkol xmin xstep xmax ystep ymin ymax
                dotkol = formDot; // точек будет столько +1 крайняя
                xmax = formXmax;
                ymin = formYmin;
                ymax = FormYmax;
                xstep = Math.Round((xmax - xmin) / dotkol, 0); // рассчитаем шаг по оси х
                ystep = Math.Round((ymax - ymin) / dotkol, 0); // рассчитаем шаг по оси у
                String xx, yy;
                double dot, dotkolpopolam;
                dotkolpopolam = Math.Round(dotkol / 2, 0); // для сложных графиков
                // 4 типа графика
                switch (formGraf)
                {
                    case 1: // чем больше - тем лучше
                        xx = xmin.ToString();
                        yy = ymin.ToString();
                        for (dot = 1; dot <= dotkol; dot++)
                        {
                            xx = xx + "," + (xmin + dot * xstep);
                            yy = yy + "," + (ymin + dot * ystep);
                        }
                        break;
                    case 2: // чем меньше - тем лучше
                        xx = xmax.ToString();
                        yy = ymin.ToString();
                        for (dot = dotkol - 1; dot >= 0; dot--)
                            xx = xx + "," + (xmin + dot * xstep);
                        for (dot = 1; dot <= dotkol; dot++)
                            yy = yy + "," + (ymin + dot * ystep);
                        break;
                    case 3: // лучше середина
                        xx = xmin.ToString();
                        yy = ymin.ToString();
                        for (dot = 1; dot <= dotkol; dot++)
                            yy = yy + "," + (ymin + dot * ystep);
                        for (dot = 1; dot < dotkolpopolam; dot++)
                            xx = xx + "," + (xmin + dot * xstep * 2); // полграфика
                        for (dot = dotkolpopolam; dot >= 0; dot--)
                            xx = xx + "," + (xmin + dot * xstep * 2); // обратно полграфика
                        break;
                    case 4: // середина хуже
                        xx = xmax.ToString();
                        yy = ymin.ToString();
                        for (dot = 1; dot <= dotkol; dot++)
                            yy = yy + "," + (ymin + dot * ystep);
                        for (dot = dotkol - 1; dot > dotkolpopolam; dot--)
                            xx = xx + "," + (xmin + dot * xstep * 2 - xmax); // полграфика
                        for (dot = dotkolpopolam; dot <= dotkol; dot++)
                            xx = xx + "," + (xmin + dot * xstep * 2 - xmax); // ещё полграфика
                        break;
                    default:
                        xx = "0";
                        yy = "0";
                        break;
                }
                MyGraphs.Rows.Add(xx, yy);
                FlagMark = true;
                TotalMark_label.Text = "Максимальные общие оценки: Муж=" + TotalBall("М") + " Жен=" + TotalBall("Ж");
                dot_label.Text = "";
            }
        }

        //===============================================================================================================
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"]) // вкладка "База данных"
            {
                FlagStud = false; // Отмечаем, что изменений пока нет
                FlagCheck = false;
                button11.Visible = false;
                button12.Visible = false;
                dataGridViewStudent.Enabled = true;
                medCheckDataGridView.Enabled = true;
                bdFileConnect(); // заново откроем базу, на вкладке ВВОД могли поменять данные
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"]) // вкладка "Ввод"
            {
                // прячем лишние элементы
                label1.Visible = false;
                box7.Visible = false;
                label8.Visible = false;
                groupBox3.Visible = false;
                groupBox4.Visible = false;
                groupBox5.Visible = false;
                box1.Focus(); // курсор на 1ый текстбокс
                ClearVvodForm();
                button9.Visible = false;
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"]) // вкладка "Оценки"
            {
                FlagMark = false; // Отмечаем, что изменений пока нет
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"]) // вкладка "Настройки"
            {
                FlagIzm = false; // Отмечаем, что изменений пока нет
            }
        }

        //===============================================================================================================
        // Очистка всех элементов формы ввода
        public void ClearVvodForm()
        {
            box0.Text = "";
            box1.Text = "";
            box2.Text = "";
            box3.Text = "";
            box4.Text = "";
            box5.Text = "";
            box6.Value = DateTime.Parse("01.01.2000");
            box7.Value = DateTime.Today;
            box8.Text = "";
            box9.Text = "";
            box10.Text = "";
            box11.Text = "";
            box12.Text = "";
            box13.Text = "";
            box14.Text = "";
            box15.Text = "";
            box16.Text = "";
            box17.Text = "";
            box18.Text = "";
            box19.Text = "";
            box20.Text = "";
            box21.Text = "";
            box22.Text = "";
            box23.Text = "";
            box24.Text = "";
            box25.Text = "";
            box26.Text = "";
            box27.Text = "";
            box28.Text = "";
            box29.Text = "";
            box30.Text = "";
            box31.Text = "";
            box32.Text = "";
            box33.Text = "";
            box34.Text = "";
            box35.Text = "";
            box36.Text = "";
            box37.Text = "";
        }

        //===============================================================================================================
        // процедура режет все нажатые клавиши, кроме цифр и Backspace (8) и минус (45)
        private void Digit_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 45) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        //===============================================================================================================
        // процедура режет все клавиши, кроме м(236) М(204) ж(230) Ж(198) и Backspace
        private void Sex_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (number != 1052 && number != 1084 && number != 1078 && number != 1046 && number != 8)
            {
                e.Handled = true;
            }
            if (box5.TextLength > 0 && number != 8)
            {
                e.Handled = true;
            }
        }

        //===============================================================================================================
        // функция возвращает сумму максимальных баллов по всем оценкам указанного пола
        // сумма максимальных иксов по всем графикам
        private Double TotalBall(string osex)
        {
            Double Total_Ball = 0;
            for (int mi = 0; mi < MyGraphs.Rows.Count; mi++)
                if (MySpisok.Rows[mi]["вкл"].Equals(true)) // оценка включена
                    if (CalculateSex(MySpisok.Rows[mi]["пол"].ToString(), osex)) // оценка нужного пола
                        Total_Ball = Total_Ball + MaxBall(mi);
            return Total_Ball;
        }

        //===============================================================================================================
        // функция возвращает сумму максимального балла по текущей оценке
        // максимальный икс в текущем графике
        private Double MaxBall(int Ngraf)
        {
            Double Max_Ball = 0;
            // получим строку
            string x_csv = MyGraphs.Rows[Ngraf]["x"].ToString();
            // преобразуем строки в массив
            string[] xx = x_csv.Split(',');
            // пробежим по иксам и ищем мамсимум
            for (int i = 0; i < xx.Length; i++)
                if (Double.Parse(xx[i]) > Max_Ball) Max_Ball = Double.Parse(xx[i]);
            return Max_Ball;
        }

        //===============================================================================================================
        // при удалении оценки
        private void dataGridView3_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            int SelectedFilter = -1;
            if (dataGridView3.CurrentCell != null)
            {
                SelectedFilter = dataGridView3.CurrentCell.RowIndex;
                dataGridView3.CurrentCell = null;
                // удаляем график
                MyGraphs.Rows[SelectedFilter].Delete();
                groupBox2.Visible = false; // закроем правую панель
                TotalMark_label.Text = "Максимальные общие оценки: Муж=" + TotalBall("М") + " Жен=" + TotalBall("Ж");
                FlagMark = true;
            }
        }

        //===============================================================================================================

        private void dataGridView3_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            dataGridView3.CurrentCell = null;
            FlagMark = true;
        }

        //===============================================================================================================

        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            checkPrint = 0;
        }

        //===============================================================================================================

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            // Print the content of RichTextBox. Store the last character printed.
            checkPrint = richTextBoxPrintCtrl1.Print(checkPrint, richTextBoxPrintCtrl1.TextLength, e);
            // Check for more pages
            if (checkPrint < richTextBoxPrintCtrl1.TextLength)
                e.HasMorePages = true;
            else
                e.HasMorePages = false;
        }

        //===============================================================================================================
        // Подготовка паспорта ФР
        public void PrintPassport(string p_passport)
        {
            string print_filters1, print_filters2, print_sex,temp_data;
            Bitmap graf;
            string print_id = "0";
            DataSet printDataSet = new DataSet();
            string commandText = "SELECT * FROM Student WHERE Student.passport = @passport;";
            SQLiteCommand cmd = new SQLiteCommand(commandText, dbConn);
            SQLiteParameter Param = new SQLiteParameter("@passport", DbType.StringFixedLength);
            Param.Value = p_passport;
            cmd.Parameters.Add(Param);
            SQLiteDataAdapter printDataAdapter = new SQLiteDataAdapter(cmd);
            printDataAdapter.Fill(printDataSet);
            int i = 0;
            if (printDataSet.Tables[0].Rows.Count > 0) // обследуемых в выборке должно быть больше 1, вообщето 1, берем первого
            {
                // Загрузить шаблон
                Word.Document doc = null;
                Word.Application app = new Word.Application(); // Создаём объект приложения
                try
                {
                    doc = app.Documents.Open(dbPath + "\\" + dbName + passportDOCX);
                    // doc = app.Documents.Open(dbPath + "\\" + dbName + passportDOCX, System.Reflection.Missing.Value, true); // 3ИЙ ПАРАМЕТР - только для чтения
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка. Шаблон для Паспорта не загружен!" + Environment.NewLine + "(" + ex.Message + ")");
                }
                doc.Activate();
                print_id = printDataSet.Tables[0].Rows[i]["Id"].ToString(); // сохраним его Id (int->string)
                Word.Bookmarks wBookmarks = doc.Bookmarks; // содержит все закладки
                // документ
                for (int nb = 1; nb <= wBookmarks.Count; nb++)
                    if (wBookmarks[nb].Name == "doc") wBookmarks[nb].Range.Text = printDataSet.Tables[0].Rows[i]["passport"].ToString();
                // фамилия
                for (int nb = 1; nb <= wBookmarks.Count; nb++)
                    if (wBookmarks[nb].Name == "f") wBookmarks[nb].Range.Text = printDataSet.Tables[0].Rows[i]["f"].ToString();
                // имя
                for (int nb = 1; nb <= wBookmarks.Count; nb++)
                    if (wBookmarks[nb].Name == "i") wBookmarks[nb].Range.Text = printDataSet.Tables[0].Rows[i]["i"].ToString();
                // отчество
                for (int nb = 1; nb <= wBookmarks.Count; nb++)
                    if (wBookmarks[nb].Name == "o") wBookmarks[nb].Range.Text = printDataSet.Tables[0].Rows[i]["o"].ToString();
                // пол
                print_sex = printDataSet.Tables[0].Rows[i]["sex"].ToString(); // сохраним, понадобится
                for (int nb = 1; nb <= wBookmarks.Count; nb++)
                    if (wBookmarks[nb].Name == "sex") wBookmarks[nb].Range.Text = printDataSet.Tables[0].Rows[i]["sex"].ToString();
                // дата рождения
                temp_data = printDataSet.Tables[0].Rows[i]["born"].ToString();
                if (temp_data.Length > 10) temp_data = temp_data.Substring(0, 10);
                for (int nb = 1; nb <= wBookmarks.Count; nb++)
                    if (wBookmarks[nb].Name == "born") wBookmarks[nb].Range.Text = temp_data;
                // -----------------МЕДОСМОТРЫ------------------------
                DataSet printDataSet2 = new DataSet();
                commandText = "SELECT * FROM MedCheck WHERE MedCheck.idStudent = @idStudent;";
                cmd = new SQLiteCommand(commandText, dbConn);
                Param = new SQLiteParameter("@idStudent", DbType.Int32);
                Param.Value = Convert.ToInt32(print_id); // выбранный Обследуемый
                cmd.Parameters.Add(Param);
                SQLiteDataAdapter printDataAdapter2 = new SQLiteDataAdapter(cmd);
                printDataAdapter2.Fill(printDataSet2);
                // фильтры берутся из последнего медосмотра
                int kol_checks = printDataSet2.Tables[0].Rows.Count;
                // выводим в 2 строки
                print_filters1 = "Медосмотры не найдены.";
                print_filters2 = "";
                if (kol_checks > 0) // если есть хоть 1 медосмотр
                {
                    // 1ая строка- названия и значения фильтров с fil1 по fil5
                    print_filters1 = "";
                    if (MyFilters.Rows[0].ItemArray[1].Equals(true))
                    {
                        print_filters1 += MyFilters.Rows[0].ItemArray[2].ToString(); // название
                        print_filters1 += ": ";
                        print_filters1 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil01"].ToString(); //значение
                        print_filters1 += "    ";
                    }
                    if (MyFilters.Rows[1].ItemArray[1].Equals(true))
                    {
                        print_filters1 += MyFilters.Rows[1].ItemArray[2].ToString(); // название
                        print_filters1 += ": ";
                        print_filters1 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil02"].ToString(); //значение
                        print_filters1 += "    ";
                    }
                    if (MyFilters.Rows[2].ItemArray[1].Equals(true))
                    {
                        print_filters1 += MyFilters.Rows[2].ItemArray[2].ToString(); // название
                        print_filters1 += ": ";
                        print_filters1 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil03"].ToString(); //значение
                        print_filters1 += "    ";
                    }
                    if (MyFilters.Rows[3].ItemArray[1].Equals(true))
                    {
                        print_filters1 += MyFilters.Rows[3].ItemArray[2].ToString(); // название
                        print_filters1 += ": ";
                        print_filters1 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil04"].ToString(); //значение
                        print_filters1 += "    ";
                    }
                    if (MyFilters.Rows[4].ItemArray[1].Equals(true))
                    {
                        print_filters1 += MyFilters.Rows[4].ItemArray[2].ToString(); // название
                        print_filters1 += ": ";
                        print_filters1 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil05"].ToString(); //значение
                    }
                    // 2ая строка- названия и значения фильтров с fil6 по fil10
                    if (MyFilters.Rows[5].ItemArray[1].Equals(true))
                    {
                        print_filters2 += MyFilters.Rows[5].ItemArray[2].ToString(); // название
                        print_filters2 += ": ";
                        print_filters2 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil06"].ToString(); //значение
                        print_filters2 += "    ";
                    }
                    if (MyFilters.Rows[6].ItemArray[1].Equals(true))
                    {
                        print_filters2 += MyFilters.Rows[6].ItemArray[2].ToString(); // название
                        print_filters2 += ": ";
                        print_filters2 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil07"].ToString(); //значение
                        print_filters2 += "    ";
                    }
                    if (MyFilters.Rows[7].ItemArray[1].Equals(true))
                    {
                        print_filters2 += MyFilters.Rows[7].ItemArray[2].ToString(); // название
                        print_filters2 += ": ";
                        print_filters2 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil08"].ToString(); //значение
                        print_filters2 += "    ";
                    }
                    if (MyFilters.Rows[8].ItemArray[1].Equals(true))
                    {
                        print_filters2 += MyFilters.Rows[8].ItemArray[2].ToString(); // название
                        print_filters2 += ": ";
                        print_filters2 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil09"].ToString(); //значение
                        print_filters2 += "    ";
                    }
                    if (MyFilters.Rows[9].ItemArray[1].Equals(true))
                    {
                        print_filters2 += MyFilters.Rows[9].ItemArray[2].ToString(); // название
                        print_filters2 += ": ";
                        print_filters2 += printDataSet2.Tables[0].Rows[kol_checks - 1]["fil10"].ToString(); //значение
                    }
                    // запишем в docx
                    for (int nb = 1; nb <= wBookmarks.Count; nb++)
                        if (wBookmarks[nb].Name == "filt") wBookmarks[nb].Range.Text = print_filters1 + " " + print_filters2; // фильтры
                    // ------------ТАБЛИЦА-------------
                    Word.Table tbl = app.ActiveDocument.Tables[1];
                    // шапка таблицы
                    double vremMark, vremBall;
                    double[] Itogo = new double[5] { 0, 0, 0, 0, 0 }; // Итоговые суммы
                    tbl.Rows[1].Cells[1].Range.Text = "Медосмотры:"; // 1ая таблицы строка уже есть, пишем в неё
                    for (int k = 0; k < kol_checks; k++)
                    {
                        temp_data = printDataSet2.Tables[0].Rows[k]["data"].ToString();
                        if (temp_data.Length > 10) temp_data = temp_data.Substring(0, 10);
                        tbl.Rows[1].Cells[k + 2].Range.Text = temp_data;
                    }
                    // тело таблицы - ИЗМЕРЕНИЯ
                    for (int mi = 0; mi < MyIzm.Rows.Count; mi++) // цикл по измерениям
                    {
                        if (MyIzm.Rows[mi]["вкл"].Equals(true)) // if == true (галочка стоит)
                        {
                            tbl.Rows.Add();
                            tbl.Rows[tbl.Rows.Count].Cells[1].Range.Text = MyIzm.Rows[mi]["название"].ToString();
                            for (int k = 0; k < kol_checks; k++)
                                tbl.Rows[tbl.Rows.Count].Cells[k + 2].Range.Text = printDataSet2.Tables[0].Rows[k][mi + 13].ToString();
                        } // оценка включена
                    } // по измерениям
                    // тело таблицы - ОЦЕНКИ
                    // Сначала просто посчитаем включенные оценки
                    int kol_izm = 0;
                    for (int mi = 0; mi < MySpisok.Rows.Count; mi++) // цикл по оценкам
                        if ((MySpisok.Rows[mi]["вкл"].Equals(true)) && (CalculateSex(MySpisok.Rows[mi]["пол"].ToString(), print_sex))) // если оценка включена и нужный нам пол
                            kol_izm++; // считаем для графика
                    // в этих 3ех массивах данные для графика (данные последнего ! техосмотра)
                    string[] graf1 = new string[kol_izm];   // массив названий для графика
                    double[] graf2 = new double[kol_izm]; // массив баллов для графика
                    double[] graf3 = new double[kol_izm]; // массив макс.баллов для графика
                    kol_izm = 0; // для запоминания индекса - mi не подходит, есть нулевые же...
                    for (int mi = 0; mi < MySpisok.Rows.Count; mi++) // цикл по оценкам
                    {
                        // если оценка включена и нужный нам пол
                        if ((MySpisok.Rows[mi]["вкл"].Equals(true)) && (CalculateSex(MySpisok.Rows[mi]["пол"].ToString(), print_sex)))
                        {
                            tbl.Rows.Add();
                            tbl.Rows[tbl.Rows.Count].Cells[1].Range.Text = MySpisok.Rows[mi]["название"].ToString() + " (" + MySpisok.Rows[mi]["ед.изм."].ToString() + ")";
                            for (int k = 0; k < kol_checks; k++)
                            {
                                vremMark = CalculateMark(mi, printDataSet2.Tables[0].Rows[k]);
                                vremBall = CalculateBall(mi, vremMark);
                                graf1[kol_izm] = MySpisok.Rows[mi]["название"].ToString(); // и сюда дублируем
                                graf2[kol_izm] = vremBall;
                                graf3[kol_izm] = MaxBall(mi);
                                tbl.Rows[tbl.Rows.Count].Cells[k + 2].Range.Text = vremMark.ToString() + " (" + vremBall.ToString() + "/" + MaxBall(mi) + ")";
                                Itogo[k] += vremBall;
                            }
                            kol_izm++; // считаем ненулевые измерения
                        } // оценка включена
                    } // по оценкам
                      // низ таблицы - ИТОГО
                    tbl.Rows.Add();
                    tbl.Rows[tbl.Rows.Count].Cells[1].Range.Text = "Итого баллов:";
                    for (int k = 0; k < kol_checks; k++)
                        tbl.Rows[tbl.Rows.Count].Cells[k + 2].Range.Text = "(" + Itogo[k].ToString() + " из " + TotalBall(print_sex).ToString() + ")";

                    // -----------------------ГРАФИК------------------------
                    GraphPane myPane = zedGraphControl2.GraphPane;
                    myPane.CurveList.Clear();
                    myPane.GraphObjList.Clear();
                    myPane.Legend.IsVisible = false;
                    myPane.Border.IsVisible = false;
                    myPane.Title.Text = "Оценки последнего медосмотра";
                    myPane.BarSettings.Type = BarType.PercentStack;
                    myPane.BarSettings.MinClusterGap = 0.0f;
                    myPane.XAxis.Title.IsVisible = false;
                    //myPane.YAxis.Title.IsVisible = false;
                    myPane.YAxis.Title.Text = "%";
                    myPane.YAxis.Scale.Max = 100;
                    myPane.YAxis.Title.FontSpec.Angle = 90;
                    myPane.XAxis.Scale.FontSpec.Angle = 90;
                    myPane.XAxis.Scale.FontSpec.Size = 15;
                    myPane.XAxis.Type = AxisType.Text;
                    myPane.XAxis.Scale.TextLabels = graf1;
                    PointPairList PPLa = new PointPairList();
                    PointPairList PPLb = new PointPairList();
                    for (int mi = 0; mi < graf1.Length; mi++) // цикл по сохраненным баллам
                    {
                        PPLa.Add(mi, graf2[mi]);
                        PPLb.Add(mi, graf3[mi] - graf2[mi]);
                    }
                    myPane.AddBar("+", PPLa, Color.Red);
                    myPane.AddBar("-", PPLb, Color.Pink);
                    zedGraphControl2.AxisChange();
                    zedGraphControl2.Invalidate();
                    // Получаем картинку, соответствующую панели
                    graf = myPane.GetImage();
                    Clipboard.SetImage(graf);
                    Word.Range selection = app.Selection.Range; // по умолчанию место курсора
                    for (int nb = 1; nb <= wBookmarks.Count; nb++)
                        if (wBookmarks[nb].Name == "graf") selection  = wBookmarks[nb].Range; // другое место
                    selection.Paste();
                } // если хоть один медосмотр
                app.Visible = true;
                // doc.Close(); // Закрываем документ
            } // Обследуемый найден
            else // обследуемый не найден
            {
                MessageBox.Show("Ошибка! Обследуемый с номером документа " + p_passport + " не найден !");
            }

        } // of printPassport

        //===============================================================================================================
        // Печать того, что в RichTexte
        private void button10_Click(object sender, EventArgs e)
        {
                printDialog1.Document = printDocument1;
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        printDocument1.Print();
                        //StringReader reader = new StringReader(this.richTextBox1.Text);
                        //stringToPrint = reader.ReadToEnd();
                        //this.docToPrint.PrintPage += new PrintPageEventHandler(this.docToPrintCustom);
                        //this.docToPrint.Print();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка печати!" + Environment.NewLine + "(" + ex.Message + ")");
                    }
                }
            }

        //===============================================================================================================
        // вычисляем оценку
        private double CalculateMark(int Nizm, DataRow print_medcheck)
        // print_medcheck - строка из таблицы - результата запроса по базе медосмотров (номера измерений начинаются с 13ой позиции)
        {
            double CalcMark = 0;
            String zznak;
            int nomer_izm1, nomer_izm2, multiplayer;
            int value_izm1, value_izm2;
            // находим и вычисляем значение оценки
            zznak = MySpisok.Rows[Nizm]["znak"].ToString();
            nomer_izm1 = Convert.ToInt32(MySpisok.Rows[Nizm]["izm1"].ToString());
            nomer_izm2 = Convert.ToInt32(MySpisok.Rows[Nizm]["izm2"].ToString());
            multiplayer = Convert.ToInt32(MySpisok.Rows[Nizm]["mult"].ToString());
            switch (zznak)
            {
                case "0": //копируется
                    CalcMark = Convert.ToInt32(print_medcheck[nomer_izm1 + 13].ToString()); // измерения в массиве с 13 столбца
                    break;
                case "1":// вычисляется +
                    value_izm1 = Convert.ToInt32(print_medcheck[nomer_izm1 + 13].ToString());
                    value_izm2 = Convert.ToInt32(print_medcheck[nomer_izm2 + 13].ToString());
                    CalcMark = (value_izm1 + value_izm2) * multiplayer;
                    break;
                case "2":// вычисляется -
                    value_izm1 = Convert.ToInt32(print_medcheck[nomer_izm1 + 13].ToString());
                    value_izm2 = Convert.ToInt32(print_medcheck[nomer_izm2 + 13].ToString());
                    CalcMark = (value_izm1 - value_izm2) * multiplayer;
                    break;
                case "3":// вычисляется *
                    value_izm1 = Convert.ToInt32(print_medcheck[nomer_izm1 + 13].ToString());
                    value_izm2 = Convert.ToInt32(print_medcheck[nomer_izm2 + 13].ToString());
                    CalcMark = value_izm1 * value_izm2 * multiplayer;
                    break;
                case "4":// вычисляется /
                    value_izm1 = Convert.ToInt32(print_medcheck[nomer_izm1 + 13].ToString());
                    value_izm2 = Convert.ToInt32(print_medcheck[nomer_izm2 + 13].ToString());
                    if (value_izm2 == 0)
                        CalcMark = 0;
                    else
                        CalcMark = Math.Round((float)value_izm1 / (float)value_izm2 * multiplayer, 0);
                    break;
                default:
                    // неверное значение
                    break;
            }
            return (CalcMark);
        }

        //===============================================================================================================
        // вычисляем балл
        private double CalculateBall(int Nizm, double Mark)
        {
            double CalcBall = 0;
            // точки x переводим в массив
            string x_csv = MyGraphs.Rows[Nizm]["x"].ToString();
            string[] xx = x_csv.Split(',');
            string y_csv = MyGraphs.Rows[Nizm]["y"].ToString();
            string[] yy = y_csv.Split(',');
            // найдем 2 точки, между которыми значение Y
            int index_up = 0;
            int index_down = 0;
            int otrezokY,kusokY;
            int otrezokX;
            float koef;
            bool index_found = false;
            for (int iy = 0; iy < yy.Length; iy++)
            {
                if (!index_found && Convert.ToInt32(yy[iy]) >= Mark)
                {
                    index_found = true;
                    index_up = iy;
                }
            }
            index_down = index_up - 1; // предыдущая точка
            if (index_down < 0) index_down = 0;
            otrezokY = Convert.ToInt32(yy[index_up]) - Convert.ToInt32(yy[index_down]); // разница между точками
            kusokY = Convert.ToInt32(Mark) - Convert.ToInt32(yy[index_down]); // разница между нашим значение и нижней точкой
            if (otrezokY != 0)
                koef = (float)kusokY / (float)otrezokY;
            else
                koef = 0;
            // с Yами разобрались, теперь с Xами
            otrezokX = Convert.ToInt32(xx[index_up]) - Convert.ToInt32(xx[index_down]); // разница между точками
            CalcBall = Math.Round(Convert.ToInt32(xx[index_down])+(otrezokX* koef),0);
            // если не нашли - сообщим !!!
            if (!index_found)
                MessageBox.Show("Невозможно посчитать оценку "+ MySpisok.Rows[Nizm][1]+". Значение "+ Mark+" находится за пределами графика!");
            return (CalcBall);
        }

        //===============================================================================================================
        // сравниваем М,Ж,м и ж
        // если true - значит нужный нам пол
        private bool CalculateSex(string sex1, string sex2)
        {
            bool osex = false;
            if ((sex1 == "М") && (sex2 == "М")) osex = true;
            if ((sex1 == "М") && (sex2 == "м")) osex = true;
            if ((sex1 == "м") && (sex2 == "М")) osex = true;
            if ((sex1 == "м") && (sex2 == "м")) osex = true;
            if ((sex1 == "Ж") && (sex2 == "Ж")) osex = true;
            if ((sex1 == "Ж") && (sex2 == "ж")) osex = true;
            if ((sex1 == "ж") && (sex2 == "Ж")) osex = true;
            if ((sex1 == "ж") && (sex2 == "ж")) osex = true;
            return (osex);
        }

        //===============================================================================================================
        // Форматируем строку для вывода в таблицу
        // принимает строку, добавляет пробелов, чтобы она выглядела в ячейке посередине
        // MaxSimbols - максимальное количество символов в ячейке таблицы 
        private string forma(string bilo, int MaxSimbols)
        {
            int kolSpaces = (MaxSimbols - bilo.Length) / 2; // количество пробелов для добавления перед строчкой
            string stalo=bilo.PadLeft(bilo.Length+ kolSpaces);
            return (stalo);
        }

        //===============================================================================================================
        // Форма отчета1
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            // заряжаем в таблицу названия измерений ( для отображения в комбобоксе формы6)
            System.Data.DataTable tempIzm1 = new System.Data.DataTable(); //врем.таблица
            tempIzm1.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm1.Columns.Add(new DataColumn("name", typeof(String)));
            for (int i = 0; i < MyIzm.Rows.Count; i++)
            {
                if (MyIzm.Rows[i]["вкл"].Equals(true)) // if == true (галочка стоит)
                    tempIzm1.Rows.Add(i + 1, MyIzm.Rows[i]["название"]);
            }
            Form6.Izm1table = tempIzm1;
            // заряжаем в таблицу названия оценок ( для отображения в комбобоксе формы6)
            System.Data.DataTable tempIzm2 = new System.Data.DataTable();
            tempIzm2.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm2.Columns.Add(new DataColumn("name", typeof(String)));
            for (int i = 0; i < MySpisok.Rows.Count; i++)
            {
                if (MySpisok.Rows[i]["вкл"].Equals(true)) // if == true (галочка стоит)
                    tempIzm2.Rows.Add(i + 1, MySpisok.Rows[i]["название"]+" ("+MySpisok.Rows[i]["пол"]+")");
            }
            Form6.Izm2table = tempIzm2;
            // заряжаем в таблицу названия полов ( для отображения в комбобоксе формы6)
            System.Data.DataTable tempIzm3 = new System.Data.DataTable();
            tempIzm3.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm3.Columns.Add(new DataColumn("name", typeof(String)));
            tempIzm3.Rows.Add(1, "мужчины");
            tempIzm3.Rows.Add(2, "женщины");
            Form6.Izm3table = tempIzm3;
            // заряжаем в таблицу названия фильтров( для отображения в комбобоксе формы6)
            System.Data.DataTable tempIzm4 = new System.Data.DataTable();
            System.Data.DataTable tempIzm5 = new System.Data.DataTable();
            tempIzm4.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm4.Columns.Add(new DataColumn("name", typeof(String)));
            tempIzm5.Columns.Add(new DataColumn("n", typeof(String)));
            tempIzm5.Columns.Add(new DataColumn("name", typeof(String)));
            for (int i = 0; i < MyFilters.Rows.Count; i++)
            {
                if (MyFilters.Rows[i]["вкл"].Equals(true)) // if == true (галочка стоит)
                {
                    tempIzm4.Rows.Add(i + 1, MyFilters.Rows[i]["название"]);
                    tempIzm5.Rows.Add(i + 1, MyFilters.Rows[i]["название"]);
                }
            }
            Form6.Izm4table = tempIzm4;
            Form6.Izm5table = tempIzm5;
            // копируем фильтры для отображения в форме6
            Form6.MyFilters = MyFilters;
            Form6.MyFilterValues1 = MyFilterValues1;
            Form6.MyFilterValues2 = MyFilterValues2;
            Form6.MyFilterValues3 = MyFilterValues3;
            Form6.MyFilterValues4 = MyFilterValues4;
            Form6.MyFilterValues5 = MyFilterValues5;
            Form6.MyFilterValues6 = MyFilterValues6;
            Form6.MyFilterValues7 = MyFilterValues7;
            Form6.MyFilterValues8 = MyFilterValues8;
            Form6.MyFilterValues9 = MyFilterValues9;
            Form6.MyFilterValues10 = MyFilterValues10;

            // ГОТОВО, открываем форму
            Form ConstrForm6 = new Form6();
            ConstrForm6.ShowDialog();
            if (form6Cancel == false) // если не нажималась отмена, продолжаем конструктор
                Report1(form6param1, form6param19, form6param20, form6param2, form6param3, form6param4, form6param5, form6param6, form6param7, form6param8, form6param9, form6param10, form6param11, form6param12, form6param13, form6param14, form6param15, form6param16, form6param17, form6param18);
            /*Report1("izm01","","",false, DateTime.Parse("01.01.2000"), DateTime.Parse("01.01.2000"),
                    "", "", "", "", "", "", "", "", "", "", "", "");*/
        }

        //===============================================================================================================
        // вычисляем среднеквадратичное отклонение
        /*
        Пусть оценки учеников класса следующие:
        2,4,4,4,5,5,7,9
        
        Тогда средняя оценка равна:
        (2+4+4+4+5+5+7+9) / 8 = 5
        
        Вычислим квадраты отклонений оценок учеников от их средней оценки:
        (2-5)^2=(-3)^2=9        (5-5)^2=0^2=0
        (4-5)^2=(-1)^2=1        (5-5)^2=0^2=0
        (4-5)^2=(-1)^2=1        (7-5)^2=2^2=4
        (4-5)^2=(-1)^2=1        (9-5)^2=4^2=16
        
        Среднее арифметическое этих значений называется дисперсией:
        sigma^2=(9+1+1+1+0+0+4+16) / 8 = 4
        
        Стандартное отклонение равно квадратному корню дисперсии:
        sigma =sqrt 4 = 2
        */
        private double SredneKvadratichnoe(int s_sum, int s_kol, string s_all)
        {
            double my_sigma, s_kol_double;
            string[] xx = s_all.Split(',');
            int zn, sred;
            double summ_quad = 0;
            if (s_kol > 0) sred = s_sum / s_kol; else sred = 0;
            // Вычислим сумму квадратов отклонений значений от их среднего
            for (int i = 0; i < xx.Length; i++)
            {
                if (xx[i].Length > 0)
                {
                    zn = Convert.ToInt32(xx[i]);
                    summ_quad += (zn - sred) * (zn - sred);
                }
            }
            if (s_kol > 0)
            {
                s_kol_double = s_kol;
                my_sigma = Math.Sqrt(summ_quad / s_kol_double);
            }
            else
                my_sigma = 0;
            return (my_sigma);
        }


        //===============================================================================================================
        // Подготовка отчета1
        // параметр p_izm может быть = название, потом переводим в (izm01-izm10) или пустой
        // параметр p_mark может быть = название или пустой
        // параметр f_sex может быть = м,ж или пустой
        // f_period - стоит галочка ПЕРИОД ОБСЛЕДОВАНИЯ, тогда должны быть указаны dataot и datado
        // f_n - стоит галочка ПОКАЗЫВАТЬ КОЛИЧЕСТВО, тогда в таблице выводится (n)
        // f_sigma - стоит галочка ПОКАЗЫВАТЬ СРЕДНЕКВАДРАТИЧНОЕ ОТКЛОНЕНИЕ, тогда в таблице выводится (СИГМА)
        // параметры f_filхх - если пустые, значит галочка не стоит, не фильтруем
        // параметры detal могут быть = SEX,YEAR,fil01-fil10 или пустой
        // detal1 - это строки отчета, detal2 - столбцы отчета
        public void Report1(string p_izm, bool f_n, bool f_sigma, string p_mark, string f_sex, bool f_period, DateTime dataot, DateTime datado,
            string f_fil01, string f_fil02, string f_fil03, string f_fil04, string f_fil05, string f_fil06, string f_fil07, string f_fil08, string f_fil09, string f_fil10,
            string detal1, string detal2)
        {
            string commandText;
            bool print_errors = false;
            string strRTF = "";
            string insertion;
            int ins_Start;
            System.Data.DataTable ReportData, ReportKol, Sigma;
            DataSet printDataSet;
            SQLiteDataAdapter printDataAdapter;
            SQLiteCommand cmd;

            string print_fil = "", print_title = "";
            // Загрузить шаблон в RichTextBox
            try
            {
                richTextBoxPrintCtrl1.LoadFile(dbPath + "\\" + dbName + report01RTF);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка. Шаблон для Отчета не загружен!" + Environment.NewLine + "(" + ex.Message + ")");
                print_errors = true;
            }
            // выведем название отчета
            if (!print_errors)
            {
                strRTF = richTextBoxPrintCtrl1.Rtf;
                insertion = "<#title#>";
                ins_Start = strRTF.IndexOf(insertion); // ищем ключ
                if (ins_Start != -1) // ключ найден
                {
                    strRTF = strRTF.Remove(ins_Start, insertion.Length); // удалим ключ
                    if (f_sex.Length != 0)
                    {
                        if (p_izm.Length != 0) print_title += p_izm + " (" + f_sex + ")";
                        else print_title += p_mark + " (" + f_sex + ")";
                    }
                    else if (p_izm.Length != 0) print_title += p_izm;
                    else print_title += p_mark;


                    strRTF = strRTF.Insert(ins_Start, print_title);
                }
                else // ключ не найден
                {
                    MessageBox.Show("Ошибка! В шаблоне " + report01RTF + " нет ключа " + insertion + " !");
                    print_errors = true;
                }
            }
            // выведем настройки отчета
            if (!print_errors)
            {
                insertion = "<#filter#>";
                ins_Start = strRTF.IndexOf(insertion); // ищем ключ
                if (ins_Start != -1) // ключ найден
                {
                    strRTF = strRTF.Remove(ins_Start, insertion.Length); // удалим ключ
                    if (f_fil01.Length != 0) print_fil += MyFilters.Rows[0].ItemArray[2].ToString() + " = " + f_fil01 + "   ";
                    if (f_fil02.Length != 0) print_fil += MyFilters.Rows[1].ItemArray[2].ToString() + " = " + f_fil02 + "   ";
                    if (f_fil03.Length != 0) print_fil += MyFilters.Rows[2].ItemArray[2].ToString() + " = " + f_fil03 + "   ";
                    if (f_fil04.Length != 0) print_fil += MyFilters.Rows[3].ItemArray[2].ToString() + " = " + f_fil04 + "   ";
                    if (f_fil05.Length != 0) print_fil += MyFilters.Rows[4].ItemArray[2].ToString() + " = " + f_fil05 + "   ";
                    if (f_fil06.Length != 0) print_fil += MyFilters.Rows[5].ItemArray[2].ToString() + " = " + f_fil06 + "   ";
                    if (f_fil07.Length != 0) print_fil += MyFilters.Rows[6].ItemArray[2].ToString() + " = " + f_fil07 + "   ";
                    if (f_fil08.Length != 0) print_fil += MyFilters.Rows[7].ItemArray[2].ToString() + " = " + f_fil08 + "   ";
                    if (f_fil09.Length != 0) print_fil += MyFilters.Rows[8].ItemArray[2].ToString() + " = " + f_fil09 + "   ";
                    if (f_fil10.Length != 0) print_fil += MyFilters.Rows[9].ItemArray[2].ToString() + " = " + f_fil10 + "   ";
                    if (print_fil.Length != 0) print_fil += "\\par";
                    if (f_period) print_fil += "по медосмотрам за период с " + dataot.ToString("dd-MM-yyyy") + " по " + datado.ToString("dd-MM-yyyy") + "\\par";
                    print_fil += "Среднее значение";
                    if (f_n) print_fil += " (Количество)";
                    if (f_sigma) print_fil += " Среднеквадратичное отклонение";
                    //print_fil += "\\par";
                    strRTF = strRTF.Insert(ins_Start, print_fil);
                }
                else // ключ не найден
                {
                    MessageBox.Show("Ошибка! В шаблоне " + report01RTF + " нет ключа " + insertion + " !");
                    print_errors = true;
                }
            }
            // переводим название параметра izm в формат izm01-izm20
            string n_izm = "";
            if (p_izm.Length != 0)
            {
                for (int i = 0; i < MyIzm.Rows.Count; i++)
                    if (MyIzm.Rows[i].ItemArray[2].ToString() == p_izm)
                        if (i < 9) n_izm = "izm0" + (i + 1); else n_izm = "izm" + (i + 1);
            }
            // переводим название параметра оценки в её номер в списке таблицы оценок
            int n_mark = -1;
            if (p_mark.Length != 0)
            {
                for (int i = 0; i < MySpisok.Rows.Count; i++)
                    if (MySpisok.Rows[i]["название"] + " (" + MySpisok.Rows[i]["пол"] + ")" == p_mark)
                        n_mark = i;
            }

            // ДЕТАЛИЗАЦИЯ СТОЛБЦЫ - переводим значение СТОЛБЦОВ в формат SEX,data,fil01 - fil10
            string MYdetal1 = "";
            switch (detal1)
            {
                case "":
                    MYdetal1 = "";
                    break;
                case "SEX":
                    MYdetal1 = detal1; // не меняем
                    break;
                case "YEAR": // меняем на "data"
                    MYdetal1 = "data";
                    break;
                default:
                    for (int i = 0; i < MyFilters.Rows.Count; i++)
                        if (MyFilters.Rows[i]["название"].ToString() == detal1)
                            if (i < 9) MYdetal1 = "fil0" + (i + 1); else MYdetal1 = "fil" + (i + 1);
                    break;
            }

            // ДЕТАЛИЗАЦИЯ СТРОКИ - переводим значение СТРОК в формат SEX,data,fil01 - fil10
            string MYdetal2 = "";
            switch (detal2)
            {
                case "":
                    MYdetal2 = "";
                    break;
                case "SEX":
                    MYdetal2 = detal2; // не меняем
                    break;
                case "YEAR": // меняем на "data"
                    MYdetal2 = "data";
                    break;
                default:
                    for (int i = 0; i < MyFilters.Rows.Count; i++)
                        if (MyFilters.Rows[i]["название"].ToString() == detal2)
                            if (i < 9) MYdetal2 = "fil0" + (i + 1); else MYdetal2 = "fil" + (i + 1);
                    break;
            }

            // Составляем запрос по ИЗМЕРЕНИЮ
            if (p_izm.Length != 0)
            {
                commandText = "SELECT " + n_izm;
                if (MYdetal1.Length != 0) commandText += ", " + MYdetal1;
                if (MYdetal2.Length != 0) commandText += ", " + MYdetal2;
                commandText += " FROM Student, MedCheck WHERE Student.Id = MedCheck.idStudent";
                if (f_sex.Length != 0) commandText += " AND Student.sex = '" + f_sex.Substring(0, 1) + "'";// пол, 1ая буква от слова (м или ж)
                if (f_fil01.Length != 0) commandText += " AND MedCheck.fil01 = '" + f_fil01 + "'";
                if (f_fil02.Length != 0) commandText += " AND MedCheck.fil02 = '" + f_fil02 + "'";
                if (f_fil03.Length != 0) commandText += " AND MedCheck.fil03 = '" + f_fil03 + "'";
                if (f_fil04.Length != 0) commandText += " AND MedCheck.fil04 = '" + f_fil04 + "'";
                if (f_fil05.Length != 0) commandText += " AND MedCheck.fil05 = '" + f_fil05 + "'";
                if (f_fil06.Length != 0) commandText += " AND MedCheck.fil06 = '" + f_fil06 + "'";
                if (f_fil07.Length != 0) commandText += " AND MedCheck.fil07 = '" + f_fil07 + "'";
                if (f_fil08.Length != 0) commandText += " AND MedCheck.fil08 = '" + f_fil08 + "'";
                if (f_fil09.Length != 0) commandText += " AND MedCheck.fil09 = '" + f_fil09 + "'";
                if (f_fil10.Length != 0) commandText += " AND MedCheck.fil10 = '" + f_fil10 + "'";
                if (f_period) commandText += " AND data BETWEEN '" + dataot.ToString("yyyy-MM-dd HH:mm:tt") + "' and '" + datado.ToString("yyyy-MM-dd HH:mm:tt") + "'";
                commandText += " ORDER BY ";
                if (MYdetal1.Length != 0) commandText += MYdetal1;
                if ((MYdetal1.Length != 0)&&(MYdetal2.Length != 0)) commandText += ", ";
                if (MYdetal2.Length != 0) commandText += MYdetal2;
                commandText += ";";
                cmd = new SQLiteCommand(commandText, dbConn);
                printDataSet = new DataSet();
                printDataAdapter = new SQLiteDataAdapter(cmd);
                printDataAdapter.Fill(printDataSet);

                // ОБРАБОТКА РЕЗУЛЬТАТОВ ЗАПРОСА

                // проходим по всем ячейкам таблицы и собираем в списки названия строк и столбцов
                int Stol = printDataSet.Tables[0].Columns.Count;
                int Strok = printDataSet.Tables[0].Rows.Count;
                List<string> Stolnew = new List<string>(); // в этом списке будут названия столбцов
                List<string> Stroknew = new List<string>(); // в этом списке будут названия строк
                // по результатам в таблице printDataSet.Tables[0] может быть 1,2 или 3 столбца
                // в 1ом цифры, во 2ом - столбцы, в 3ем - строки
                bool jfound;
                string jcurrent;
                for (int i = 0; i < Strok; i++) // по всем строкам
                    for (int j = 1; j < Stol; j++) // начнем с 1, а не с 0 - пропустим первый столбец (там цифры)
                    {
                        if (j == 1) // здесь названия столбцов
                        {
                            if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[i][j]))
                            {
                                jcurrent = printDataSet.Tables[0].Rows[i][j].ToString().Trim();
                                if ((detal1 == "YEAR")&&(jcurrent.Length==18))
                                    // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                    jcurrent = jcurrent.Substring(6, 4);
                            }
                            else
                                jcurrent = "";
                            jfound = false;
                            for (int x = 0; x < Stolnew.Count; x++)
                                if (Stolnew[x] == jcurrent)
                                    jfound = true;
                            if (!jfound) // если не нашли в списке столбцов
                                Stolnew.Add(jcurrent);// добавляем 
                        }
                        if (j == 2) // здесь названия строк
                        { 
                            if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[i][j]))
                            {
                                jcurrent = printDataSet.Tables[0].Rows[i][j].ToString().Trim();
                                if ((detal2 == "YEAR") && (jcurrent.Length == 18))
                                    // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                    jcurrent = jcurrent.Substring(6, 4);
                            }
                            else
                                jcurrent = "";
                            jfound = false;
                            for (int x = 0; x < Stroknew.Count; x++)
                                if (Stroknew[x] == jcurrent)
                                    jfound = true;
                            if (!jfound) // если не нашли в списке строк
                                Stroknew.Add(jcurrent);// добавляем 
                        }
                    }
                if (Stroknew.Count == 0)
                    Stroknew.Add("");// добавляем 
                if (Stolnew.Count > 5)
                {
                    MessageBox.Show("Ошибка! В отчете получается более 5 столбцов (" + Stolnew.Count + "). Попробуйте в настройках детализации поменять строки и столбцы местами!");
                    print_errors = true;
                }

                // Создаем 3 параллельные таблицы и массив
                ReportData = new System.Data.DataTable(); // в этой будут суммы средних значений
                ReportKol = new System.Data.DataTable();  // в этой будут количества
                Sigma = new System.Data.DataTable();  // в этой будут список всех значений в строковом формате через запятую
                // Добавляем столбцы
                for (int j = 0; j < Stolnew.Count; j++)
                {
                    ReportData.Columns.Add(new DataColumn(Stolnew[j], typeof(int)));
                    ReportKol.Columns.Add(new DataColumn(Stolnew[j], typeof(int)));
                    Sigma.Columns.Add(new DataColumn(Stolnew[j], typeof(double)));
                }
                // Цикл по количеству будущих строк
                string xcurrent, ycurrent;
                int s_sum, s_kol;
                string s_all;
                for (int i = 0; i < Stroknew.Count; i++)
                {
                    ReportData.Rows.Add();
                    ReportKol.Rows.Add();
                    Sigma.Rows.Add();
                    for (int j = 0; j < Stolnew.Count; j++)
                    {
                        s_sum = 0;
                        s_kol = 0;
                        s_all = "";
                        for (int y = 0; y < Strok; y++) // цикл по всем строкам таблицы запроса
                        {
                            xcurrent = "";
                            ycurrent = "";
                            if (Stol > 1) // если столбец есть
                                if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[y][1]))
                                {
                                    // запоминаем значения 2ого столбца в этой строке (1ый - нулевой!!!)
                                    xcurrent = printDataSet.Tables[0].Rows[y][1].ToString().Trim();
                                    if ((detal1 == "YEAR") && (xcurrent.Length == 18))
                                        // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                        xcurrent = xcurrent.Substring(6, 4);
                                }
                            if (Stol > 2) // если столбец есть
                                if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[y][2]))
                                {
                                    // запоминаем значения 3его столбца в этой строке
                                    ycurrent = printDataSet.Tables[0].Rows[y][2].ToString().Trim();
                                    if ((detal2 == "YEAR") && (ycurrent.Length == 18))
                                        // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                        ycurrent = ycurrent.Substring(6, 4);
                                }
                            if ((Stroknew[i] == ycurrent) && (Stolnew[j] == xcurrent))
                            {
                                if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[y][0]))
                                {
                                    s_sum += Convert.ToInt32(printDataSet.Tables[0].Rows[y][0].ToString());
                                    s_kol += 1;
                                    if (f_sigma) // если галочка не стоит - не тратим время на подсчет
                                        s_all += printDataSet.Tables[0].Rows[y][0].ToString()+",";
                                }
                            }
                        }
                        ReportData.Rows[i][j] = s_sum;
                        ReportKol.Rows[i][j] = s_kol;
                        if (f_sigma) // если галочка не стоит - не тратим время на подсчет
                            Sigma.Rows[i][j] = SredneKvadratichnoe(s_sum,s_kol,s_all);
                    }
                }

                // ВЫВОДИМ ОТЧЕТ
                if (!print_errors)
                {
                    insertion = "<#tab#>";
                    ins_Start = strRTF.IndexOf(insertion); // ищем ключ
                    if (ins_Start != -1) // ключ найден
                    {
                        strRTF = strRTF.Remove(ins_Start, insertion.Length); // удалим ключ
                        if (Stroknew.Count > 0) // если есть хоть 1 
                        {
                            tableRtf = new StringBuilder();
                            tableRtf.Append(@"\trowd");
                            tableRtf.Append(@"\trql"); // trqc - по центру
                            tableRtf.Append(@"\cellx" + passport_wide.ToString());// столбец на название оценки
                            int cells = passport_wide;
                            for (int k = 0; k < Stolnew.Count; k++)
                            {
                                cells += passport_wide2; // ширина колонок
                                tableRtf.Append(@"\cellx" + cells.ToString()); // добавим столько столбцов, сколько колонок
                            }
                            // шапка таблицы
                            string print_name_izm="";
                            if (detal2.Length > 0)
                                print_name_izm = detal2 + " / " + detal1;
                            else
                                print_name_izm = detal1;
                            // заменим SEX на Пол
                            int x = print_name_izm.IndexOf("SEX");
                            if (x != -1)
                            {
                                print_name_izm = print_name_izm.Remove(x, 3); // удалим 3 символа
                                print_name_izm = print_name_izm.Insert(x, "Пол");
                            }
                            // заменим YEAR на Год
                            x = print_name_izm.IndexOf("YEAR");
                            if (x != -1)
                            {
                                print_name_izm = print_name_izm.Remove(x, 4); // удалим 3 символа
                                print_name_izm = print_name_izm.Insert(x, "Год");
                            }
                            string print_col1;
                            string print_col2;
                            string print_col3;
                            string print_col4;
                            string print_col5;
                            switch (Stolnew.Count)
                            {
                                case 0:
                                    MessageBox.Show("Пустой отчет, нет данных! Измените условия фильтрации!");
                                    break;
                                case 1:
                                    print_col1 = Stolnew[0];
                                    if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell\row", print_name_izm, print_col1));
                                    break;
                                case 2:
                                    print_col1 = Stolnew[0];
                                    if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                    print_col2 = Stolnew[1];
                                    if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell\row", print_name_izm, print_col1, print_col2));
                                    break;
                                case 3:
                                    print_col1 = Stolnew[0];
                                    if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                    print_col2 = Stolnew[1];
                                    if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                    print_col3 = Stolnew[2];
                                    if (print_col3.Length > 10) print_col3 = print_col3.Substring(0, 10);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell        {3}\cell\row", print_name_izm, print_col1, print_col2, print_col3));
                                    break;
                                case 4:
                                    print_col1 = Stolnew[0];
                                    if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                    print_col2 = Stolnew[1];
                                    if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                    print_col3 = Stolnew[2];
                                    if (print_col3.Length > 10) print_col3 = print_col3.Substring(0, 10);
                                    print_col4 = Stolnew[3];
                                    if (print_col4.Length > 10) print_col4 = print_col4.Substring(0, 10);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell        {3}\cell        {4}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4));
                                    break;
                                case 5:
                                    print_col1 = Stolnew[0];
                                    if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                    print_col2 = Stolnew[1];
                                    if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                    print_col3 = Stolnew[2];
                                    if (print_col3.Length > 10) print_col3 = print_col3.Substring(0, 10);
                                    print_col4 = Stolnew[3];
                                    if (print_col4.Length > 10) print_col4 = print_col4.Substring(0, 10);
                                    print_col5 = Stolnew[4];
                                    if (print_col5.Length > 10) print_col5 = print_col5.Substring(0, 10);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell        {3}\cell        {4}\cell        {5}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4, print_col5));
                                    break;
                                default:
                                    MessageBox.Show("Ошибка! Столбцов больше 5ти!");
                                    print_errors = true;
                                    break;
                            }
                            // тело таблицы
                            int v1, v2, v3;
                            double v4=0;
                            int B0 = 0, B1 = 0, B2 = 0, B3 = 0, B4 = 0; // Итоговые суммы
                            for (int mi = 0; mi < Stroknew.Count; mi++) // цикл по строкам
                            {
                                print_name_izm = Stroknew[mi];
                                print_col1 = "";
                                print_col2 = "";
                                print_col3 = "";
                                print_col4 = "";
                                print_col5 = "";
                                for (int k = 0; k < Stolnew.Count; k++)
                                {
                                    v1 = Convert.ToInt32(ReportData.Rows[mi][k].ToString());
                                    v2 = Convert.ToInt32(ReportKol.Rows[mi][k].ToString());
                                    if (f_sigma) 
                                        v4 = Convert.ToDouble(Sigma.Rows[mi][k].ToString());
                                    // проверка деления на 0
                                    if (v2 > 0) v3 = v1 / v2; else v3 = 0;
                                    if (f_sigma) 
                                        v4 = Math.Round(v4, 2);
                                    if (k == 0)
                                    {
                                        B0 += v2;
                                        print_col1 = v3.ToString();
                                        if (f_n)        print_col1 += " (" + v2.ToString() + ")";
                                        if (f_sigma)    print_col1 += " " + v4.ToString();
                                        print_col1 = forma(print_col1, 22);
                                    }
                                    if (k == 1)
                                    {
                                        B1 += v2;
                                        print_col2 = v3.ToString();
                                        if (f_n)        print_col2 += " (" + v2.ToString() + ")";
                                        if (f_sigma)    print_col2 += " " + v4.ToString();
                                        print_col2 = forma(print_col2, 22);
                                    }
                                    if (k == 2)
                                    {
                                        B2 += v2;
                                        print_col3 = v3.ToString();
                                        if (f_n)        print_col3 += " (" + v2.ToString() + ")";
                                        if (f_sigma)    print_col3 += " " + v4.ToString();
                                        print_col3 = forma(print_col3, 22);
                                    }
                                    if (k == 3)
                                    {
                                        B3 += v2;
                                        print_col4 = v3.ToString();
                                        if (f_n)        print_col4 += " (" + v2.ToString() + ")";
                                        if (f_sigma)    print_col4 += " " + v4.ToString();
                                        print_col4 = forma(print_col4, 22);
                                    }
                                    if (k == 4)
                                    {
                                        B4 += v2;
                                        print_col5 = v3.ToString() + " (" + v2.ToString() + ")";
                                        if (f_n)        print_col4 += " (" + v2.ToString() + ")";
                                        if (f_sigma)    print_col4 += " " + v4.ToString();
                                        print_col5 = forma(print_col5, 22);
                                    }
                                }
                                switch (Stolnew.Count)
                                {
                                    case 0:
                                        break;
                                    case 1:
                                        tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell\row", print_name_izm, print_col1));
                                        break;
                                    case 2:
                                        tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell\row", print_name_izm, print_col1, print_col2));
                                        break;
                                    case 3:
                                        tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell\row", print_name_izm, print_col1, print_col2, print_col3));
                                        break;
                                    case 4:
                                        tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4));
                                        break;
                                    case 5:
                                        tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell {5}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4, print_col5));
                                        break;
                                    default:
                                        print_errors = true;
                                        break;
                                }
                            } // по строкам
                            // низ таблицы - ИТОГО
                            print_name_izm = "Итого количество:";
                            switch (Stolnew.Count)
                            {
                                case 0:
                                    break;
                                case 1:
                                    print_col1 = forma(B0.ToString(), 22);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell\row", print_name_izm, print_col1));
                                    break;
                                case 2:
                                    print_col1 = forma(B0.ToString(), 22);
                                    print_col2 = forma(B1.ToString(), 22);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell\row", print_name_izm, print_col1, print_col2));
                                    break;
                                case 3:
                                    print_col1 = forma(B0.ToString(), 22);
                                    print_col2 = forma(B1.ToString(), 22);
                                    print_col3 = forma(B2.ToString(), 22);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell\row", print_name_izm, print_col1, print_col2, print_col3));
                                    break;
                                case 4:
                                    print_col1 = forma(B0.ToString(), 22);
                                    print_col2 = forma(B1.ToString(), 22);
                                    print_col3 = forma(B2.ToString(), 22);
                                    print_col4 = forma(B3.ToString(), 22);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4));
                                    break;
                                case 5:
                                    print_col1 = forma(B0.ToString(), 22);
                                    print_col2 = forma(B1.ToString(), 22);
                                    print_col3 = forma(B2.ToString(), 22);
                                    print_col4 = forma(B3.ToString(), 22);
                                    print_col5 = forma(B4.ToString(), 22);
                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell {5}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4, print_col5));
                                    break;
                                default:
                                    print_errors = true;
                                    break;
                            }
                            tableRtf.Append(@"\pard");
                            strRTF = strRTF.Insert(ins_Start, tableRtf.ToString()); // таблица готова - вставляем
                        } // если хоть один медосмотр
                    }
                    else // ключ не найден
                    {
                        MessageBox.Show("Ошибка! В шаблоне " + report01RTF + " нет ключа " + insertion + " !");
                        print_errors = true;
                    }
                } // проверки print_errors
                // готово
                if (!print_errors)
                {
                    richTextBoxPrintCtrl1.Rtf = strRTF; //сохраняем изменения в форму
                    button10.Visible = true;
                }
            }
            // Составляем запрос по ОЦЕНКЕ
            else MessageBox.Show("Обратитесь к разработчику!");
        }

        //===============================================================================================================
        //************************************************

        //===============================================================================================================
        // Подготовка отчета2
        // параметр f_sex может быть = м,ж или пустой
        // f_period - стоит галочка ПЕРИОД ОБСЛЕДОВАНИЯ, тогда должны быть указаны dataot и datado
        // параметры f_filхх - если пустые, значит галочка не стоит, не фильтруем
        // таблица Coltab1 - это копия таблицы измерений (отключенные удалены), где галочками выделены объекты, которые будут считаться в столбцах отчета
        // таблица Coltab2 - это копия таблицы оценок    (отключенные удалены), где галочками выделены объекты, которые будут считаться в столбцах отчета
        public void Report2(string f_sex, bool f_period, DateTime dataot, DateTime datado,
            string f_fil01, string f_fil02, string f_fil03, string f_fil04, string f_fil05, string f_fil06, string f_fil07, string f_fil08, string f_fil09, string f_fil10,
            DataTable Coltab1, DataTable Coltab2)
        {
            string commandText;
            bool print_errors = false;
            string strRTF = "";
            string insertion;
            int ins_Start;
            System.Data.DataTable ReportData, ReportKol, Sigma;
            DataSet printDataSet;
            SQLiteDataAdapter printDataAdapter;
            SQLiteCommand cmd;

            string print_fil = "", print_title = "";
            // Загрузить шаблон в RichTextBox
            try
            {
                richTextBoxPrintCtrl1.LoadFile(dbPath + "\\" + dbName + report01RTF);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка. Шаблон для Отчета не загружен!" + Environment.NewLine + "(" + ex.Message + ")");
                print_errors = true;
            }
            // выведем название отчета
            if (!print_errors)
            {
                strRTF = richTextBoxPrintCtrl1.Rtf;
                insertion = "<#title#>";
                ins_Start = strRTF.IndexOf(insertion); // ищем ключ
                if (ins_Start != -1) // ключ найден
                {
                    strRTF = strRTF.Remove(ins_Start, insertion.Length); // удалим ключ
                    print_title = "Список обследований";
                    strRTF = strRTF.Insert(ins_Start, print_title);
                }
                else // ключ не найден
                {
                    MessageBox.Show("Ошибка! В шаблоне " + report01RTF + " нет ключа " + insertion + " !");
                    print_errors = true;
                }
            }
            // выведем настройки отчета
            if (!print_errors)
            {
                insertion = "<#filter#>";
                ins_Start = strRTF.IndexOf(insertion); // ищем ключ
                if (ins_Start != -1) // ключ найден
                {
                    strRTF = strRTF.Remove(ins_Start, insertion.Length); // удалим ключ
                    if (f_fil01.Length != 0) print_fil += MyFilters.Rows[0].ItemArray[2].ToString() + " = " + f_fil01 + "   ";
                    if (f_fil02.Length != 0) print_fil += MyFilters.Rows[1].ItemArray[2].ToString() + " = " + f_fil02 + "   ";
                    if (f_fil03.Length != 0) print_fil += MyFilters.Rows[2].ItemArray[2].ToString() + " = " + f_fil03 + "   ";
                    if (f_fil04.Length != 0) print_fil += MyFilters.Rows[3].ItemArray[2].ToString() + " = " + f_fil04 + "   ";
                    if (f_fil05.Length != 0) print_fil += MyFilters.Rows[4].ItemArray[2].ToString() + " = " + f_fil05 + "   ";
                    if (f_fil06.Length != 0) print_fil += MyFilters.Rows[5].ItemArray[2].ToString() + " = " + f_fil06 + "   ";
                    if (f_fil07.Length != 0) print_fil += MyFilters.Rows[6].ItemArray[2].ToString() + " = " + f_fil07 + "   ";
                    if (f_fil08.Length != 0) print_fil += MyFilters.Rows[7].ItemArray[2].ToString() + " = " + f_fil08 + "   ";
                    if (f_fil09.Length != 0) print_fil += MyFilters.Rows[8].ItemArray[2].ToString() + " = " + f_fil09 + "   ";
                    if (f_fil10.Length != 0) print_fil += MyFilters.Rows[9].ItemArray[2].ToString() + " = " + f_fil10 + "   ";
                    if (print_fil.Length != 0) print_fil += "\\par";
                    if (f_period) print_fil += "по медосмотрам за период с " + dataot.ToString("dd-MM-yyyy") + " по " + datado.ToString("dd-MM-yyyy") + "\\par";
                    print_fil += "Среднее значение";
                    strRTF = strRTF.Insert(ins_Start, print_fil);
                }
                else // ключ не найден
                {
                    MessageBox.Show("Ошибка! В шаблоне " + report01RTF + " нет ключа " + insertion + " !");
                    print_errors = true;
                }
            }
            // Составляем запрос
            commandText = "SELECT f,i,o ";
            // добавляем все измерения, даже которые не выводятся, они могут понадобиться в расчете оценки
            for (int i = 0; i < Coltab1.Rows.Count; i++) 
                commandText += ", "+ Coltab1.Rows[i]["izm"].ToString();
            commandText += " FROM Student, MedCheck WHERE Student.Id = MedCheck.idStudent";
                            if (f_sex.Length != 0) commandText += " AND Student.sex = '" + f_sex.Substring(0, 1) + "'";// пол, 1ая буква от слова (м или ж)
                            if (f_fil01.Length != 0) commandText += " AND MedCheck.fil01 = '" + f_fil01 + "'";
                            if (f_fil02.Length != 0) commandText += " AND MedCheck.fil02 = '" + f_fil02 + "'";
                            if (f_fil03.Length != 0) commandText += " AND MedCheck.fil03 = '" + f_fil03 + "'";
                            if (f_fil04.Length != 0) commandText += " AND MedCheck.fil04 = '" + f_fil04 + "'";
                            if (f_fil05.Length != 0) commandText += " AND MedCheck.fil05 = '" + f_fil05 + "'";
                            if (f_fil06.Length != 0) commandText += " AND MedCheck.fil06 = '" + f_fil06 + "'";
                            if (f_fil07.Length != 0) commandText += " AND MedCheck.fil07 = '" + f_fil07 + "'";
                            if (f_fil08.Length != 0) commandText += " AND MedCheck.fil08 = '" + f_fil08 + "'";
                            if (f_fil09.Length != 0) commandText += " AND MedCheck.fil09 = '" + f_fil09 + "'";
                            if (f_fil10.Length != 0) commandText += " AND MedCheck.fil10 = '" + f_fil10 + "'";
                            if (f_period) commandText += " AND data BETWEEN '" + dataot.ToString("yyyy-MM-dd HH:mm:tt") + "' and '" + datado.ToString("yyyy-MM-dd HH:mm:tt") + "'";
                            /* commandText += " ORDER BY ";
                            if (MYdetal1.Length != 0) commandText += MYdetal1;
                            if ((MYdetal1.Length != 0) && (MYdetal2.Length != 0)) commandText += ", ";
                            if (MYdetal2.Length != 0) commandText += MYdetal2;*/
                            commandText += ";";
                            cmd = new SQLiteCommand(commandText, dbConn);
                            printDataSet = new DataSet();
                            printDataAdapter = new SQLiteDataAdapter(cmd);
                            printDataAdapter.Fill(printDataSet);

                            // ОБРАБОТКА РЕЗУЛЬТАТОВ ЗАПРОСА

                            // проходим по всем ячейкам таблицы и собираем в списки названия строк и столбцов
                            int Stol = printDataSet.Tables[0].Columns.Count;
                            int Strok = printDataSet.Tables[0].Rows.Count;
                            List<string> Stolnew = new List<string>(); // в этом списке будут названия столбцов
                            List<string> Stroknew = new List<string>(); // в этом списке будут названия строк
                            // по результатам в таблице printDataSet.Tables[0] может быть 1,2 или 3 столбца
                            // в 1ом цифры, во 2ом - столбцы, в 3ем - строки
                            bool jfound;
                            string jcurrent;
                            for (int i = 0; i < Strok; i++) // по всем строкам
                                for (int j = 1; j < Stol; j++) // начнем с 1, а не с 0 - пропустим первый столбец (там цифры)
                                {
                                    if (j == 1) // здесь названия столбцов
                                    {
                                        if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[i][j]))
                                        {
                                            jcurrent = printDataSet.Tables[0].Rows[i][j].ToString().Trim();
                                            if ((detal1 == "YEAR") && (jcurrent.Length == 18))
                                                // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                                jcurrent = jcurrent.Substring(6, 4);
                                        }
                                        else
                                            jcurrent = "";
                                        jfound = false;
                                        for (int x = 0; x < Stolnew.Count; x++)
                                            if (Stolnew[x] == jcurrent)
                                                jfound = true;
                                        if (!jfound) // если не нашли в списке столбцов
                                            Stolnew.Add(jcurrent);// добавляем 
                                    }
                                    if (j == 2) // здесь названия строк
                                    {
                                        if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[i][j]))
                                        {
                                            jcurrent = printDataSet.Tables[0].Rows[i][j].ToString().Trim();
                                            if ((detal2 == "YEAR") && (jcurrent.Length == 18))
                                                // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                                jcurrent = jcurrent.Substring(6, 4);
                                        }
                                        else
                                            jcurrent = "";
                                        jfound = false;
                                        for (int x = 0; x < Stroknew.Count; x++)
                                            if (Stroknew[x] == jcurrent)
                                                jfound = true;
                                        if (!jfound) // если не нашли в списке строк
                                            Stroknew.Add(jcurrent);// добавляем 
                                    }
                                }
                            if (Stroknew.Count == 0)
                                Stroknew.Add("");// добавляем 
                            if (Stolnew.Count > 5)
                            {
                                MessageBox.Show("Ошибка! В отчете получается более 5 столбцов (" + Stolnew.Count + "). Попробуйте в настройках детализации поменять строки и столбцы местами!");
                                print_errors = true;
                            }

                            // Создаем 3 параллельные таблицы и массив
                            ReportData = new System.Data.DataTable(); // в этой будут суммы средних значений
                            ReportKol = new System.Data.DataTable();  // в этой будут количества
                            Sigma = new System.Data.DataTable();  // в этой будут список всех значений в строковом формате через запятую
                            // Добавляем столбцы
                            for (int j = 0; j < Stolnew.Count; j++)
                            {
                                ReportData.Columns.Add(new DataColumn(Stolnew[j], typeof(int)));
                                ReportKol.Columns.Add(new DataColumn(Stolnew[j], typeof(int)));
                                Sigma.Columns.Add(new DataColumn(Stolnew[j], typeof(double)));
                            }
                            // Цикл по количеству будущих строк
                            string xcurrent, ycurrent;
                            int s_sum, s_kol;
                            string s_all;
                            for (int i = 0; i < Stroknew.Count; i++)
                            {
                                ReportData.Rows.Add();
                                ReportKol.Rows.Add();
                                Sigma.Rows.Add();
                                for (int j = 0; j < Stolnew.Count; j++)
                                {
                                    s_sum = 0;
                                    s_kol = 0;
                                    s_all = "";
                                    for (int y = 0; y < Strok; y++) // цикл по всем строкам таблицы запроса
                                    {
                                        xcurrent = "";
                                        ycurrent = "";
                                        if (Stol > 1) // если столбец есть
                                            if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[y][1]))
                                            {
                                                // запоминаем значения 2ого столбца в этой строке (1ый - нулевой!!!)
                                                xcurrent = printDataSet.Tables[0].Rows[y][1].ToString().Trim();
                                                if ((detal1 == "YEAR") && (xcurrent.Length == 18))
                                                    // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                                    xcurrent = xcurrent.Substring(6, 4);
                                            }
                                        if (Stol > 2) // если столбец есть
                                            if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[y][2]))
                                            {
                                                // запоминаем значения 3его столбца в этой строке
                                                ycurrent = printDataSet.Tables[0].Rows[y][2].ToString().Trim();
                                                if ((detal2 == "YEAR") && (ycurrent.Length == 18))
                                                    // меняем в таблице временные значения на год. например, ""01.01.2018 0:00:00"" меняем на "2018"
                                                    ycurrent = ycurrent.Substring(6, 4);
                                            }
                                        if ((Stroknew[i] == ycurrent) && (Stolnew[j] == xcurrent))
                                        {
                                            if (!Convert.IsDBNull(printDataSet.Tables[0].Rows[y][0]))
                                            {
                                                s_sum += Convert.ToInt32(printDataSet.Tables[0].Rows[y][0].ToString());
                                                s_kol += 1;
                                                if (f_sigma) // если галочка не стоит - не тратим время на подсчет
                                                    s_all += printDataSet.Tables[0].Rows[y][0].ToString() + ",";
                                            }
                                        }
                                    }
                                    ReportData.Rows[i][j] = s_sum;
                                    ReportKol.Rows[i][j] = s_kol;
                                    if (f_sigma) // если галочка не стоит - не тратим время на подсчет
                                        Sigma.Rows[i][j] = SredneKvadratichnoe(s_sum, s_kol, s_all);
                                }
                            }

                            // ВЫВОДИМ ОТЧЕТ
                            if (!print_errors)
                            {
                                insertion = "<#tab#>";
                                ins_Start = strRTF.IndexOf(insertion); // ищем ключ
                                if (ins_Start != -1) // ключ найден
                                {
                                    strRTF = strRTF.Remove(ins_Start, insertion.Length); // удалим ключ
                                    if (Stroknew.Count > 0) // если есть хоть 1 
                                    {
                                        tableRtf = new StringBuilder();
                                        tableRtf.Append(@"\trowd");
                                        tableRtf.Append(@"\trql"); // trqc - по центру
                                        tableRtf.Append(@"\cellx" + passport_wide.ToString());// столбец на название оценки
                                        int cells = passport_wide;
                                        for (int k = 0; k < Stolnew.Count; k++)
                                        {
                                            cells += passport_wide2; // ширина колонок
                                            tableRtf.Append(@"\cellx" + cells.ToString()); // добавим столько столбцов, сколько колонок
                                        }
                                        // шапка таблицы
                                        string print_name_izm = "";
                                        if (detal2.Length > 0)
                                            print_name_izm = detal2 + " / " + detal1;
                                        else
                                            print_name_izm = detal1;
                                        // заменим SEX на Пол
                                        int x = print_name_izm.IndexOf("SEX");
                                        if (x != -1)
                                        {
                                            print_name_izm = print_name_izm.Remove(x, 3); // удалим 3 символа
                                            print_name_izm = print_name_izm.Insert(x, "Пол");
                                        }
                                        // заменим YEAR на Год
                                        x = print_name_izm.IndexOf("YEAR");
                                        if (x != -1)
                                        {
                                            print_name_izm = print_name_izm.Remove(x, 4); // удалим 3 символа
                                            print_name_izm = print_name_izm.Insert(x, "Год");
                                        }
                                        string print_col1;
                                        string print_col2;
                                        string print_col3;
                                        string print_col4;
                                        string print_col5;
                                        switch (Stolnew.Count)
                                        {
                                            case 0:
                                                MessageBox.Show("Пустой отчет, нет данных! Измените условия фильтрации!");
                                                break;
                                            case 1:
                                                print_col1 = Stolnew[0];
                                                if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell\row", print_name_izm, print_col1));
                                                break;
                                            case 2:
                                                print_col1 = Stolnew[0];
                                                if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                                print_col2 = Stolnew[1];
                                                if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell\row", print_name_izm, print_col1, print_col2));
                                                break;
                                            case 3:
                                                print_col1 = Stolnew[0];
                                                if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                                print_col2 = Stolnew[1];
                                                if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                                print_col3 = Stolnew[2];
                                                if (print_col3.Length > 10) print_col3 = print_col3.Substring(0, 10);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell        {3}\cell\row", print_name_izm, print_col1, print_col2, print_col3));
                                                break;
                                            case 4:
                                                print_col1 = Stolnew[0];
                                                if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                                print_col2 = Stolnew[1];
                                                if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                                print_col3 = Stolnew[2];
                                                if (print_col3.Length > 10) print_col3 = print_col3.Substring(0, 10);
                                                print_col4 = Stolnew[3];
                                                if (print_col4.Length > 10) print_col4 = print_col4.Substring(0, 10);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell        {3}\cell        {4}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4));
                                                break;
                                            case 5:
                                                print_col1 = Stolnew[0];
                                                if (print_col1.Length > 10) print_col1 = print_col1.Substring(0, 10);
                                                print_col2 = Stolnew[1];
                                                if (print_col2.Length > 10) print_col2 = print_col2.Substring(0, 10);
                                                print_col3 = Stolnew[2];
                                                if (print_col3.Length > 10) print_col3 = print_col3.Substring(0, 10);
                                                print_col4 = Stolnew[3];
                                                if (print_col4.Length > 10) print_col4 = print_col4.Substring(0, 10);
                                                print_col5 = Stolnew[4];
                                                if (print_col5.Length > 10) print_col5 = print_col5.Substring(0, 10);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell       {1}\cell        {2}\cell        {3}\cell        {4}\cell        {5}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4, print_col5));
                                                break;
                                            default:
                                                MessageBox.Show("Ошибка! Столбцов больше 5ти!");
                                                print_errors = true;
                                                break;
                                        }
                                        // тело таблицы
                                        int v1, v2, v3;
                                        double v4 = 0;
                                        int B0 = 0, B1 = 0, B2 = 0, B3 = 0, B4 = 0; // Итоговые суммы
                                        for (int mi = 0; mi < Stroknew.Count; mi++) // цикл по строкам
                                        {
                                            print_name_izm = Stroknew[mi];
                                            print_col1 = "";
                                            print_col2 = "";
                                            print_col3 = "";
                                            print_col4 = "";
                                            print_col5 = "";
                                            for (int k = 0; k < Stolnew.Count; k++)
                                            {
                                                v1 = Convert.ToInt32(ReportData.Rows[mi][k].ToString());
                                                v2 = Convert.ToInt32(ReportKol.Rows[mi][k].ToString());
                                                if (f_sigma)
                                                    v4 = Convert.ToDouble(Sigma.Rows[mi][k].ToString());
                                                // проверка деления на 0
                                                if (v2 > 0) v3 = v1 / v2; else v3 = 0;
                                                if (f_sigma)
                                                    v4 = Math.Round(v4, 2);
                                                if (k == 0)
                                                {
                                                    B0 += v2;
                                                    print_col1 = v3.ToString();
                                                    if (f_n) print_col1 += " (" + v2.ToString() + ")";
                                                    if (f_sigma) print_col1 += " " + v4.ToString();
                                                    print_col1 = forma(print_col1, 22);
                                                }
                                                if (k == 1)
                                                {
                                                    B1 += v2;
                                                    print_col2 = v3.ToString();
                                                    if (f_n) print_col2 += " (" + v2.ToString() + ")";
                                                    if (f_sigma) print_col2 += " " + v4.ToString();
                                                    print_col2 = forma(print_col2, 22);
                                                }
                                                if (k == 2)
                                                {
                                                    B2 += v2;
                                                    print_col3 = v3.ToString();
                                                    if (f_n) print_col3 += " (" + v2.ToString() + ")";
                                                    if (f_sigma) print_col3 += " " + v4.ToString();
                                                    print_col3 = forma(print_col3, 22);
                                                }
                                                if (k == 3)
                                                {
                                                    B3 += v2;
                                                    print_col4 = v3.ToString();
                                                    if (f_n) print_col4 += " (" + v2.ToString() + ")";
                                                    if (f_sigma) print_col4 += " " + v4.ToString();
                                                    print_col4 = forma(print_col4, 22);
                                                }
                                                if (k == 4)
                                                {
                                                    B4 += v2;
                                                    print_col5 = v3.ToString() + " (" + v2.ToString() + ")";
                                                    if (f_n) print_col4 += " (" + v2.ToString() + ")";
                                                    if (f_sigma) print_col4 += " " + v4.ToString();
                                                    print_col5 = forma(print_col5, 22);
                                                }
                                            }
                                            switch (Stolnew.Count)
                                            {
                                                case 0:
                                                    break;
                                                case 1:
                                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell\row", print_name_izm, print_col1));
                                                    break;
                                                case 2:
                                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell\row", print_name_izm, print_col1, print_col2));
                                                    break;
                                                case 3:
                                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell\row", print_name_izm, print_col1, print_col2, print_col3));
                                                    break;
                                                case 4:
                                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4));
                                                    break;
                                                case 5:
                                                    tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell {5}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4, print_col5));
                                                    break;
                                                default:
                                                    print_errors = true;
                                                    break;
                                            }
                                        } // по строкам
                                        // низ таблицы - ИТОГО
                                        print_name_izm = "Итого количество:";
                                        switch (Stolnew.Count)
                                        {
                                            case 0:
                                                break;
                                            case 1:
                                                print_col1 = forma(B0.ToString(), 22);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell\row", print_name_izm, print_col1));
                                                break;
                                            case 2:
                                                print_col1 = forma(B0.ToString(), 22);
                                                print_col2 = forma(B1.ToString(), 22);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell\row", print_name_izm, print_col1, print_col2));
                                                break;
                                            case 3:
                                                print_col1 = forma(B0.ToString(), 22);
                                                print_col2 = forma(B1.ToString(), 22);
                                                print_col3 = forma(B2.ToString(), 22);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell\row", print_name_izm, print_col1, print_col2, print_col3));
                                                break;
                                            case 4:
                                                print_col1 = forma(B0.ToString(), 22);
                                                print_col2 = forma(B1.ToString(), 22);
                                                print_col3 = forma(B2.ToString(), 22);
                                                print_col4 = forma(B3.ToString(), 22);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4));
                                                break;
                                            case 5:
                                                print_col1 = forma(B0.ToString(), 22);
                                                print_col2 = forma(B1.ToString(), 22);
                                                print_col3 = forma(B2.ToString(), 22);
                                                print_col4 = forma(B3.ToString(), 22);
                                                print_col5 = forma(B4.ToString(), 22);
                                                tableRtf.Append(String.Format(@"\intbl       {0}\cell {1}\cell {2}\cell {3}\cell {4}\cell {5}\cell\row", print_name_izm, print_col1, print_col2, print_col3, print_col4, print_col5));
                                                break;
                                            default:
                                                print_errors = true;
                                                break;
                                        }
                                        tableRtf.Append(@"\pard");
                                        strRTF = strRTF.Insert(ins_Start, tableRtf.ToString()); // таблица готова - вставляем
                                    } // если хоть один медосмотр
                                }
                                else // ключ не найден
                                {
                                    MessageBox.Show("Ошибка! В шаблоне " + report01RTF + " нет ключа " + insertion + " !");
                                    print_errors = true;
                                }
                            } // проверки print_errors
                            // готово
                            if (!print_errors)
                            {
                                richTextBoxPrintCtrl1.Rtf = strRTF; //сохраняем изменения в форму
                                button10.Visible = true;
                            }
                    }

                    //===============================================================================================================
                    //************************************************


            */


        }

        }//classForm
}//FormApp
