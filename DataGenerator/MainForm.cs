using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DataGenerator.Properties;
using Microsoft.Office.Interop.Excel;
using static DataGenerator.SettingsOip;
using static DataGenerator.SettingsTimer;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;

namespace DataGenerator
{
    enum TableOperations
    {
        Delete,
        Insert,
        Copy,
        Paste
    } 

    public partial class MainForm : Form
    {
        #region Переменные
        private Excel _fileIo; //файл информационный
        private Excel _fileMsg; //файл сообщений
        private OleDb _oleDb; //оюъект базы данных
        private SqlDb _sqlDb; //оюъект базы данных
        private OleDbDataAdapter _daOip; //адаптер к БД в OIP
        private OleDbDataAdapter _daOipType; //адаптер к БД в OIPType
        private OleDbDataAdapter _daPlaceOip; //адаптер к БД в 
        private OleDbDataAdapter _daSysObject; //адаптер к БД в 
        private OleDbDataAdapter _daObject; //адаптер к БД в 
        private OleDbDataAdapter _daMessage; //адаптер к БД в 
        private OleDbDataAdapter _daTimer; //адаптер к БД в 
        private OleDbDataAdapter _daPlaceTimer; //адаптер к БД в 
        readonly string _pathimportDirectory = Application.StartupPath + "\\Import";
        private readonly bool _useAdrBegDenominator;

        Dictionary<string, Dictionary<string, string>> _usingTablesStructure; 
        #endregion

        #region Инициализация

        public MainForm()
        {
            InitializeComponent();
            Text = "Редактор базы данных " + AssemblyInfo.AssemblyVersion;
            toolStripStatusLabel1.Text = @"Готов";
            ReadOipNumSetting();
            ReadTimerSetting();
            _useAdrBegDenominator = Program.Settings.UseAdrBeginning;

            Program.Settings.Save();
            button_Check.Enabled = false;
            button_ExportToCSV.Enabled = false;
            button_WriteToDB.Enabled = false;
            checkBox_SIKN.Enabled = checkBox_Object.Checked; 
            checkBox_MNS_PNS.Enabled = checkBox_Object.Checked;
            checkBox_RP.Enabled = checkBox_Object.Checked;
            checkBox_PT.Enabled = checkBox_Object.Checked;
            checkBox_SAR.Enabled = checkBox_Object.Checked;
            label_MPSA.Enabled = checkBox_Object.Checked;

       

         var files = FindFiles("*.xls");
            if (files != null)
            {
                foreach (string str in files)
                {
                    if (str.Contains("~")) continue;
                    if (str.Contains("_MSG_")) textBox_path_MSG.Text = str;
                    if (str.Contains("_IO_"))
                    {
                        textBox_path_IO.Text = str;
                        if (str.Contains("СИКН") || str.Contains("SIKN")) checkBox_SIKN.Checked = true;
                        if (str.Contains("ПНС") || str.Contains("PNS")) checkBox_MNS_PNS.Checked = true;
                        if (str.Contains("РП") || str.Contains("RP")) checkBox_RP.Checked = true;
                        if (str.Contains("ПТ") || str.Contains("PT")) checkBox_PT.Checked = true;
                        if (str.Contains("САР") || str.Contains("SAR")) checkBox_SAR.Checked = true;
                    }
                }
            }
            files = FindFiles("*.sql");
            if (files != null)
            {
                List<string> list = new List<string>();
                foreach (string filePath in files)
                {
                    if (!File.Exists(filePath)) return;
                    var reader = new StreamReader(File.OpenRead(filePath), Encoding.Default);
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (line != null && line.Contains("USE")) continue;
                        if (line != null && line.Contains("use")) continue;
                        if (String.Equals(line, "go", StringComparison.OrdinalIgnoreCase)) continue;
                        list.Add(line);
                    }
                    reader.Close();
                }
                richTextBox_SQL_commands.Lines = list.ToArray();
            }
        }

        private string[] FindFiles(string extensionPattern) //ищет инф файлы в стартовой директории
        {
            try
            {
                var path = Application.StartupPath;
                string[] files = Directory.GetFiles(path, extensionPattern, SearchOption.TopDirectoryOnly);
                return files;
            }
            catch (UnauthorizedAccessException uex)
            {
                toolStripStatusLabel1.Text = uex.Source + @": " + uex.Message;
                return null;
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Source + @": " + ex.Message;
                return null;
            }
        }

        private void DataGridViewInicialization()
        {
            dataGridView_OIP.AutoGenerateColumns = true;
            dataGridView_OIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView_OIP.AllowUserToResizeColumns = true;
            dataGridView_OIP.DataSource = Data.DataSet.Tables["OIP"];
            if (dataGridView_OIP.Columns.Count > 1) dataGridView_OIP.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            if (Program.Settings.DB_version > 8)
            {
                dataGridView_OIPType.AutoGenerateColumns = true;
                dataGridView_OIPType.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView_OIPType.AllowUserToResizeColumns = true;
                dataGridView_OIPType.DataSource = Data.DataSet.Tables["OIPTYPE"];
            }

            dataGridView_PlaceOIP.AutoGenerateColumns = true;
            dataGridView_PlaceOIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView_PlaceOIP.DataSource = Data.DataSet.Tables["PlaceOIP"];

            dataGridView_SysObject.AutoGenerateColumns = true;
            dataGridView_SysObject.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView_SysObject.DataSource = Data.DataSet.Tables["SysObject"];

            dataGridView_Message.AutoGenerateColumns = true;
            dataGridView_Message.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView_Message.DataSource = Data.DataSet.Tables["Message"];
            if (dataGridView_Message.Columns.Count > 3) dataGridView_Message.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            dataGridView_Object.AutoGenerateColumns = true;
            dataGridView_Object.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView_Object.DataSource = Data.DataSet.Tables["Object"];

            dataGridView_Timer.AutoGenerateColumns = true;
            dataGridView_Timer.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView_Timer.DataSource = Data.DataSet.Tables["Timer"];


            dataGridView_PlaceTimer.AutoGenerateColumns = true;
            dataGridView_PlaceTimer.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView_PlaceTimer.DataSource = Data.DataSet.Tables["PlaceTimer"];
        }

        private void DataGridViewDataSourceClear()
        {
            dataGridView_OIP.DataSource = null;
            if (Program.Settings.DB_version > 8) dataGridView_OIPType.DataSource = null;
            dataGridView_PlaceOIP.DataSource = null;
            dataGridView_SysObject.DataSource = null;
            dataGridView_Message.DataSource = null;
            dataGridView_Object.DataSource = null;
            dataGridView_Timer.DataSource = null;
            dataGridView_PlaceTimer.DataSource = null;
        }

        private void DataGridViewDataSourceSet()
        {
            dataGridView_OIP.DataSource = Data.DataSet.Tables["OIP"];
            if (Program.Settings.DB_version > 8) dataGridView_OIPType.DataSource = Data.DataSet.Tables["OIPType"];
            dataGridView_PlaceOIP.DataSource = Data.DataSet.Tables["PlaceOIP"];
            dataGridView_SysObject.DataSource = Data.DataSet.Tables["SysObject"];
            dataGridView_Message.DataSource = Data.DataSet.Tables["Message"];
            dataGridView_Object.DataSource = Data.DataSet.Tables["Object"];
            dataGridView_Timer.DataSource = Data.DataSet.Tables["Timer"];
            dataGridView_PlaceTimer.DataSource = Data.DataSet.Tables["PlaceTimer"];
        }


        #endregion

        #region Кнопки 
        private void button_Preview_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Stopwatch watch = new Stopwatch();
                watch.Start();
                if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                    !checkBox_Object.Checked && !checkBox_Timer.Checked)
                {
                    toolStripStatusLabel1.Text = @"Ничего не выбрано";
                    return;
                }
                button_Check.Enabled = true;
                button_ExportToCSV.Enabled = true;
                button_WriteToDB.Enabled = true;

                button_Check.BackColor = DefaultBackColor;
                button_Check.UseVisualStyleBackColor = true;
                if (!File.Exists(textBox_path_IO.Text)) throw new Exception("Файл (информационный) не найден");
                if (!File.Exists(textBox_path_MSG.Text)) throw new Exception("Файл (сообщений) не найден");
                _fileIo = new Excel(textBox_path_IO.Text, false, false);
                _fileMsg = new Excel(textBox_path_MSG.Text, false, true);
                if (!checkBox_bdEditor.Checked)
                {
                    Data.IninicializationDataSet();
                    toolStripStatusLabel1.Text = @"Генерация из Excel";
                    Refresh();
                }
                else
                {
                    toolStripStatusLabel1.Text = @" --> генерация из Excel";
                }

                if (checkBox_OIP.Checked) ReadOip();
                if (checkBox_SysObject.Checked) ReadSysObject();
                if (checkBox_Message.Checked) ReadMessage();
                if (checkBox_Object.Checked) ReadObject();
                if (checkBox_Timer.Checked) ReadTimer();
                DataGridViewInicialization();
                watch.Stop();
                toolStripStatusLabel1.Text += @" [прошло " + watch.ElapsedMilliseconds + @"мс]";
                Cursor = Cursors.Default;
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = exception.Source + @": " + exception.Message;
            }
            finally
            {
                Cursor = Cursors.Default;
                _fileIo?.Close();
                _fileMsg?.Close();
            }
        }

        private void button_Check_Click(object sender, EventArgs e)
        {
            if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                !checkBox_Object.Checked && !checkBox_Timer.Checked)
            {
                toolStripStatusLabel1.Text = @"Ничего не выбрано";
                return;
            }
            //checkBox_OIP.BackColor = Color.FromName("Control");
            if (checkBox_OIP.Checked) checkBox_OIP.BackColor = CheckOip() ? Color.IndianRed : Color.ForestGreen;
            if (checkBox_SysObject.Checked)
                checkBox_SysObject.BackColor = CheckSysObject() ? Color.IndianRed : Color.ForestGreen;
            if (checkBox_Message.Checked)
                checkBox_Message.BackColor = CheckMessage() ? Color.IndianRed : Color.ForestGreen;
            if (checkBox_Object.Checked) checkBox_Object.BackColor = CheckObject() ? Color.IndianRed : Color.ForestGreen;
            if (checkBox_Timer.Checked) checkBox_Timer.BackColor = CheckTimer() ? Color.IndianRed : Color.ForestGreen;
            toolStripStatusLabel1.Text = @"Проверка выполнена";
        }

        private void button_WriteDB_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Stopwatch watch = new Stopwatch();
                watch.Start();
                if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                    !checkBox_Object.Checked && !checkBox_Timer.Checked)
                {
                    toolStripStatusLabel1.Text = @"Ничего не выбрано";
                    return;
                }

                //button_ClearBD_Click(null,null); //чистим базу
                DataGridViewDataSourceClear();

                _oleDb = new OleDb();
                _oleDb.Open(); //соединение с базой
                //проверяем структуру базы
                _usingTablesStructure = OleDb.GetUsingTablesStructure();
                foreach (KeyValuePair<string, Dictionary<string, string>> table in _usingTablesStructure)
                {
                    if (_oleDb.CheckTable(table.Key, table.Value) == false)
                        throw new Exception("Ошибка структуры базы данных в таблице " + table.Key);
                }
                _oleDb.Close();

                toolStripStatusLabel1.Text = @"Запись в БД";
                button_Check.BackColor = DefaultBackColor;
                button_Check.UseVisualStyleBackColor = true;

                if (checkBox_OIP.Checked) WriteOip();
                if (checkBox_SysObject.Checked) WriteSysObject();
                if (checkBox_Object.Checked) WriteObject();
                if (checkBox_Message.Checked) WriteMessage();
                if (checkBox_Timer.Checked) WriteTimer();

                watch.Stop();
                toolStripStatusLabel1.Text += @" [прошло " + watch.ElapsedMilliseconds + @"мс]";
                Cursor = Cursors.Default;
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка записи в БД [" + exception.Source + @"]: " + exception.Message;
                if (_oleDb != null) toolStripStatusLabel1.Text += @" SQL: " + _oleDb.MessErr;
            }
            finally
            {
                DataGridViewDataSourceSet();
                Cursor = Cursors.Default;
                if (_oleDb != null && _oleDb.SqlConnected()) _oleDb.Close();
            }
        }

        private void button_ClearBD_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Stopwatch watch = new Stopwatch();
                watch.Start();
                if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                    !checkBox_Object.Checked && !checkBox_Timer.Checked)
                {
                    toolStripStatusLabel1.Text = @"Ничего не выбрано";
                    return;
                }

                toolStripStatusLabel1.Text = @"Очистка БД";

                _oleDb = new OleDb();
                _oleDb.Open(); //соединение с базой
                if (checkBox_OIP.Checked)
                {
                    _oleDb.TruncateTable("OIP");
                    _oleDb.TruncateTable("PlaceOIP");
                    if (Program.Settings.DB_version > 8) _oleDb.TruncateTable("OIPType");
                }
                if (checkBox_Message.Checked) _oleDb.TruncateTable("Message");
                if (checkBox_Object.Checked) _oleDb.TruncateTable("Object");
                if (checkBox_Timer.Checked)
                {
                    _oleDb.TruncateTable("Timer");
                    _oleDb.TruncateTable("PlaceTimer");
                }
                if (checkBox_SysObject.Checked) _oleDb.TruncateTable("SysObject");
                _oleDb.Close();

                watch.Stop();
                toolStripStatusLabel1.Text += @" [прошло " + watch.ElapsedMilliseconds + @"мс]";
                Cursor = Cursors.Default;
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка очистки БД [" + exception.Source + @"]: " + exception.Message;
                if (_oleDb != null) toolStripStatusLabel1.Text += @" SQL: " + _oleDb.MessErr;
            }
            finally
            {
                Cursor = Cursors.Default;
                if (_oleDb != null && _oleDb.SqlConnected()) _oleDb.Close();
            }
        }

        private void button_Explorer_Click(object sender, EventArgs e)
        {
            Process.Start(Application.StartupPath);
        }

        #endregion

        #region Экспорт

        private void button_Export_Click(object sender, EventArgs e)
        {
            try
            {
                if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                    !checkBox_Object.Checked && !checkBox_Timer.Checked)
                {
                    toolStripStatusLabel1.Text = @"Ничего не выбрано";
                    return;
                }
                string ioFilePath = Path.GetFileNameWithoutExtension(textBox_path_IO.Text);
                string exportPath = Application.StartupPath + "\\Export_" + ioFilePath;
                if (checkBox_bdEditor.Checked && _oleDb != null) exportPath = Application.StartupPath + "\\ExportDB_" + _oleDb.InitialCatalog;
                if (!Directory.Exists(exportPath)) Directory.CreateDirectory(exportPath);
                if (!Directory.Exists(_pathimportDirectory)) Directory.CreateDirectory(_pathimportDirectory);
                toolStripStatusLabel1.Text = @"Экспорт -->";
                if (checkBox_OIP.Checked)
                {
                    ExportToCsvFile(Data.DataSet.Tables["OIP"], exportPath + "\\OIP.csv");
                    if (Program.Settings.DB_version > 8) ExportToCsvFile(Data.DataSet.Tables["OIPType"], exportPath + "\\OIPType.csv");
                    ExportToCsvFile(Data.DataSet.Tables["PlaceOIP"], exportPath + "\\PlaceOIP.csv");
                    CopyFile(exportPath + "\\OIP.csv", _pathimportDirectory + "\\OIP.csv");
                    if (Program.Settings.DB_version > 8) CopyFile(exportPath + "\\OIPType.csv", _pathimportDirectory + "\\OIPType.csv");
                    CopyFile(exportPath + "\\PlaceOIP.csv", _pathimportDirectory + "\\PlaceOIP.csv");
                    if (Program.Settings.DB_version > 8) toolStripStatusLabel1.Text += @"OIP PlaceOIP OIPType";
                    else
                    {
                        toolStripStatusLabel1.Text += @"OIP PlaceOIP";
                    }
                }
                if (checkBox_SysObject.Checked)
                {
                    ExportToCsvFile(Data.DataSet.Tables["SysObject"], exportPath + "\\SysObject.csv"); toolStripStatusLabel1.Text += @" SysObject ";
                    CopyFile(exportPath + "\\SysObject.csv", _pathimportDirectory + "\\SysObject.csv");
                }
                if (checkBox_Message.Checked)
                {
                    ExportToCsvFile(Data.DataSet.Tables["Message"], exportPath + "\\Message.csv"); toolStripStatusLabel1.Text += @" Message ";
                    CopyFile(exportPath + "\\Message.csv", _pathimportDirectory + "\\Message.csv");
                }
                if (checkBox_Object.Checked)
                {
                    ExportToCsvFile(Data.DataSet.Tables["Object"], exportPath + "\\Object.csv"); toolStripStatusLabel1.Text += @" Object ";
                    CopyFile(exportPath + "\\Object.csv", _pathimportDirectory + "\\Object.csv");
                }
                if (checkBox_Timer.Checked)
                {
                    ExportToCsvFile(Data.DataSet.Tables["Timer"], exportPath + "\\Timer.csv");
                    CopyFile(exportPath + "\\Timer.csv", _pathimportDirectory + "\\Timer.csv");
                    ExportToCsvFile(Data.DataSet.Tables["PlaceTimer"], exportPath + "\\PlaceTimer.csv");
                    CopyFile(exportPath + "\\PlaceTimer.csv", _pathimportDirectory + "\\PlaceTimer.csv");
                    toolStripStatusLabel1.Text += @" Timer PlaceTimer";
                }
                toolStripStatusLabel1.Text += @" выполнен в " + exportPath;
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка экспорта [" + exception.Source + @"]: " + exception.Message;
            }
        }

        private void CopyFile(string sourcefn, string destinfn)
        {
            try
            {
                FileInfo fn = new FileInfo(sourcefn);
                fn.CopyTo(destinfn, true);
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка копирования файлов из" + sourcefn + @" в " + destinfn + @" !! " + exception.Message;
            }
        }

        private void ExportToCsvFile(DataTable dtTable, string exportPath)
        {
            if (dtTable == null) return;
            DataTable tableExport = dtTable.Copy();
            StringBuilder sbldr = new StringBuilder();
            if (tableExport.Columns.Count != 0)
            {
                foreach (DataColumn col in tableExport.Columns) //пишем заголовок
                {
                    if (col.Ordinal == tableExport.Columns.Count - 1)
                    {
                        sbldr.Append(col.ColumnName); //в последнюю строку
                    }
                    else
                    {
                        sbldr.Append(col.ColumnName + ';');
                    }
                }
                sbldr.Append("\r\n");
                foreach (DataRow row in tableExport.Rows)
                {
                    foreach (DataColumn column in tableExport.Columns)
                    {
                        if (column.DataType == Type.GetType("System.String") && column.ColumnName.ToUpper() != "UNIT")
                        {
                            if (column.Ordinal > 0) row[column] = "\"" + row[column].ToString().ToUpper() + "\"";
                            if (row[column].ToString() == "\"\"") row[column] = "\"NaN\"";
                        }
                        if (column.ColumnName.ToUpper() == "UNIT")
                        {
                            if (row[column].ToString().ToUpper() == "МПА") row[column] = "МПа";
                        }
                        if (column.Ordinal == tableExport.Columns.Count - 1)
                        {
                            sbldr.Append(row[column]); //в последнюю строку
                        }
                        else
                        {
                            sbldr.Append(row[column].ToString() + ';');
                        }
                    }
                    sbldr.Append("\r\n");
                }
            }
            StreamWriter sw = new StreamWriter(new FileStream(exportPath, FileMode.Create, FileAccess.Write), Encoding.Default);
            sw.Write(sbldr);
            sw.Close();
        }

        #endregion

        #region OIP
        /// <summary>
        /// чтение OIP
        /// </summary>
        private void ReadOip()
        {

            var sysName = "OIP";
            var sheetOip = _fileIo.GetSheet("ИП");
            var startRow = OipStartRow;
            var endRow = _fileIo.LastRowCell(ref sheetOip);
            int num = endRow - startRow + 1;

            Dictionary<string, int> placeOipDictionary = new Dictionary<string, int>(); //словарик для PlaceOip
            Dictionary<string, int> OipTypeDictionary = new Dictionary<string, int>(); //словарик для OipType
            object[,] range = {}; // целевая таблица Для OIP

            int OIPrownumber = 23;

            switch (Program.Settings.DB_version )
            {
                case 4:
                    OIPrownumber = 23;
                    break;
                case 5:
                    OIPrownumber = 23;
                    break;
                case  6:
                    OIPrownumber = 23;
                    break;
                case 7:
                    OIPrownumber = 24;
                    break;
                case 8:
                    OIPrownumber = 30;
                    break;
                case 9:
                    OIPrownumber = 33;
                    break;
            }
            range = new object[num, OIPrownumber]; 
            
            List<object[,]> blocks = new List<object[,]>(); // лист из блоков, каждый блок - двумерный массив // спасибо, так гораздо понятнее (нет)
            //заполняем столбец адресов
            for (int i = 0; i < range.GetLength(0); i++)
                range[i, 21] = OipAdressBeginning + 34*i;

            //заполняем лист:
            //ID формируется автоматически
            blocks.Add(
                sheetOip.Range[
                    sheetOip.Cells[startRow, OipParamName],
                    sheetOip.Cells[endRow, OipParamId]].Value); //0 - имя и идентификатор параметра
            blocks.Add(
                sheetOip.Range[
                    sheetOip.Cells[startRow, OipUnit],
                    sheetOip.Cells[endRow, OipUnit]].Value); //1 - ед изм
            blocks.Add(
                sheetOip.Range[
                    sheetOip.Cells[startRow, OipScaleBeginning],
                    sheetOip.Cells[endRow, OipScaleBeginning]].Value); //2 - нижний предел изм
            blocks.Add(
                sheetOip.Range[
                    sheetOip.Cells[startRow, OipUst0],
                    sheetOip.Cells[endRow, OipUst11]].Value); //3 - уставки 0-11
            blocks.Add(
                sheetOip.Range[
                    sheetOip.Cells[startRow, OipScaleEnd],
                    sheetOip.Cells[endRow, OipLimSpd]].Value); //4 - верхний предел изм, гистерезеис, Зона нечувствит-сти, Предел скорости изменения
            //Адрес уже записали ранее

            if (Program.Settings.DB_version < 7) //раньше был видеокадр вместо принадлжености
            {
                blocks.Add(
                    sheetOip.Range[
                        sheetOip.Cells[startRow, OipPlaceOld],
                        sheetOip.Cells[endRow, OipPlaceOld]].Value); //5 - место
            }
            else //принадлежность
            {
                blocks.Add(
                    sheetOip.Range[
                        sheetOip.Cells[startRow, OipPlace],
                        sheetOip.Cells[endRow, OipPlace]].Value); //5 - место
            }



            //просмотрим blocks[5] и добавим уникальные PlaceOip в словарь + заменим исходный
            int j = 1;
            for (int i = 0; i < blocks[5].GetLength(0); i++)
            {
                if (blocks[5][i + 1, 1] == null) continue; //нули пропускаем
                if (!placeOipDictionary.ContainsKey(blocks[5][i + 1, 1].ToString()))
                {
                    //если такого ключа в словаре нет - добавляем
                    placeOipDictionary.Add(blocks[5][i + 1, 1].ToString(), j);
                    blocks[5][i + 1, 1] = j;
                    j++;
                }
                else
                {
                    //если есть заменяем исходный на значение ключа
                    blocks[5][i + 1, 1] = placeOipDictionary[blocks[5][i + 1, 1].ToString()];
                }
            }

            if (Program.Settings.DB_version >= 9)
                //c 9-й версии добавился столбец OIPType, добавляем его следующим блоком.
            {
                blocks.Add(
                    sheetOip.Range[
                        sheetOip.Cells[startRow, OipType],
                        sheetOip.Cells[endRow, OipType]].Value); //6 - место тип аналога


                // blocks[blocks.Count] = массив OIPType
                int jj = 1;
                for (int i = 0; i < blocks[6].GetLength(0); i++)
                {
                    if (blocks[6][i + 1, 1] == null) continue; //нули пропускаем
                    if (!OipTypeDictionary.ContainsKey(blocks[6][i + 1, 1].ToString().Split('-')[1].Trim()))
                    {
                        //если такого ключа в словаре нет - добавляем
                        OipTypeDictionary.Add(blocks[6][i + 1, 1].ToString().Split('-')[1].Trim(),
                            Convert.ToInt32(blocks[6][i + 1, 1].ToString().Split('-')[0].Trim()));
                        blocks[6][i + 1, 1] = jj;
                        jj++;
                    }
                    else
                    {
                        //если есть заменяем исходный на значение ключа
                        blocks[6][i + 1, 1] = OipTypeDictionary[blocks[6][i + 1, 1].ToString().Split('-')[1].Trim()];
                    }
                }
            }

            //готовим целевой массив
            for (int i = 0; i < range.GetLength(0); i++) //проходим по строкам
            {
                int sdvig = 0;
                foreach (object[,] block in blocks) // проходим по всем блокам
                {
                    // todo с этим сдвигом могут появиться баги. Теперь столбец не последний.
                    if (blocks.IndexOf(block) == 5) sdvig = sdvig + 1; //учитываем для <последнего> 5-го столбца в сдвиге то, что адрес мы заполнили сами
                    for (int k = 0; k < block.GetLength(1); k++) //столбцы 
                    {
                        range[i, k + 1 + sdvig] = block[i + 1, k + 1]; //проходим по элементам
                    }
                    sdvig = sdvig + block.GetLength(1);
                }
            }

            //заполняем datatable OIP
            object[] rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
            for (int i = 0; i < range.GetLength(0); i++)
            {
                for (int k = 1; k < range.GetLength(1); k++)
                {
                    rowObjects[k] = range[i, k];
                }
                if (rowObjects[1] == null) continue;
                if (Program.Settings.Gen_OIP_NoGap)
                { 
                    if (rowObjects[4] == null) rowObjects[4] = 0; //ДОзаполняем ScaleBegin
                    if (rowObjects[17] == null) rowObjects[17] = 100; //ДОзаполняем ScaleEnd
                    if (rowObjects[18] == null) rowObjects[18] = 0; //ДОзаполняем Hist
                    if (rowObjects[19] == null) rowObjects[19] = 0; //ДОзаполняем Delta
                    if (rowObjects[20] == null) rowObjects[20] = 0; //ДОзаполняем LimSpd
                    if (rowObjects[22] == null) rowObjects[22] = Program.Settings.Gen_OIP_NoGapPlaceValue; //ДОзаполняем Place
                }
                if (Program.Settings.DB_version > 6)
                {
                    rowObjects[23] = 1; //ДОзаполняем PLC
                }
                if (Program.Settings.DB_version > 7)
                {
                    rowObjects[24] = 1; //ДОзаполняем KF
                    rowObjects[25] = 1; //ДОзаполняем NDiap
                    rowObjects[26] = Convert.ToInt32(OipAdressBeginningKf + i*4); //ДОзаполняем  AdressBeginningKF
                    rowObjects[27] = 1; //ДОзаполняем CtrlUst
                    rowObjects[28] = 1; //ДОзаполняем OpMask
                    rowObjects[29] = 1; //ДОзаполняем SignMask
                }
                if (Program.Settings.DB_version > 8)
                {
                    rowObjects[30] = ""; //OPCPATH
                    rowObjects[31] = placeOipDictionary.FirstOrDefault(x => x.Value == (int) rowObjects[22]).Key; //PLACENAME
                    rowObjects[32] = ParseTypeOip(sheetOip.Cells[i + startRow, OipType].Value.ToString().Split('-')[1].Trim()); //OipType
                }
                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); //если название не пустое добавляем строку
            }

            if (Program.Settings.DB_version >= 9) //заполнение OPCPATH
            {
                //формирование иерархии ИП
                object[,] rangeIe = sheetOip.Range[sheetOip.Cells[startRow, 46], sheetOip.Cells[endRow, 46]].Value;
                //46 - номер столбца полной иерархии
                string stationTag = Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + "."; //получим иерархию проекта
                for (int i = 0; i < Data.DataSet.Tables[sysName].Rows.Count; i++)
                {
                    Data.DataSet.Tables[sysName].Rows[i][30] = stationTag + rangeIe[i+1,1]; 
                }
            }

            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                             Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
            //заполняем datatable PlaceOIP
            sysName = "PlaceOIP";
            foreach (KeyValuePair<string,int> placeObj in placeOipDictionary)
            {
                Data.DataSet.Tables[sysName].Rows.Add(null, placeObj.Value, placeObj.Key);
            }
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                                         Data.DataSet.Tables[sysName].Rows.Count + @" параметров";


            if (Program.Settings.DB_version > 8)
            {
                //заполняем datatable OIPType
                sysName = "OipType";
                foreach (KeyValuePair<string, int> OIPunit in OipTypeDictionary)
                {
                    Data.DataSet.Tables[sysName].Rows.Add(null, OIPunit.Value, OIPunit.Key);
                        // Длина входного массива больше числа столбцов в этой таблице.
                }
                toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                                             Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
            }



        }
        /// <summary>
        /// проверка OIP
        /// </summary>
        private bool CheckOip() 
        {
            ResetColorTable(dataGridView_OIP);
            ResetColorTable(dataGridView_PlaceOIP);
            bool res = false;
            foreach (DataGridViewRow row in dataGridView_OIP.Rows)
            {
                bool resIn = CheckColumnInTable(row, 1) || CheckColumnInTable(row, 2) || CheckColumnInTable(row, 3) || CheckColumnInTable(row, 4) ||
                    CheckColumnInTable(row, 17) || CheckColumnInTable(row, 18) || CheckColumnInTable(row, 19) || CheckColumnInTable(row, 20) || CheckColumnInTable(row, 21) || CheckColumnInTable(row, 22);
                if (resIn) res = true;
            }
            if (dataGridView_OIP.Rows.Count == 0) res = true;
            res = res || CheckTable(dataGridView_PlaceOIP);
            return res;
        }

        private void WriteOip()
        {
            string sysName = "PlaceOIP";
            Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
            {
                foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                {
                    if (row[column].ToString() == "") row[column] = "NaN";
                }
                row[2] = row[2].ToString().ToUpper();

                _oleDb.WriteOperMess_PlaceOIP((uint)row.ItemArray[1], row.ItemArray[2].ToString());
            }
            Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";

            sysName = "OIP";
                Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
                {
                    foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                    {
                        if (row[column].ToString() == "") row[column] = "NaN";
                    }
                    row[1] = row[1].ToString().ToUpper();

                    if (Program.Settings.DB_version < 7)
                    {
                        _oleDb.WriteOperMess_OIP(row.ItemArray[1].ToString(), row.ItemArray[2].ToString(),
                            row.ItemArray[3].ToString(),
                            row.ItemArray[4].ToString(), row.ItemArray[5].ToString(), row.ItemArray[6].ToString(),
                            row.ItemArray[7].ToString(),
                            row.ItemArray[8].ToString(), row.ItemArray[9].ToString(), row.ItemArray[10].ToString(),
                            row.ItemArray[11].ToString(),
                            row.ItemArray[12].ToString(), row.ItemArray[13].ToString(), row.ItemArray[14].ToString(),
                            row.ItemArray[15].ToString(),
                            row.ItemArray[16].ToString(), row.ItemArray[17].ToString(), row.ItemArray[18].ToString(),
                            row.ItemArray[19].ToString(),
                            row.ItemArray[20].ToString(), Convert.ToUInt32(row.ItemArray[21]), Convert.ToUInt32(row.ItemArray[22]));
                    }
                    else
                    {
                        _oleDb.WriteOperMess_OIP(row.ItemArray[1].ToString(), row.ItemArray[2].ToString(),
                            row.ItemArray[3].ToString(),
                            row.ItemArray[4].ToString(), row.ItemArray[5].ToString(), row.ItemArray[6].ToString(),
                            row.ItemArray[7].ToString(),
                            row.ItemArray[8].ToString(), row.ItemArray[9].ToString(), row.ItemArray[10].ToString(),
                            row.ItemArray[11].ToString(),
                            row.ItemArray[12].ToString(), row.ItemArray[13].ToString(), row.ItemArray[14].ToString(),
                            row.ItemArray[15].ToString(),
                            row.ItemArray[16].ToString(), row.ItemArray[17].ToString(), row.ItemArray[18].ToString(),
                            row.ItemArray[19].ToString(),
                            row.ItemArray[20].ToString(), (uint)row.ItemArray[21], (uint)row.ItemArray[22],
                            (uint)row.ItemArray[23]);
                    }
                }
                Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";
        }

        public string ParseTypeOip(string arg)
        {
           string[] typeOip = arg.Split(Convert.ToChar("-"));
            return arg; //todo typeOip[1].Trim();

        }

        private static bool CheckColumnInTable(DataGridViewRow row, int columnNumber)
        {
            bool res = false;
            if (row.Cells[columnNumber].Value.ToString() == "")
            {
                row.Cells[columnNumber].Style.BackColor = Color.Red;
                res = true;
            }
            return res;
        }

        private bool CheckTable(DataGridView dataGridView)
        {
            bool res = false;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (row.Cells[i].Value.ToString() == "")
                    {
                        row.Cells[i].Style.BackColor = Color.Red;
                        res = true;
                    }
                }
            }
            if (dataGridView.Rows.Count == 0) res = true;
            return res;
        }

        private void ResetColorTable(DataGridView dataGridView)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++) row.Cells[i].Style.BackColor = Color.White;
            }
        }

        #endregion
        
        #region SysObject

        private void ReadSysObject() //чтение 
        {
            var sysName = "SysObject";
            var sheetSysId = _fileMsg.GetSheet("SysID");
            var startRow = 4;
            var endRow = _fileMsg.LastRowCell(ref sheetSysId);

            //заполняем массив:
            dynamic range = sheetSysId.Range[sheetSysId.Cells[startRow, 1], sheetSysId.Cells[endRow, 4]].Value;

            //заполняем datatable
            object[] rowObjects = new object[range.GetLength(1)]; //массив представляющий строку
            for (int i = 0; i < range.GetLength(0); i++)
            {
                for (int k = 0; k < range.GetLength(1) - 1; k++)
                {
                    rowObjects[k + 1] = range[i + 1, k + 1];
                }
                if (rowObjects[2] == null) continue;
                if (rowObjects[2].ToString() == "") continue;
                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); //если название не пустое добавляем строку
            }
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                                         Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
        }

        private bool CheckSysObject() //проверка
        {
            ResetColorTable(dataGridView_SysObject);
            return CheckTable(dataGridView_SysObject);
        }

        private void WriteSysObject()
        {
            string sysName = "SysObject";
            Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
            {
                foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                {
                    if (row[column].ToString() == "") row[column] = "NaN";
                }
                row[2] = row[2].ToString().ToUpper();
                row[3] = row[3].ToString().ToUpper();

                _oleDb.WriteOperMess_SysObject((uint) row.ItemArray[1], row.ItemArray[2].ToString(),
                        row.ItemArray[3].ToString());
            }
            Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";
        }

        #endregion SysObject

        #region Message

        private void ReadMessage() //чтение 
        {
            var sysName = "Message";
            //читаем SysObject
            foreach (DataRow row in Data.DataSet.Tables["SysObject"].Rows)
            {
                if (row[2] == null) continue;
                var sheetSys = _fileMsg.GetSheet(row[2].ToString());
                if (sheetSys == null) continue;
                var startRow = 2;
                var endRow = _fileMsg.LastRowCell(ref sheetSys);
                //заполняем массив:
                dynamic range = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 6]].Value;
                //заполняем datatable
                object[] rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                for (int i = 0; i < range.GetLength(0) - 1; i++)
                {
                    rowObjects[3] = range[i + 1, 3]; //сначала проверазвание
                    if (rowObjects[3] == null) continue; //выход если ноль
                    if (rowObjects[3].ToString().Trim() == "") continue; //выход если ноль
                    rowObjects[1] = row[1];
                    rowObjects[2] = range[i + 1, 1];
                    rowObjects[4] = DetectKind(range[i + 1, 4]);
                    rowObjects[5] = DetectPriority(range[i + 1, 5]);
                    rowObjects[6] = DetectSound(range[i + 1, 6]);
                    rowObjects[7] = DetectIdSound(range[i + 1, 6]);
                    rowObjects[8] = 1; //что то тут хотели сделать
                    rowObjects[9] = DetectIsAck(range[i + 1, 5]);
                    rowObjects[10] = DetectIdColor(range[i + 1, 5]);
                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                }
            }
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                                         Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
        }

        private bool CheckMessage() //проверка
        {
            ResetColorTable(dataGridView_Message);
            bool res = false;
            foreach (DataGridViewRow row in dataGridView_Message.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (i > 3)
                    {
                        if (row.Cells[i].Value.ToString() == "" || row.Cells[i].Value.ToString() == "-1")
                        {
                            row.Cells[i].Style.BackColor = Color.Red;
                            res = true;
                        }
                    }
                    else
                    {
                        if (row.Cells[i].Value.ToString() == "")
                        {
                            row.Cells[i].Style.BackColor = Color.Red;
                            res = true;
                        }
                    }
                }
            }
            if (dataGridView_Message.Rows.Count == 0) res = true;
            return res;
        }

        private void WriteMessage()
        {
            string sysName = "Message";
            Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
            {
                foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                {
                    if (row[column].ToString() == "") row[column] = "NaN";
                }
                row[3] = row[3].ToString().ToUpper();

                    _oleDb.WriteOperMess_Message((uint) row.ItemArray[1], (uint)row.ItemArray[2],
                        row.ItemArray[3].ToString(), (uint)row.ItemArray[4], (uint)row.ItemArray[5], (uint)row.ItemArray[6],
                       (uint)row.ItemArray[7], (uint) row.ItemArray[8], (uint) row.ItemArray[9], (uint)row.ItemArray[10]);

            }
            Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";
        }

        private int DetectKind(string strKind)
        {
            switch (strKind)
            {
                case "Нет":
                    return 0;
                case "C":
                    return 1;
                case "С":
                    return 1;
            }
            return -1;
        }

        private int DetectPriority(string strPriority)
        {
            switch (strPriority)
            {
                case "Нормальный":
                case "Норм":
                    return 1;
                case "Низкий":
                case "Синий":
                case "Низ":
                    return 2;
                case "Средний":
                case "Сред":
                    return 3;
                case "Высокий":
                case "Высш":
                    return 4;
            }
            return -1;
        }

        private int DetectSound(string strSound)
        {
            switch (strSound)
            {
                case "Нет":
                    return 0;
                case "Однокр":
                    return 1;
                case "Многокр":
                    return 2;
            }
            return -1;
        }

        private int DetectIdSound(string strSound)
        {
            switch (strSound)
            {
                case "Нет":
                    return 1;
                case "Однокр":
                    return 1;
                case "Многокр":
                    return 2;
            }
            return -1;
        }

        private int DetectIsAck(string strPriority)
        {
            switch (strPriority)
            {
                case "Нормальный":
                case "Норм":
                    return 0;
                case "Низкий":
                case "Синий":
                case "Низ":
                    return 0;
                case "Средний":
                case "Сред":
                    return 1;
                case "Высокий":
                case "Высш":
                    return 1;
            }
            return -1;
        }

        private int DetectIdColor(string strPriority)
        {
            switch (strPriority)
            {
                case "Нормальный":
                case "Норм":
                    return 4; //белый
                case "Низкий":              
                case "Низ":
                    return 1; //зеленый
                case "Средний":
                case "Сред":
                    return 2; //желтый
                case "Высокий":
                case "Высш":
                    return 3; //красный
                case "Синий":
                    return 5; //синий
            }
            return -1;
        }

        #endregion Message

        #region Object

        private void ReadObject() //чтение 
        {
            var sysName = "Object";
            //читаем лист SysID блоком и заполняем таблицу SysID
            // 1-Id / 2-sysId/3-Система/4-Описание/5-комменты/6-Лист в IO/7-НачСтрока/8-Столбец название/9-Столбец идентификатор/10-Столбец с SysNum

            #region SysID

            Data.IninicializationDataSetSysId();
            var sheetSysId = _fileMsg.GetSheet("SysID");
            dynamic sysIdArray = sheetSysId.Range[sheetSysId.Cells[4, 1], sheetSysId.Cells[258, 10]].Value;
            object[] rowObjects = new object[sysIdArray.GetLength(1)]; //массив представляющий строку
            for (int i = 0; i < sysIdArray.GetLength(0); i++)
            {
                for (int k = 0; k < sysIdArray.GetLength(1) - 1; k++) rowObjects[k + 1] = sysIdArray[i + 1, k + 1];
                if (rowObjects[2] == null) continue;
                if (rowObjects[2].ToString() == "") continue;
                Data.DataSetSysId.Tables["SysID"].Rows.Add(rowObjects); //если название не пустое добавляем строку
            }

            #endregion

            //читаем SysID
            dynamic rangeDrop = null;
            dynamic rangeDropSARD = null;
            bool dropFirstRead=true;
            bool dropFirstReadSARD = true;
            foreach (DataRow row in Data.DataSetSysId.Tables["SysID"].Rows)
            {
                if (row[5].ToString() == "") continue;
                var sheetSys = _fileIo.GetSheet(row[5].ToString());
                if (sheetSys == null) continue;
                var startRow = 0;
                if (row[6].ToString() != "") startRow = Convert.ToInt32(row[6]);
                var endRow = _fileIo.LastRowCell(ref sheetSys);
                var colParamName = 0;
                if (row[7].ToString() != "") colParamName = Convert.ToInt32(row[7]);
                var colParamId = 0;
                if (row[8].ToString() != "") colParamId = Convert.ToInt32(row[8]);
                var colParamSysNum = 0;
                if (row[9].ToString() != "") colParamSysNum = Convert.ToInt32(row[9]);
                dynamic rangeParamName;
                dynamic rangeParamSysNum;
                if (row[5].ToString() == "Drop" && dropFirstRead) //читаем один раз дроп для диагностики если есть
                {
                    rangeDrop = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 23]].Value; //весь дроп
                    dropFirstRead = false;
                }
                if (row[5].ToString() == "DropSAR" && dropFirstReadSARD) //читаем один раз дроп для диагностики если есть
                {
                    rangeDropSARD = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 23]].Value; //весь дроп
                    dropFirstReadSARD = false;
                }
                int numPlc;
                switch (row[2].ToString()) //sysId
                {
                    #region Автоматика МНС - ПНС

                    case "KTPR":
                        if (!checkBox_MNS_PNS.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;

                    case "KTPRA":
                        #region KTPRA
                            var blockStart = false; //признак окончания блока
                            var blockName = "";
                            var blockNum = 0;
                            var numStart = 1; //стартовый номер для блока
                        try
                        {
                            //НЕ анализирует столбец - "наличие защиты" в перспективе надо сделать унифицировано с готовностями
                            if (!checkBox_MNS_PNS.Checked) break;
                            startRow = startRow - 1; //чтобы поймать первое название
                            rangeParamName =
                                sheetSys.Range[
                                    sheetSys.Cells[startRow, colParamName - 1], sheetSys.Cells[endRow, colParamName]]
                                    .Value; //читаем весь кусок этой какули

                            for (int i = 1; i < rangeParamName.GetLength(0); i++)
                            {
                                //порядок проверок условий важен!
                                if (rangeParamName[i, 1] == null) continue; //обработка пустых строк
								var str = rangeParamName[i, 1].ToString();
                                int res;
                                bool isInt = Int32.TryParse(str, out res);			  
                                if (!isInt) //поиск начала блока
                                {
                                    blockName = rangeParamName[i, 1].ToString().Trim();
                                    numStart = 1 + 96 * blockNum;
                                    blockNum++;
                                    blockStart = true; //ставим флаг что блок кончился
                                    continue;
                                }
                                if (rangeParamName[i, 2] == null)
                                {
                                    numStart++;
                                    continue;
                                }
                                if (rangeParamName[i, 2].ToString().Trim() != "" && blockStart == false)
                                    //обработка других строк блока с данными
                                {
                                    rangeParamName[i, 1] = numStart;
                                    numStart++;
                                    rangeParamName[i, 2] = "АГРЕГАТНАЯ ЗАЩИТА " + rangeParamName[i, 2].ToString() + " " + blockName;
                                }
                                if (rangeParamName[i, 2].ToString().Trim() != "" && blockStart)
                                    //обработка первой строки блока с данными
                                {
                                    blockStart = false;
                                    rangeParamName[i, 1] = numStart;
                                    numStart++;
                                    rangeParamName[i, 2] = "АГРЕГАТНАЯ ЗАЩИТА " + rangeParamName[i, 2].ToString() + " " + blockName;
                                }
                            }
                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeParamName.GetLength(0) - 1; i++)
                            {
                                if (rangeParamName[i + 1, 2] == null) continue;
                                if (rangeParamName[i + 1, 2].ToString().Trim() == "") continue; //пропускаем пропуски
                                rowObjects[1] = row[1]; //sysId
                                rowObjects[2] = rangeParamName[i + 1, 1];
                                rowObjects[3] = rangeParamName[i + 1, 2];
                                rowObjects[4] = -1;
                                rowObjects[5] = -1;
                                rowObjects[6] = -1;
                                if (Program.Settings.DB_version > 6)
                                {
                                    rowObjects[7] = -1;
                                    rowObjects[8] = -1;
                                    rowObjects[9] = -1;
                                }
                                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в KTPRA";
                        }
                        #endregion KTPRA
                        break;

                    case "KGMPNA":
                        #region KGMPNA 

                        try
                        {
                            //анализирует столбец №3 - "наличие готовности"
                            if (!checkBox_MNS_PNS.Checked) break;
                            startRow = startRow - 1; //чтобы поймать первое название
                            rangeParamName =
                                sheetSys.Range[
                                    sheetSys.Cells[startRow, colParamName - 1], sheetSys.Cells[endRow, colParamName + 1]]
                                    .Value; //читаем весь кусок этой какули
                            blockStart = false; //признак окончания блока
                            blockName = "";
                            blockNum = 0;
                            numStart = 1; //стартовый номер для блока
                            for (int i = 1; i < rangeParamName.GetLength(0); i++)
                            {
                                //порядок проверок условий важен!
                                if (rangeParamName[i, 1] == null) continue; //обработка пустых строк
                                var str = rangeParamName[i, 1].ToString();
                                int res;
                                bool isInt = Int32.TryParse(str, out res);
                                if (!isInt) //поиск начала блока
                                {
                                    blockName = rangeParamName[i, 1].ToString().Trim();
                                    numStart = 1 + 96 * blockNum;
                                    blockNum++;
                                    blockStart = true; //ставим флаг что блок кончился
                                    continue;
                                }
                                if (rangeParamName[i, 2] == null)
                                {
                                    numStart++;
                                    continue;
                                }
                                if (rangeParamName[i, 2].ToString().Trim() != "" && blockStart == false)
                                    //обработка других строк блока с данными
                                {
                                    rangeParamName[i, 1] = numStart;
                                    numStart++;
                                    rangeParamName[i, 2] = "ГОТОВНОСТЬ " + blockName + ". " + rangeParamName[i, 2].ToString();
                                }
                                if (rangeParamName[i, 2].ToString().Trim() != "" && blockStart)
                                    //обработка первой строки блока с данными
                                {
                                    blockStart = false;
                                    rangeParamName[i, 1] = numStart;
                                    numStart++;
                                    rangeParamName[i, 2] = "ГОТОВНОСТЬ " + blockName + ". " + rangeParamName[i, 2].ToString();
                                }
                            }
                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeParamName.GetLength(0) - 1; i++)
                            {
                                if (rangeParamName[i + 1, 2] == null) continue;
                                if (rangeParamName[i + 1, 3] == null) continue;
                                if (rangeParamName[i + 1, 2].ToString().Trim() == "") continue; //пропускаем пропуски
                                if (rangeParamName[i + 1, 3].ToString().Trim() == "") continue; //пропускаем пропуски
                                rowObjects[1] = row[1]; //sysId
                                rowObjects[2] = rangeParamName[i + 1, 1];
                                rowObjects[3] = rangeParamName[i + 1, 2];
                                rowObjects[4] = -1;
                                rowObjects[5] = -1;
                                rowObjects[6] = -1;
                                if (Program.Settings.DB_version > 6)
                                {
                                    rowObjects[7] = -1;
                                    rowObjects[8] = -1;
                                    rowObjects[9] = -1;
                                }
                                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в KGMPNA";
                        }

                        #endregion KTPRA
                        break;

                    #endregion

                    #region Диагностика

                    case "diagUSO":
                        #region Диагностика УСО

                        try
                        {
                            rangeParamName =
                                sheetSys.Range[sheetSys.Cells[startRow, colParamName], sheetSys.Cells[endRow, colParamName]]
                                    .Value;
                            rangeParamSysNum =
                                sheetSys.Range[
                                    sheetSys.Cells[startRow, colParamSysNum], sheetSys.Cells[endRow, colParamSysNum]].Value;

                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeParamName.GetLength(0) - 1; i++)
                            {
                                if (rangeParamName[i + 1, 1] == null || rangeParamName[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                                if (rangeParamSysNum[i + 1, 1] == null || rangeParamSysNum[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски и КЦ
                                rowObjects[1] = row[1]; //sysId
                                rowObjects[2] = Convert.ToInt32(rangeParamSysNum[i + 1, 1]);
                                rowObjects[3] = rangeParamName[i + 1, 1];
                                rowObjects[4] = -1;
                                rowObjects[5] = -1;
                                rowObjects[6] = -1;
                                if (Program.Settings.DB_version > 6)
                                {
                                    rowObjects[7] = -1;
                                    rowObjects[8] = -1;
                                    rowObjects[9] = -1;
                                }
                                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagUSO";
                        }

                        #endregion
                        break;
                    case "diagDrop":
                        #region Диагностика корзин ввода-вывода
                        try
                        {
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeDrop.GetLength(0) - 1; i++)
                            {
                                if (rangeDrop[i + 1, 1] == null) continue;
                                if (rangeDrop[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                                if (rangeDrop[i + 1, 1].ToString().ToUpper().Contains("САР")) continue; //пропускаем пропуски strUSO.ToUpper().Contains("САР"))
                                if (rangeDrop[i + 1, 5] == null) continue;
                                if (rangeDrop[i + 1, 5].ToString().Trim() == "") continue; //пропускаем пропуски
                                rowObjects[1] = row[1]; //sysId
                                rowObjects[2] = (int)(rangeDrop[i + 1, 5]);
                                rowObjects[3] = rangeDrop[i + 1, 1].ToString() + " " + rangeDrop[i + 1, 7].ToString();
                                rowObjects[4] = -1;
                                rowObjects[5] = -1;
                                rowObjects[6] = -1;
                                if (Program.Settings.DB_version > 6)
                                {
                                    rowObjects[7] = -1;
                                    rowObjects[8] = -1;
                                    rowObjects[9] = -1;
                                }
                                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagDrop";
                        }
                        #endregion
                        break;
                    case "diagDropSARD":
                        //ошибка с дублирование sysnum?
                        #region Диагностика корзин ввода-вывода
                        //заполняем datatable
                        try
                        {
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeDropSARD.GetLength(0) - 1; i++)
                            {
                                if (rangeDropSARD[i + 1, 1] == null) continue;
                                if (rangeDropSARD[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                                if (rangeDropSARD[i + 1, 5] == null) continue;
                                if (rangeDropSARD[i + 1, 5].ToString().Trim() == "") continue; //пропускаем пропуски
                                if (rangeDropSARD[i + 1, 1].ToString().ToUpper().Contains("САР"))
                                {
                                    rowObjects[1] = row[1]; //sysId

                                    rowObjects[2] = (int) (rangeDropSARD[i + 1, 5]);
                                    rowObjects[3] = rangeDropSARD[i + 1, 1].ToString() + " " +
                                                    rangeDropSARD[i + 1, 7].ToString();
                                    rowObjects[4] = -1;
                                    rowObjects[5] = -1;
                                    rowObjects[6] = -1;

                                    if (Program.Settings.DB_version > 6)
                                    {
                                        rowObjects[7] = -1;
                                        rowObjects[8] = -1;
                                        rowObjects[9] = -1;
                                    }

                                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                }
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagDropSARD";
                        }

                        #endregion
                        break;
                    case "diagPLC":
                        #region Диагностика ПЛК

                        int plcNum = 1;
                        //заполняем datatable
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int i = 0; i < rangeDrop.GetLength(0) - 1; i++)
                        {
                            if (rangeDrop[i + 1, 1] == null) continue;
                            if (rangeDrop[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                            if (rangeDrop[i + 1, 4] == null) continue;
                            if (rangeDrop[i + 1, 4].ToString().Trim() == "") continue; //пропускаем пропуски
                            var str = rangeDrop[i + 1, 4].ToString();
                            if (!str.Contains("КЦ")) continue; //пропускаем если не основные контроллеры
                            rowObjects[1] = row[1]; //sysId
                            rowObjects[2] = plcNum;
                            plcNum++;
                            rowObjects[3] = rangeDrop[i + 1, 1].ToString() + " " + rangeDrop[i + 1, 4].ToString() + " " + rangeDrop[i + 1, 7].ToString();
                            rowObjects[4] = -1;
                            rowObjects[5] = -1;
                            rowObjects[6] = -1;
                            if (Program.Settings.DB_version > 6)
                            {
                                rowObjects[7] = -1;
                                rowObjects[8] = -1;
                                rowObjects[9] = -1;
                            }
                            Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                        }

                        #endregion
                        break;
                    case "diagNet":
                        //нет
                        break;
                    case "diagModDrop":
                        #region Диагностика модулей в корзинах

                        //dynamic rangeDrop = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 23]].Value; //весь дроп
                        //заполняем datatable
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int i = 0; i < rangeDrop.GetLength(0) - 1; i++)
                        {
                            if (rangeDrop[i + 1, 5] != null && rangeDrop[i + 1, 5].ToString().Trim() != "")
                            {
                                for (int j = 8; j < rangeDrop.GetLength(1); j++)
                                {
                                    if (rangeDrop[i + 1, j] == null || rangeDrop[i + 1, j].ToString().Trim() == "") continue;
                                    if (rangeDrop[i + 2, j] == null || rangeDrop[i + 2, j].ToString().Trim() == "") continue;
                                    if (rangeDrop[i + 1, 1].ToString().ToUpper().Contains("САР")) continue;

                                    rowObjects[1] = row[1]; //sysId
                                    rowObjects[2] = 16 * ((int)rangeDrop[i + 1, 5] - 1) + (j - 7);
                                    rowObjects[3] = rangeDrop[i + 1, 1].ToString() + " " + rangeDrop[i + 1, j].ToString() + " (" + rangeDrop[i + 2, j].ToString() + ")";
                                    rowObjects[4] = -1;
                                    rowObjects[5] = -1;
                                    rowObjects[6] = -1;
                                    if (Program.Settings.DB_version > 6)
                                    {
                                        rowObjects[7] = -1;
                                        rowObjects[8] = -1;
                                        rowObjects[9] = -1;
                                    }
                                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку  
                                }

                            }
                        }

                    #endregion
                        break;
                    case "diagModDropSARD":
                        #region Диагностика модулей в корзинах

                        try
                        {
                            //dynamic rangeDrop = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 23]].Value; //весь дроп
                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeDropSARD.GetLength(0) - 1; i++)
                            {
                                if (rangeDropSARD[i + 1, 5] != null && rangeDropSARD[i + 1, 5].ToString().Trim() != "")
                                {
                                    for (int j = 8; j < rangeDropSARD.GetLength(1); j++)
                                    {
                                        if (rangeDropSARD[i + 1, j] == null ||
                                            rangeDropSARD[i + 1, j].ToString().Trim() == "") continue;
                                        if (rangeDropSARD[i + 2, j] == null ||
                                            rangeDropSARD[i + 2, j].ToString().Trim() == "") continue;
                                        if (rangeDropSARD[i + 1, 1].ToString().ToUpper().Contains("САР"))
                                        {
                                            rowObjects[1] = row[1]; //sysId

                                            rowObjects[2] = 16*((int) rangeDropSARD[i + 1, 5] - 1) + (j - 7);
                                            rowObjects[3] = rangeDropSARD[i + 1, 1].ToString() + " " +
                                                            rangeDropSARD[i + 1, j].ToString() + " (" +
                                                            rangeDropSARD[i + 2, j].ToString() + ")";
                                            rowObjects[4] = -1;
                                            rowObjects[5] = -1;
                                            rowObjects[6] = -1;
                                            if (Program.Settings.DB_version > 6)
                                            {
                                                rowObjects[7] = -1;
                                                rowObjects[8] = -1;
                                                rowObjects[9] = -1;
                                            }
                                            Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку  
                                        }
                                    }
                                
                                }
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagModDropSARD";
                        }

                        #endregion
                        break;
                    case "diagModPLC":
                        #region Диагностика модулей ПЛК

                        try
                        {
                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            numPlc = 0;
                            for (int i = 0; i < rangeDrop.GetLength(0) - 1; i++)
                            {
                                if (rangeDrop[i + 1, 4] == null) continue;
                                string str = rangeDrop[i + 1, 4].ToString();
                                if (str.ToUpper().Contains("КЦ") || str.ToUpper().Contains("КОНТРОЛЛЕР")) //пропускаем если не плк
                                {
                                    numPlc++;
                                    for (int j = 8; j < rangeDrop.GetLength(1); j++)
                                    {
                                        if (rangeDrop[i + 1, j] == null || rangeDrop[i + 1, j].ToString().Trim() == "") continue;
                                        if (rangeDrop[i + 2, j] == null || rangeDrop[i + 2, j].ToString().Trim() == "") continue;
                                        rowObjects[1] = row[1]; //sysId
                                        rowObjects[2] = 16 * (numPlc - 1) + (j - 7);
                                        rowObjects[3] = rangeDrop[i + 1, 4].ToString() + " " + rangeDrop[i + 1, j].ToString() + " (" + rangeDrop[i + 2, j].ToString() + ")";
                                        rowObjects[4] = -1;
                                        rowObjects[5] = -1;
                                        rowObjects[6] = -1;
                                        if (Program.Settings.DB_version > 6)
                                        {
                                            rowObjects[7] = -1;
                                            rowObjects[8] = -1;
                                            rowObjects[9] = -1;
                                        }
                                        Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку  
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagModPLC";
                        }

                    #endregion

                        break;
                    case "diagModPLCSARD":
                        #region Диагностика модулей ПЛК САРД
                        if (checkBox_SAR.Checked) break;
                        if (checkBox_PT.Checked) break;
                        if (checkBox_SIKN.Checked) break;
                        //заполняем datatable
                        try
                        {
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            numPlc = 0;

                            if (dropFirstReadSARD == true) rangeDropSARD = rangeDrop; // проверить, работает ли присвоение по ссылке

                            for (int i = 0; i < rangeDropSARD.GetLength(0) - 1; i++)
                            {
                                if (rangeDropSARD[i + 1, 4] == null) continue;
                                string strUSO = rangeDropSARD[i + 1, 1].ToString();
                                string str = rangeDropSARD[i + 1, 4].ToString();

                                if ((str.ToUpper().Contains("КЦ") || str.ToUpper().Contains("КОНТРОЛЛЕР")) & strUSO.ToUpper().Contains("САР"))  //пропускаем если не плк
                                {
                                    numPlc++;
                                    for (int j = 8; j < rangeDropSARD.GetLength(1); j++)
                                    {
                                        if (rangeDropSARD[i + 1, j] == null || rangeDropSARD[i + 1, j].ToString().Trim() == "") continue;
                                        if (rangeDropSARD[i + 2, j] == null || rangeDropSARD[i + 2, j].ToString().Trim() == "") continue;
                                        rowObjects[1] = row[1]; //sysId
                                        rowObjects[2] = 16 * (numPlc - 1) + (j - 7);
                                        rowObjects[3] = rangeDropSARD[i + 1, 1].ToString() + ". " + rangeDropSARD[i + 1, 4].ToString() + " " + rangeDropSARD[i + 1, j].ToString() + " (" + rangeDropSARD[i + 2, j].ToString() + ")";
                                        rowObjects[4] = -1;
                                        rowObjects[5] = -1;
                                        rowObjects[6] = -1;
                                        if (Program.Settings.DB_version > 6)
                                        {
                                            rowObjects[7] = -1;
                                            rowObjects[8] = -1;
                                            rowObjects[9] = -1;
                                        }
                                        Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку  
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagModPLCSARD";
                        }

                        #endregion
                        break;
                    case "diagConnect":
                        #region Диагностика связи с устройствами по интерфейсу

                        try
                        {
                            rangeParamName =
                                sheetSys.Range[sheetSys.Cells[startRow, colParamName], sheetSys.Cells[endRow, colParamName]]
                                    .Value;
                            rangeParamSysNum =
                                sheetSys.Range[
                                    sheetSys.Cells[startRow, colParamSysNum], sheetSys.Cells[endRow, colParamSysNum]].Value;

                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeParamName.GetLength(0) - 1; i++)
                            {
                                if (rangeParamName[i + 1, 1] == null || rangeParamName[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                                if (rangeParamSysNum[i + 1, 1] == null || rangeParamSysNum[i + 1, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                                rowObjects[1] = row[1]; //sysId
                                rowObjects[2] = Convert.ToInt32(rangeParamSysNum[i + 1, 1]);
                                rowObjects[3] = rangeParamName[i + 1, 1];
                                rowObjects[4] = -1;
                                rowObjects[5] = -1;
                                rowObjects[6] = -1;
                                if (Program.Settings.DB_version > 6)
                                {
                                    rowObjects[7] = -1;
                                    rowObjects[8] = -1;
                                    rowObjects[9] = -1;
                                }
                                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в diagConnect";
                        }

                        #endregion
                        break;
                    case "diagSwitch":
                            GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                        #endregion

                        #region Пожарка

                    case "APT_PI":
                        if (!checkBox_PT.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                    case "APT_ST":
                        if (!checkBox_PT.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                    case "APT_GPZ":
                        if (!checkBox_PT.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                    case "APT_ATP":
                        if (!checkBox_PT.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                    case "APT_UBD":
                        if (!checkBox_PT.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                    case "RPZV":
                        if (!checkBox_PT.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;

                    #endregion

                        #region САРД

                        case "SARD_SetPoint_Spec":
                            if (checkBox_PT.Checked) break;
                            if (checkBox_SIKN.Checked) break;

                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                            break;
                        case "SARD_Mode":
                            if (checkBox_PT.Checked) break;
                            if (checkBox_SIKN.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                            break;
                        case "SARD_RD":
                            if (checkBox_PT.Checked) break;
                            if (checkBox_SIKN.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                            break;
                        case "SARD_Other":
                            if (checkBox_PT.Checked) break;
                            if (checkBox_SIKN.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                            break;

                        #endregion

                        #region Резервуарный парк

                    case "KTPRR":
                        if (!checkBox_RP.Checked) break;
                        #region KTPRR

                        try
                        {
                            //НЕ анализирует столбец - "наличие защиты" в перспективе надо сделать унифицировано с готовностями
                            startRow = startRow - 1; //чтобы поймать первое название
                            rangeParamName =
                                sheetSys.Range[
                                    sheetSys.Cells[startRow, colParamName - 1], sheetSys.Cells[endRow, colParamName]]
                                    .Value; //читаем весь кусок этой какули
                            blockStart = false; //признак окончания блока
                            blockName = "";
                            blockNum = 0;
                            numStart = 1; //стартовый номер для блока
                            for (int i = 1; i < rangeParamName.GetLength(0); i++)
                            {
                                //порядок проверок условий важен!
                                if (rangeParamName[i, 1] == null) continue; //обработка пустых строк
                                string str = rangeParamName[i, 1].ToString();
                                int res;
                                bool isInt = Int32.TryParse(str, out res);
                                if (!isInt) //поиск начала блока
                                {
                                    blockName = rangeParamName[i, 1].ToString().Trim();
                                    numStart = 1 + 96 * blockNum;
                                    blockNum++;
                                    blockStart = true; //ставим флаг что блок кончился
                                    continue;
                                }
                                if (rangeParamName[i, 2] == null)
                                {
                                    numStart++;
                                    continue;
                                }
                                if (rangeParamName[i, 2].ToString().Trim() != "" && blockStart == false)
                                    //обработка других строк блока с данными
                                {
                                    rangeParamName[i, 1] = numStart;
                                    numStart++;
                                    rangeParamName[i, 2] = blockName + ". " + rangeParamName[i, 2].ToString();
                                }
                                if (rangeParamName[i, 2].ToString().Trim() != "" && blockStart)
                                    //обработка первой строки блока с данными
                                {
                                    blockStart = false;
                                    rangeParamName[i, 1] = numStart;
                                    numStart++;
                                    rangeParamName[i, 2] = blockName + ". " + rangeParamName[i, 2].ToString();
                                }
                            }
                            //заполняем datatable
                            rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                            for (int i = 0; i < rangeParamName.GetLength(0) - 1; i++)
                            {
                                if (rangeParamName[i + 1, 2] == null) continue;
                                if (rangeParamName[i + 1, 2].ToString().Trim() == "") continue; //пропускаем пропуски
                                rowObjects[1] = row[1]; //sysId
                                rowObjects[2] = rangeParamName[i + 1, 1];
                                rowObjects[3] = rangeParamName[i + 1, 2];
                                rowObjects[4] = -1;
                                rowObjects[5] = -1;
                                rowObjects[6] = -1;
                                if (Program.Settings.DB_version > 6)
                                {
                                    rowObjects[7] = -1;
                                    rowObjects[8] = -1;
                                    rowObjects[9] = -1;
                                }
                                Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                            }
                        }
                        catch (Exception)
                        {
                            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка в KTPRR";
                        }

                        #endregion KTPRA
                        break;
                    case "Tank":
                        if (!checkBox_RP.Checked) break;
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;

                    #endregion

                    //Обработка простых систем
                    default:
                        GetDefaultObject(sheetSys, startRow, endRow, colParamName, colParamSysNum, colParamId, sysName, row);
                        break;
                }
                //имена для диагностики
            }
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows) //заполняем -1
            {
                int upCol;
                if (Program.Settings.DB_version < 7) upCol = 6; else upCol = 9;
                for (int i = 4; i <= upCol; i++)
                {
                    row[i] = -1;
                }
            }
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                                         Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
        }

        private void GetDefaultObject(Worksheet sheetSys, int startRow, int endRow, int colParamName, int colParamSysNum,
            int colParamId, string sysName, DataRow row)
        {
            try
            {
                dynamic rangeParamName = sheetSys.Range[sheetSys.Cells[startRow, colParamName], sheetSys.Cells[endRow, colParamName]]
                    .Value;
                dynamic rangeParamSysNum = sheetSys.Range[
                    sheetSys.Cells[startRow, colParamSysNum], sheetSys.Cells[endRow, colParamSysNum]].Value;
                if (colParamId != 0) //если у параметра есть идентификатор
                {
                    dynamic rangeParamId =
                        sheetSys.Range[sheetSys.Cells[startRow, colParamId], sheetSys.Cells[endRow, colParamId]]
                            .Value;
                    //преобразуем имена параметров
                    for (int i = 1; i <= rangeParamName.GetLength(0); i++)
                    {
                        if (rangeParamName[i, 1] == null) continue;
                        if (rangeParamName[i, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                        if (rangeParamId[i, 1] == null) continue;
                        if (rangeParamId[i, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                        rangeParamName[i, 1] =
                            (object)
                                (rangeParamName[i, 1].ToString() + " (" + rangeParamId[i, 1].ToString() + ")");
                    }
                }
                //заполняем datatable
                var rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count];
                if (rangeParamName is string) //случай когда всего одна строка в таблице и массив вырождается в строку
                {
                    if(rangeParamSysNum.ToString().Trim() == "" || rangeParamName.ToString().Trim() == "") return;
                    rowObjects[1] = row[1]; //sysId
                    rowObjects[2] = rangeParamSysNum;
                    rowObjects[3] = rangeParamName;
                    rowObjects[4] = -1;
                    rowObjects[5] = -1;
                    rowObjects[6] = -1;
                    if (Program.Settings.DB_version > 6)
                    {
                        rowObjects[7] = -1;
                        rowObjects[8] = -1;
                        rowObjects[9] = -1;
                    }
                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                    return;
                }

                for (int i = 0; i <= rangeParamName.GetLength(0) - 1; i++)
                {
                    if (rangeParamName[i + 1, 1] == null || rangeParamName[i + 1, 1].ToString().Trim() == "")
                        continue; //пропускаем пропуски
                    rowObjects[1] = row[1]; //sysId
                    rowObjects[2] = rangeParamSysNum[i + 1, 1];
                    if (row[2].ToString() == "KTPR")
                    {
                        rowObjects[3] = "ОБЩЕСТАНЦИОННАЯ ЗАЩИТА " + rangeParamName[i + 1, 1];
                    }
                    else
                    {
                        rowObjects[3] = rangeParamName[i + 1, 1];
                    }             
                    rowObjects[4] = -1;
                    rowObjects[5] = -1;
                    rowObjects[6] = -1;
                    if (Program.Settings.DB_version > 6)
                    {
                        rowObjects[7] = -1;
                        rowObjects[8] = -1;
                        rowObjects[9] = -1;
                    }
                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                }
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" Object: ошибка sheetSys =" + sheetSys.Name;
            }
        }

        private bool CheckObject() //проверка
        {
            ResetColorTable(dataGridView_Object);
            bool res = false;
            foreach (DataGridViewRow row in dataGridView_Object.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (row.Cells[i].Value.ToString() == "")
                    {
                        row.Cells[i].Style.BackColor = Color.Red;
                        res = true;
                    }
                }
            }
            if (dataGridView_Object.Rows.Count == 0) res = true;
            return res;
        }

        private void WriteObject()
        {
            string sysName = "Object";
            Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
            {
                foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                {
                    if (row[column].ToString() == "") row[column] = "NaN";
                }
                row[3] = row[3].ToString().ToUpper();

                if (Program.Settings.DB_version < 7)
                {
                    _oleDb.WriteOperMess_Object((uint)row.ItemArray[1], (uint)row.ItemArray[2],
                        row.ItemArray[3].ToString(), (int) row.ItemArray[4], (int)row.ItemArray[5], (int)row.ItemArray[6]);
                }
                else
                {
                    _oleDb.WriteOperMess_Object((uint)row.ItemArray[1], (uint)row.ItemArray[2],
                        row.ItemArray[3].ToString(), (int)row.ItemArray[4], (int)row.ItemArray[5], (int)row.ItemArray[6], (string)row.ItemArray[7], (string)row.ItemArray[8], (string)row.ItemArray[9]);
                }
            }
            Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";
        }

        #endregion Object

        #region Timer

        private void ReadTimer() //чтение = генерация из Excel
        {
            if (checkBox_SAR.Checked) return;
            var sysName = "Timer";
            string[] adressTimer = Program.Settings.TimerAdressSettings.Split(';'); //адреса временных настроек
            string[] listTimer = Program.Settings.TimerListSettings.Split(';'); //листы временных настроек
            string[] strTimer = Program.Settings.TimerStrSettings.Split(';'); //нач строки временных настроек
            Dictionary<int, string> placeTimerDictionary = new Dictionary<int, string>();
            int numPlace = 1;
            int idIterator = 1;
            int[] pojZoneAddr = new int[100];
            //var sheetOip = _fileIo.GetSheet("ИП");

            for (int index = 0; index < listTimer.Length; index++)
            {
                string listTimerName = listTimer[index];
                var sheetSys = _fileIo.GetSheet(listTimerName);
                if (sheetSys == null) continue;
                int startRow = Convert.ToInt32(strTimer[index]);
                int endRow = _fileIo.LastRowCell(ref sheetSys); //окончание определяется любым значением последней ячейки, даже рамка ячейки будет определяться концом
                int pojZoneAddrIndex = 0;
                dynamic rangeTable; //основная таблица куда читаем из ЁкСеЛя
                int adressIterator;
                switch (listTimerName)  //Рез;twList4;twNA;twCommTimes;twZDV;twVS
                {
                    case "Станц Защ":
                        #region Станц Защ
                        if (checkBox_PT.Checked) break;
                        string sysTimerName = "СТАНЦИОННЫЕ ЗАЩИТЫ";
                        int startrowStZash = Program.Settings.startRow_StZash;
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 20]].Value;
                        adressIterator = 0;
                        //заполняем datatable
                        var rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count];
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {
                            if (rangeTable[i, 19] == null || rangeTable[i, 19].ToString() != "1")
                            {
                                adressIterator++;
                                continue; //Проверка на наличие защиты
                            }
                            //if (rangeTable[i, 19] != null) // Столбец "Наличие" имеет номер 19.
                            // {
                            //if (rangeTable[i, 19].ToString() == "1") // Столбец "Наличие" имеет номер 19.
                            //{
                            if (rangeTable[i, 2] == null || rangeTable[i, 2].ToString().Trim() == "")
                                continue; //пропускаем пропуски в имени
                            rowObjects[1] = rangeTable[i, 2]; //имя
                            if (rangeTable[i, 3] == null || rangeTable[i, 3].ToString().Trim() == "")
                            {
                                rowObjects[2] = "T" + idIterator; //ID
                                idIterator++;
                            }
                            else
                            {
                                rowObjects[2] = rangeTable[i, 3]; //ID
                            }
                            rowObjects[3] = "СЕК"; //ед изм
                            rowObjects[4] = rangeTable[i, 7]; //уставка времени
                            if (_useAdrBegDenominator)
                            {
                                rowObjects[5] = 10;
                            }
                            else
                            {
                                rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                            }

                            //AdressBeginning
                            adressIterator++;
                            rowObjects[6] = numPlace; //Place
                            rowObjects[7] = "1"; // PLC
                            if (Program.Settings.DB_version >= 9)
                            {
                                string stationTag =
                                    Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                //получим иерархию проекта
                                rowObjects[8] = stationTag +
                                                Convert.ToString(
                                                    _fileIo.GetSheet("Станц Защ").Cells[
                                                        startrowStZash - 1 + i, 22]
                                                        .Value); // 22 - номер столбца Иерархии в конфигураторе
                                rowObjects[9] = sysTimerName; //заполняем PLACENAME
                            }

                            Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                                                               //}
                                                                               //}
                        }
                        if (!placeTimerDictionary.ContainsKey(numPlace)) placeTimerDictionary.Add(numPlace, sysTimerName); //кладем Place в словарь
                        numPlace++;
                        #endregion
                        break;
                    case "Защ НА":
                        #region Защ НА
                        if (!checkBox_MNS_PNS.Checked) break;
                        if (checkBox_SIKN.Checked) break;
                        if (checkBox_PT.Checked) break;
                        sysTimerName = "АГРЕГАТНЫЕ ЗАЩИТЫ";
                        startRow = startRow - 1; //чтобы поймать первое название
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 16]].Value;
                        var blockStart = false; //признак окончания блока
                        var blockName = "";
                        var blockNum = 0;
                        var numStart = 1; //стартовый номер для блока
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {
                            // if (rangeTable[i, 14] != null && rangeTable[i, 14].ToString() == "1") //Проверка на наличие защиты
                            if (rangeTable[i, 14] == null || rangeTable[i, 14].ToString() != "1") continue; //Проверка на наличие защиты
  
                            //порядок проверок условий важен!
                            if (rangeTable[i, 1] == null) continue; //обработка пустых строк
                            var str = rangeTable[i, 1].ToString();
                            int res;
                            bool isInt = Int32.TryParse(str, out res);
                            if (!isInt) //поиск начала блока
                            {
                                blockName = rangeTable[i, 1].ToString().Trim();
                                numStart = 1 + 96 * blockNum;
                                blockNum++;
                                blockStart = true; //ставим флаг что блок кончился
                                rangeTable[i, 6] = blockNum; //сюда запишем номер блока для place
                                rangeTable[i, 7] = blockName;
                                continue;
                            }
                            if (rangeTable[i, 2] == null)
                            {
                                numStart++;
                                continue;
                            }
                            if (rangeTable[i, 2].ToString().Trim() != "" && blockStart == false)
                            //обработка других строк блока с данными
                            {
                                rangeTable[i, 6] = blockNum; //сюда запишем номер блока для place
                                rangeTable[i, 7] = blockName;
                                rangeTable[i, 1] = numStart;
                                numStart++;
                                rangeTable[i, 2] = blockName + ". " + rangeTable[i, 2].ToString();
                            }
                            if (rangeTable[i, 2].ToString().Trim() != "" && blockStart)
                            //обработка первой строки блока с данными
                            {
                                rangeTable[i, 6] = blockNum; //сюда запишем номер блока для place
                                rangeTable[i, 7] = blockName;
                                blockStart = false;
                                rangeTable[i, 1] = numStart;
                                numStart++;
                                rangeTable[i, 2] = blockName + ". " + rangeTable[i, 2].ToString();
                            }

                        }
                        //заполняем datatable
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        adressIterator = -1;
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {

                            //adressIterator++;
                            if (rangeTable[i, 2] == null || rangeTable[i, 2].ToString().Trim() == "") continue;  //пропускаем пропуски
                            if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                            if (rangeTable[i, 14] == null || rangeTable[i, 14].ToString() != "1") continue; //пропускаем если нет наличия
                            //if (rangeTable[i, 16] == null || rangeTable[i, 16].ToString() == "0") continue; //пропускаем если нет наличия

                            rowObjects[1] = rangeTable[i, 2]; //имя
                            if (rangeTable[i, 3] == null || rangeTable[i, 3].ToString().Trim() == "")
                            {
                                rowObjects[2] = "T" + idIterator; //ID
                                idIterator++;
                            }
                            else
                            {
                                rowObjects[2] = rangeTable[i, 3]; //ID
                            }
                            rowObjects[3] = "СЕК"; //ед изм
                            rowObjects[4] = rangeTable[i, 5]; //уставка времени
                            if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                            {
                                rowObjects[5] = 10;
                            }
                            else
                            {
                                rowObjects[5] = Convert.ToInt32(adressTimer[index]) + (int)rangeTable[i, 1] - 1; //AdressBeginning
                            }
                            //adressIterator++;
                            rowObjects[6] = numPlace + (int)rangeTable[i, 6] - 1; //Place
                            rowObjects[7] = "1"; //PLC
                            if (Program.Settings.DB_version >= 9)
                            {
                                if (rangeTable[i, 14] == null || rangeTable[i, 14].ToString() != "1") continue; //пропускаем если нет наличия
                                string stationTag = Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + "."; //получим иерархию проекта
                                rowObjects[8] = stationTag + rangeTable[i, 16];
                                rowObjects[9] = sysTimerName + " " + rangeTable[i, 7].ToString(); //заполняем PLACENAME
                            }
                            if (!placeTimerDictionary.ContainsKey((int)rowObjects[6])) placeTimerDictionary.Add((int)rowObjects[6], sysTimerName + " " + rangeTable[i, 7].ToString()); //кладем Place в словарь
                            Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку 
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "Защ Рез":
                        #region Защ Рез
                        if (!checkBox_RP.Checked) break;
                        if (checkBox_PT.Checked) break;
                        sysTimerName = "ЗАЩИТЫ РЕЗЕРВУАРНОГО ПАРКА";
                        startRow = startRow - 1; //чтобы поймать первое название
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 16]].Value;
                        blockStart = false; //признак окончания блока
                        blockName = "";
                        blockNum = 0;
                        numStart = 1; //стартовый номер для блока
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {
                            if (rangeTable[i, 14] == null || rangeTable[i, 14].ToString() != "1") continue;
                            //порядок проверок условий важен!
                            if (rangeTable[i, 1] == null) continue; //обработка пустых строк
                            var str = rangeTable[i, 1].ToString();
                            int res;
                            bool isInt = Int32.TryParse(str, out res);
                            if (!isInt) //поиск начала блока
                            {
                                blockName = rangeTable[i, 1].ToString().Trim();
                                numStart = 1 + 8 * blockNum;
                                blockNum++;
                                blockStart = true; //ставим флаг что блок кончился
                                rangeTable[i, 6] = blockNum; //сюда запишем номер блока для place
                                rangeTable[i, 7] = blockName;
                                continue;
                            }
                            if (rangeTable[i, 2] == null)
                            {
                                numStart++;
                                continue;
                            }
                            if (rangeTable[i, 2].ToString().Trim() != "" && blockStart == false)
                            //обработка других строк блока с данными
                            {
                                rangeTable[i, 6] = blockNum; //сюда запишем номер блока для place
                                rangeTable[i, 7] = blockName;
                                rangeTable[i, 1] = numStart;
                                numStart++;
                                rangeTable[i, 2] = blockName + ". " + rangeTable[i, 2].ToString();
                            }
                            if (rangeTable[i, 2].ToString().Trim() != "" && blockStart)
                            //обработка первой строки блока с данными
                            {
                                rangeTable[i, 6] = blockNum; //сюда запишем номер блока для place
                                rangeTable[i, 7] = blockName;
                                blockStart = false;
                                rangeTable[i, 1] = numStart;
                                numStart++;
                                rangeTable[i, 2] = blockName + ". " + rangeTable[i, 2].ToString();
                            }
                        }
                        //заполняем datatable
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {
                            if (rangeTable[i, 2] == null || rangeTable[i, 2].ToString().Trim() == "") continue; //пропускаем пропуски
                            if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                            if (rangeTable[i, 14] == null || rangeTable[i, 14].ToString() != "1") continue; //пропускаем если нет наличия
                            rowObjects[1] = rangeTable[i, 2]; //имя
                            if (rangeTable[i, 3] == null || rangeTable[i, 3].ToString().Trim() == "")
                            {
                                rowObjects[2] = "T" + idIterator; //ID
                                idIterator++;
                            }
                            else
                            {
                                rowObjects[2] = rangeTable[i, 3]; //ID
                            }
                            rowObjects[3] = "СЕК"; //ед изм
                            rowObjects[4] = rangeTable[i, 5]; //уставка времени
                            if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                            {
                                rowObjects[5] = 10;
                            }
                            else
                            {
                                rowObjects[5] = Convert.ToInt32(adressTimer[index]) + (int)rangeTable[i, 1] - 1;
                            }
                           
                            //AdressBeginning
                            //adressIterator++;
                            rowObjects[6] = numPlace + (int)rangeTable[i, 6] - 1; //Place
                            rowObjects[7] = "1"; //PLC
                            if (Program.Settings.DB_version >= 9)
                            {
                                string stationTag =
                                    Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                //получим иерархию проекта
                                rowObjects[8] = stationTag + rangeTable[i, 16];
                                rowObjects[9] = sysTimerName + " " + rangeTable[i, 7].ToString(); //заполняем PLACENAME
                            }
                            if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                placeTimerDictionary.Add((int)rowObjects[6], sysTimerName + " " + rangeTable[i, 7].ToString()); //кладем Place в словарь
                            Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "twList4": //+prot
                        #region ВРЕМЕННЫЕ УСТАВКИ ПРЕДЕЛЬНЫХ ПАРАМЕТРОВ
                        sysTimerName = "ВРЕМЕННЫЕ УСТАВКИ ПРЕДЕЛЬНЫХ ПАРАМЕТРОВ";
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 5]].Value;
                        adressIterator = 0;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {
                            if (rangeTable[i, 4] == null || rangeTable[i, 4].ToString().Trim() == "" || rangeTable[i, 4].ToString().Trim() == "0")
                            {
                                adressIterator++;
                                continue; //пропускаем пропуски
                            }
                            if (rangeTable[i, 4] != null) // Столбец "Наличие" имеет номер 4.
                            {
                                if (rangeTable[i, 4].ToString() == "1") // Столбец "Наличие" имеет номер 4.
                                {
                                    if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "")
                                        continue; //пропускаем пропуски
                                    if (rangeTable[i, 2] == null || rangeTable[i, 2].ToString().Trim() == "")
                                        continue; //пропускаем пропуски
                                    rowObjects[1] = rangeTable[i, 2].ToString(); //имя
                                    rowObjects[2] = rangeTable[i, 3]; //ID
                                    rowObjects[3] = "СЕК"; //ед изм
                                    rowObjects[4] = "0"; //уставка времени
                                    if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                                    {
                                        rowObjects[5] = 10;
                                    }
                                    else
                                    {
                                        rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                                    }
                                    
                                    //AdressBeginning
                                    rowObjects[6] = numPlace; //Place
                                    rowObjects[7] = "1"; //PLC
                                    if (Program.Settings.DB_version >= 9)
                                    {
                                        string stationTag =
                                            Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                        //получим иерархию проекта
                                        rowObjects[8] = stationTag + rangeTable[i, 5];
                                        rowObjects[9] = sysTimerName; //заполняем PLACENAME
                                    }
                                    adressIterator++;
                                    if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                        placeTimerDictionary.Add((int)rowObjects[6], sysTimerName);
                                    //кладем Place в словарь
                                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                }
                            }
                        }
                        if (rowObjects[6] != null) numPlace = (int)rowObjects[6] + 1;
                        //numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "twNA": //+prot
                        #region Насосы
                        if (!checkBox_MNS_PNS.Checked) break;
                        if (checkBox_SIKN.Checked) break;
                        if (checkBox_PT.Checked) break;
                        sysTimerName = "ВРЕМЕННЫЕ УСТАВКИ НА";
                        //int endRowTwNA = Program.Settings.Quntity_twCommTimes + 1; -закомментил на 2017/11/25 так как ограничивать кол-во уставок TwNA не уместно, кол-во должно определяться динамически
                        // rangeTable = sheetSys.Range[sheetSys.Cells[startRow - 1, 1], sheetSys.Cells[endRowTwNA,  20]].Value; --Кузьмин
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow - 1, 1], sheetSys.Cells[endRow, 20]].Value;
                        // dynamic rangeNameNA = _fileIo.GetSheet("НА").Range[sheetSys.Cells[4, 2], sheetSys.Cells[19, 2]].Value;
                        adressIterator = 0;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int j = 4; j <= rangeTable.GetLength(1) - 1; j++) // 4-й столбец соответствует МНА-1; GetLen(1) = в ширину; 1- потому что последний стоблец = Иерархия;
                        {
                            for (int i = 2; i <= rangeTable.GetLength(0); i++)
                            {
                                if (rangeTable[i, j] == null || rangeTable[i, j].ToString().Trim() == "" || rangeTable[i, j].ToString().Trim() == "0")
                                {
                                    adressIterator++;
                                    continue; //пропускаем пропуски
                                }
                                if (rangeTable[i, j] != null)
                                {
                                    if (rangeTable[i, j].ToString() == "1")
                                    // Рабочий вариант, стоит размножить на остальные листы с защитами.
                                    {
                                        rowObjects[1] = rangeTable[1, j] + ". " + rangeTable[i, 2]; //имя
                                        rowObjects[2] = rangeTable[i, 3]; //ID
                                        rowObjects[3] = "СЕК"; //ед изм
                                        rowObjects[4] = Convert.ToInt32(rangeTable[i, j]); //уставка времени
                                        if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                                        {
                                            rowObjects[5] = 10;
                                        }
                                        else
                                        {
                                            rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                                        }
                                        //AdressBeginning
                                        rowObjects[6] = numPlace + (j - 4); //Place
                                        rowObjects[7] = "1"; //PLC
                                        if (Program.Settings.DB_version >= 9)
                                        {
                                            string stationTag =
                                                Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                            //получим иерархию проекта
                                            string mnAtag =
                                                Convert.ToString(_fileIo.GetSheet("НА").Cells[j, 159].Value) + ".";
                                            // берется с листа НА.
                                            rowObjects[8] = stationTag + mnAtag + rangeTable[i, 20];
                                            rowObjects[9] = sysTimerName + " " + rangeTable[1, j].ToString(); //заполняем PLACENAME
                                        }
                                        adressIterator++;
                                        if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                            placeTimerDictionary.Add((int)rowObjects[6],
                                                sysTimerName + " " + rangeTable[1, j].ToString());
                                        //кладем Place в словарь
                                        Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                    }
                                }
                            }
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "twCommTimes": //+prot
                        #region Общие
                        sysTimerName = "ОБЩИЕ ВРЕМЕННЫЕ УСТАВКИ";
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 5]].Value;
                        adressIterator = 0;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int i = 1; i <= rangeTable.GetLength(0); i++)
                        {
                            if (rangeTable[i, 4] == null || rangeTable[i, 4].ToString().Trim() == "" || rangeTable[i, 4].ToString().Trim() == "0")
                            {
                                adressIterator++;
                                continue; //пропускаем пропуски
                            }
                            //if (rangeTable[i, 4] != null) // Столбец "Наличие" имеет номер 4.
                            //{
                            //if (rangeTable[i, 4].ToString() == "1") // Столбец "Наличие" имеет номер 4.
                            //{

                            if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "")
                                continue; //пропускаем пропуски
                            if (rangeTable[i, 2] == null || rangeTable[i, 2].ToString().Trim() == "")
                                continue; //пропускаем пропуски
                            rowObjects[1] = rangeTable[i, 2].ToString(); //имя
                            rowObjects[2] = rangeTable[i, 3]; //ID
                            rowObjects[3] = "СЕК"; //ед изм
                            rowObjects[4] = "0"; //уставка времени
                            if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                            {
                                rowObjects[5] = 10;
                            }
                            else
                            {
                                rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                            }
                           
                            //AdressBeginning
                            rowObjects[6] = numPlace; //Place
                            rowObjects[7] = "1"; //PCL
                            if (Program.Settings.DB_version >= 9)
                            {
                                //формирование иерархии twCommTimes
                                //object[,] rangeIe = sheetSys.Range[sheetSys.Cells[startRow, 5], sheetSys.Cells[endRow, 5]].Value;
                                //5 - номер столбца иерархии
                                string stationTag =
                                    Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                //получим иерархию проекта

                                rowObjects[8] = stationTag + rangeTable[i, 5];
                                rowObjects[9] = sysTimerName; //заполняем PLACENAME
                            }
                            adressIterator++;
                            if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                placeTimerDictionary.Add((int)rowObjects[6], sysTimerName);
                            //кладем Place в словарь
                            Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                                                               // }
                                                                               // }
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "twZDV": //+prot
                        #region Задвижки
                        sysTimerName = "ВРЕМЕННЫЕ УСТАВКИ ЗАДВИЖЕК";
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 5]].Value;
                        var sheetZdv = _fileIo.GetSheet("ЗДВ");
                        if (sheetZdv == null) throw new Exception("Лист \"ЗДВ\" не найден");
                        int startrowZdv = Program.Settings.startRow_ZDV;
                        var sheettwZdv = _fileIo.GetSheet("twZDV");
                        if (sheettwZdv == null) throw new Exception("Лист \"twZDV\" не найден");
                        dynamic rangeNameZdv = sheetZdv.Range[sheetZdv.Cells[4, 2], sheetZdv.Cells[_fileIo.LastRowCell(ref sheetZdv), 2]].Value;
                        adressIterator = 0;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int j = 1; j <= rangeNameZdv.GetLength(0); j++)
                        {
                            if (rangeNameZdv[j, 1] == null || rangeNameZdv[j, 1].ToString().Trim() == "")
                                continue; //пропускаем пропуски
                            for (int i = 1; i <= rangeTable.GetLength(0); i++)
                            {
                                if (rangeTable[i, 4] != null && rangeTable[i, 4].ToString() == "1")
                                {
                                    if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "")
                                        continue; //пропускаем пропуски
                                    rowObjects[1] = rangeNameZdv[j, 1].ToString() + ". " +
                                                    rangeTable[i, 2].ToString();
                                    //имя
                                    rowObjects[2] = rangeTable[i, 3]; //ID
                                    rowObjects[3] = "СЕК"; //ед изм
                                    rowObjects[4] = "0"; //уставка времени
                                    if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                                    {
                                        rowObjects[5] = 10;
                                    }
                                    else
                                    {
                                        rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                                    }
                                   
                                    //AdressBeginning
                                    rowObjects[6] = numPlace; //Place
                                    rowObjects[7] = "1"; //PLC

                                    if (Program.Settings.DB_version >= 9)
                                    {
                                        string stationTag =
                                            Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) +
                                            ".";
                                        //получим иерархию проекта
                                        rowObjects[8] = stationTag +
                                                        Convert.ToString(
                                                            sheetZdv.Cells[startrowZdv + j - 1, 204].Value) +
                                                        Convert.ToString(
                                                            _fileIo.GetSheet("twZDV").Cells[i + 1, 5].Value);
                                        rowObjects[9] = sysTimerName; //заполняем PLACENAME
                                    }

                                    adressIterator++;
                                    if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                        placeTimerDictionary.Add((int)rowObjects[6], sysTimerName);
                                    //кладем Place в словарь
                                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                }
                            }
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "twVS": //+prot
                        #region Вспомсистемы
                        sysTimerName = "ВРЕМЕННЫЕ УСТАВКИ ВСПОМСИСТЕМ";

                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 5]].Value;
                        var sheetVspom = _fileIo.GetSheet("ВСПОМ");
                        if (sheetVspom == null) throw new Exception("Лист \"ВСПОМ\" не найден");
                        int startrowVs = Program.Settings.startRow_VS;
                        dynamic rangeNameVspom = sheetVspom.Range[sheetVspom.Cells[3, 2], sheetVspom.Cells[_fileIo.LastRowCell(ref sheetVspom), 2]].Value;
                        adressIterator = 0;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int j = 1; j <= rangeNameVspom.GetLength(0); j++)
                        {
                            if (rangeNameVspom[j, 1] == null || rangeNameVspom[j, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                            for (int i = 1; i <= rangeTable.GetLength(0); i++)
                            {
                                if (rangeTable[i, 4] != null && rangeTable[i, 4].ToString() == "1")
                                {
                                    if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "")
                                        continue; //пропускаем пропуски
                                    rowObjects[1] = rangeNameVspom[j, 1].ToString() + ". " +
                                                    rangeTable[i, 2].ToString();
                                    //имя
                                    rowObjects[2] = rangeTable[i, 3]; //ID
                                    rowObjects[3] = "СЕК"; //ед изм
                                    rowObjects[4] = "0"; //уставка времени
                                    if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                                    {
                                        rowObjects[5] = 10;
                                    }
                                    else
                                    {
                                        rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                                    }
                                    
                                    //AdressBeginning
                                    rowObjects[6] = numPlace; //Place
                                    rowObjects[7] = "1"; //PLC
                                    if (Program.Settings.DB_version >= 9)
                                    {
                                        string stationTag =
                                            Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                        //получим иерархию проекта
                                        rowObjects[8] = stationTag +
                                                        Convert.ToString(
                                                            _fileIo.GetSheet("ВСПОМ").Cells[startrowVs + j - 1, 36]
                                                                .Value) +
                                                        Convert.ToString(
                                                            _fileIo.GetSheet("twVS").Cells[i + 1, 5].Value);
                                        rowObjects[9] = sysTimerName; //заполняем PLACENAME
                                    }
                                    adressIterator++;
                                    if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                        placeTimerDictionary.Add((int)rowObjects[6], sysTimerName);
                                    //кладем Place в словарь
                                    Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                }
                            }
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                    case "twASUPTzoneFoam": //+prot
                        #region Пожарные зоны
                        if (checkBox_PT.Checked == false) break;
                        sysTimerName = "ВРЕМЕННЫЕ УСТАВКИ ПОЖАРНЫХ ЗОН";
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 6]].Value;
                        var sheetPz = _fileIo.GetSheet("Зоны");
                        if (sheetPz == null) throw new Exception("Лист \"Зоны\" не найден");
                        dynamic rangeNamePz = sheetPz.Range[sheetPz.Cells[6, 2], sheetPz.Cells[_fileIo.LastRowCell(ref sheetPz), 4]].Value;
                        adressIterator = 0;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int j = 1; j <= rangeNamePz.GetLength(0); j++)
                        {

                            if (rangeNamePz[j, 1] == null || rangeNamePz[j, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                            if (rangeNamePz[j, 3].ToString() == "1") continue; //пропускаем отключенные зоны 
                            if (rangeNamePz[j, 1].ToString().ToUpper().Contains("БЕЗ"))
                            {
                                adressIterator = adressIterator + 5;
                                continue; //пропускаем зоны 
                            }
                            if (rangeNamePz[j, 1].ToString().ToUpper().Contains("ВОДО"))
                            {
                                for (int k = 0; k < 5; k++)
                                {
                                    pojZoneAddr[pojZoneAddrIndex] = Convert.ToInt32(adressTimer[index]) + adressIterator + k;
                                    pojZoneAddrIndex++;
                                }
                                adressIterator = adressIterator + 5;
                                continue; //пропускаем зоны 
                            }
                            for (int i = 1; i <= rangeTable.GetLength(0); i++)
                            {
                                if (rangeTable[i, 4] != null) // Столбец "Наличие" имеет номер 4.
                                {
                                    if (rangeTable[i, 4].ToString() == "1") // Столбец "Наличие" имеет номер 4.
                                    {
                                        if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "")
                                            continue; //пропускаем пропуски
                                        if (rangeTable[i, 2] == null || rangeTable[i, 2].ToString().Trim() == "0" ||
                                            rangeTable[i, 2].ToString().Trim() == "")
                                        {
                                            adressIterator++;
                                            continue; //пропускаем пропуски
                                        }
                                        rowObjects[1] = rangeNamePz[j, 1].ToString() + ". " +
                                                        rangeTable[i, 2].ToString();
                                        //имя
                                        rowObjects[2] = rangeTable[i, 3].ToString().ToUpper(); //ID
                                        rowObjects[3] = rangeTable[i, 5].ToString().ToUpper(); //ед изм
                                        rowObjects[4] = "0"; //уставка времени
                                        if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                                        {
                                            rowObjects[5] = 10;
                                        }
                                        else
                                        {
                                            rowObjects[5] = Convert.ToInt32(adressTimer[index]) + adressIterator;
                                        }
                                        
                                        //AdressBeginning
                                        rowObjects[6] = numPlace; //Place
                                        rowObjects[7] = "1"; //PLC
                                        if (Program.Settings.DB_version >= 9)
                                        {
                                            string stationTag =
                                                Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + ".";
                                            //получим иерархию проекта
                                            rowObjects[8] = stationTag +
                                                            Convert.ToString(
                                                                _fileIo.GetSheet("Зоны").Cells[j + 5, 204].Value) + "." +
                                                            rangeTable[i, 6];
                                            rowObjects[9] = sysTimerName; //заполняем PLACENAME
                                        }
                                        adressIterator++;
                                        if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                            placeTimerDictionary.Add((int)rowObjects[6], sysTimerName);
                                        //кладем Place в словарь
                                        Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                    }
                                }
                            }
                        }
                        numPlace = (int)rowObjects[6] + 1;

                        #endregion
                        break;
                    case "twASUPTzoneWater":
                        #region Пожарные зоны
                        if (checkBox_PT.Checked == false) break;
                        if (pojZoneAddr[0] == 0) break;
                        pojZoneAddrIndex = 0;
                        sysTimerName = "ВРЕМЕННЫЕ УСТАВКИ ПОЖАРНЫХ ЗОН";
                        rangeTable = sheetSys.Range[sheetSys.Cells[startRow, 1], sheetSys.Cells[endRow, 6]].Value;
                        sheetPz = _fileIo.GetSheet("Зоны");
                        if (sheetPz == null) throw new Exception("Лист \"Зоны\" не найден");
                        rangeNamePz = sheetPz.Range[sheetPz.Cells[6, 2], sheetPz.Cells[_fileIo.LastRowCell(ref sheetPz), 4]].Value;
                        rowObjects = new object[Data.DataSet.Tables[sysName].Columns.Count]; //массив представляющий строку
                        for (int j = 1; j <= rangeNamePz.GetLength(0); j++)
                        {

                            if (rangeNamePz[j, 1] == null || rangeNamePz[j, 1].ToString().Trim() == "") continue; //пропускаем пропуски
                            if (rangeNamePz[j, 1].ToString().ToUpper().Contains("БЕЗ")) continue; //пропускаем зоны 
                            if (rangeNamePz[j, 1].ToString().ToUpper().Contains("ПЕНО")) continue; //пропускаем зоны 
                            if (rangeNamePz[j, 3].ToString() == "1") continue; //пропускаем отключенные зоны 
                            for (int i = 1; i <= rangeTable.GetLength(0); i++)
                            {
                                if (rangeTable[i, 4] != null) // Столбец "Наличие" имеет номер 4.
                                {
                                    if (rangeTable[i, 4].ToString() == "1") // Столбец "Наличие" имеет номер 4.
                                    {
                                        if (rangeTable[i, 1] == null || rangeTable[i, 1].ToString().Trim() == "")
                                            continue; //пропускаем пропуски
                                        rowObjects[1] = rangeNamePz[j, 1].ToString() + ". " +
                                                        rangeTable[i, 2].ToString();
                                        //имя
                                        rowObjects[2] = rangeTable[i, 3].ToString().ToUpper(); //ID
                                        rowObjects[3] = rangeTable[i, 5].ToString().ToUpper(); //ед изм
                                        rowObjects[4] = "0"; //уставка времени
                                        if (_useAdrBegDenominator) //проверка на флаг AdressBeginning
                                        {
                                            rowObjects[5] = 10;
                                        }
                                        else
                                        {
                                            rowObjects[5] = pojZoneAddr[pojZoneAddrIndex].ToString(); //AdressBeginning
                                        }
                                       
                                        rowObjects[6] = numPlace - 1; //Place
                                        rowObjects[7] = "1"; //PLC
                                        if (Program.Settings.DB_version >= 9)
                                        {
                                            string stationTag = Convert.ToString(_fileIo.GetSheet("Иерархия").Cells[1, 2].Value) + "."; //получим иерархию проекта
                                            rowObjects[8] = stationTag + Convert.ToString(_fileIo.GetSheet("Зоны").Cells[j + 5, 204].Value) + "." + rangeTable[i, 6];
                                            rowObjects[9] = sysTimerName; //заполняем PLACENAME
                                        }
                                        pojZoneAddrIndex++;
                                        if (!placeTimerDictionary.ContainsKey((int)rowObjects[6]))
                                            placeTimerDictionary.Add((int)rowObjects[6], sysTimerName);
                                        //кладем Place в словарь
                                        Data.DataSet.Tables[sysName].Rows.Add(rowObjects); // добавляем строку
                                    }
                                }
                            }
                        }
                        numPlace = (int)rowObjects[6] + 1;
                        #endregion
                        break;
                }
            }
            if (Program.Settings.DB_version > 6)
            {
                foreach (DataRow row in Data.DataSet.Tables[sysName].Rows) //заполняем -1
                {
                    row[7] = 1;
                }

            }

            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                             Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
            //заполняем placetimer
            sysName = "PlaceTimer";
            var rowObjectsPlaceTimer = new object[Data.DataSet.Tables[sysName].Columns.Count];
            foreach (KeyValuePair<int, string> placeTimerKeyValuePair in placeTimerDictionary)
            {
                rowObjectsPlaceTimer[1] = placeTimerKeyValuePair.Key;
                rowObjectsPlaceTimer[2] = placeTimerKeyValuePair.Value;
                Data.DataSet.Tables[sysName].Rows.Add(rowObjectsPlaceTimer);
            }
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" " +
                             Data.DataSet.Tables[sysName].Rows.Count + @" параметров";
        }

        private bool CheckTimer()
        {
            ResetColorTable(dataGridView_Timer);
            ResetColorTable(dataGridView_PlaceTimer);
            if (checkBox_SAR.Checked) return false;
            return CheckTable(dataGridView_Timer) || CheckTable(dataGridView_PlaceTimer);
        }

        private void WriteTimer()
        {
            string sysName = "PlaceTimer";
            Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
            {
                foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                {
                    if (row[column].ToString() == "") row[column] = "NaN";
                }
                row[2] = row[2].ToString().ToUpper();


                _oleDb.WriteOperMess_PlaceTimer((uint)row.ItemArray[1], row.ItemArray[2].ToString());

            }
            Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";

            sysName = "Timer";
            Data.DataSet.Tables[sysName].AcceptChanges();
            //_oleDb = new OleDb();
            //_oleDb.Open(); //соединение с базой
            //_oleDb.TruncateTable(sysName);
            //_oleDb.Close();
            _oleDb = new OleDb();
            _oleDb.Open(); //соединение с базой
            foreach (DataRow row in Data.DataSet.Tables[sysName].Rows)
            {
                foreach (DataColumn column in Data.DataSet.Tables[sysName].Columns)
                {
                    if (row[column].ToString() == "") row[column] = "NaN";
                }
                row[1] = row[1].ToString().ToUpper();

                if (Program.Settings.DB_version < 7)
                {
                    _oleDb.WriteOperMess_Timer(row.ItemArray[1].ToString(), row.ItemArray[2].ToString(),
                        row.ItemArray[3].ToString(),
                        row.ItemArray[4].ToString(), (uint)row.ItemArray[5], (uint)row.ItemArray[6]);
                }
                else
                {
                    _oleDb.WriteOperMess_Timer(row.ItemArray[1].ToString(), row.ItemArray[2].ToString(),
                        row.ItemArray[3].ToString(),
                        row.ItemArray[4].ToString(), (uint)row.ItemArray[5], (uint)row.ItemArray[6], (uint)row.ItemArray[7]);
                }
            }
            Data.DataSet.Tables[sysName].RejectChanges();
            _oleDb.Close();
            toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + @" -> " + sysName + @" успешно";
        }

        #endregion Timer

        #region Редактор БД

        private void checkBox_bdEditor_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //button_ReadFromCSV.Enabled = checkBox_bdEditor.Checked;
                panel_typeMPSA.Enabled = !checkBox_bdEditor.Checked;
                button_GenerateFromExcel.Enabled = !checkBox_bdEditor.Checked;
                //button_Check.Enabled = !checkBox_bdEditor.Checked;
                button_WriteToDB.Enabled = !checkBox_bdEditor.Checked;
                tableLayoutPanel_Files.Enabled = !checkBox_bdEditor.Checked;
                настройкиToolStripMenuItem.Enabled = !checkBox_bdEditor.Checked;
                button_ExportToCSV.Enabled = checkBox_bdEditor.Checked;
                button_ClearBD.Enabled = !checkBox_bdEditor.Checked;
                checkBox_AutowriteSql.Visible = checkBox_bdEditor.Checked;

                if (!checkBox_bdEditor.Checked) button_AcceptedChanges.Visible = false;

                if (checkBox_bdEditor.Checked)
                {
                    if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                        !checkBox_Object.Checked && !checkBox_Timer.Checked)
                    {
                        toolStripStatusLabel1.Text = @"Ничего не выбрано";
                        return;
                    }
                    Data.DataSet = new DataSet();
                    _oleDb = new OleDb();
                    _oleDb.Open();

                    if (checkBox_OIP.Checked)
                    {
                        _daOip = new OleDbDataAdapter("Select * From OIP", _oleDb.Connection);
                        _daOip.FillSchema(Data.DataSet, SchemaType.Source, "OIP");
                        _daOip.Fill(Data.DataSet, "OIP");

                        if (Program.Settings.DB_version > 8)
                        {
                            _daOipType = new OleDbDataAdapter("Select * From OIPTYPE", _oleDb.Connection);
                            _daOipType.FillSchema(Data.DataSet, SchemaType.Source, "OIPTYPE");
                            _daOipType.Fill(Data.DataSet, "OIPTYPE"); 
                       }
                       
                        _daPlaceOip = new OleDbDataAdapter("Select * From PlaceOIP", _oleDb.Connection);
                        _daPlaceOip.FillSchema(Data.DataSet, SchemaType.Source, "PlaceOIP");
                        _daPlaceOip.Fill(Data.DataSet, "PlaceOIP");
                    }
                    if (checkBox_SysObject.Checked)
                    {
                        _daSysObject = new OleDbDataAdapter("Select * From SysObject", _oleDb.Connection);
                        _daSysObject.FillSchema(Data.DataSet, SchemaType.Source, "SysObject");
                        _daSysObject.Fill(Data.DataSet, "SysObject");
                    }
                    if (checkBox_Object.Checked)
                    {
                        _daObject = new OleDbDataAdapter("Select * From Object", _oleDb.Connection);
                        _daObject.FillSchema(Data.DataSet, SchemaType.Source, "Object");
                        _daObject.Fill(Data.DataSet, "Object");
                    }
                    if (checkBox_Message.Checked)
                    {
                        _daMessage = new OleDbDataAdapter("Select * From Message", _oleDb.Connection);
                        _daMessage.FillSchema(Data.DataSet, SchemaType.Source, "Message");
                        _daMessage.Fill(Data.DataSet, "Message");
                    }
                    if (checkBox_Timer.Checked)
                    {
                        _daTimer = new OleDbDataAdapter("Select * From Timer", _oleDb.Connection);
                        _daTimer.FillSchema(Data.DataSet, SchemaType.Source, "Timer");
                        _daTimer.Fill(Data.DataSet, "Timer");

                        _daPlaceTimer = new OleDbDataAdapter("Select * From PlaceTimer", _oleDb.Connection);
                        _daPlaceTimer.FillSchema(Data.DataSet, SchemaType.Source, "PlaceTimer");
                        _daPlaceTimer.Fill(Data.DataSet, "PlaceTimer");
                    }

                    GetCommandUpdate();

                    foreach (DataTable table in Data.DataSet.Tables)
                    {
                        table.RowChanged += Table_RowChanged;
                        table.RowDeleted += Table_RowDeleted;
                    }

                    DataGridViewInicialization();
                    toolStripStatusLabel1.Text = @"Соедиенение с " + _oleDb.DataSource + @"-->" + _oleDb.InitialCatalog +
                                                 @" установлено. Загружен режим редактора";
                }
                else
                {
                    button_WriteToDB.Enabled = false;
                    Data.ClearDataSet();
                    if (_oleDb != null && _oleDb.SqlConnected()) _oleDb.Close();
                }
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка режима редактора " + exception.Message;
                if (_oleDb != null) toolStripStatusLabel1.Text += @" SQL: " + _oleDb.MessErr;
            }
            finally
            {
                if (_oleDb != null && _oleDb.SqlConnected()) _oleDb.Close();
            }
        }

        private void GetCommandUpdate()
        {
            try
            {
                if (checkBox_OIP.Checked)
                {
                    OleDbCommandBuilder builderOip = new OleDbCommandBuilder(_daOip);
                    _daOip.UpdateCommand = builderOip.GetUpdateCommand();

                    if (Program.Settings.DB_version > 8)
                    {
                        OleDbCommandBuilder builderOipType = new OleDbCommandBuilder(_daOipType);
                        _daOipType.UpdateCommand = builderOipType.GetUpdateCommand();
                    }

                    OleDbCommandBuilder builderPlaceOip = new OleDbCommandBuilder(_daPlaceOip);
                    _daPlaceOip.UpdateCommand = builderPlaceOip.GetUpdateCommand();
                }
                if (checkBox_SysObject.Checked)
                {
                    _daSysObject.UpdateCommand = new OleDbCommandBuilder(_daSysObject).GetUpdateCommand();
                }
                if (checkBox_Object.Checked)
                {
                    _daObject.UpdateCommand = new OleDbCommandBuilder(_daObject).GetUpdateCommand();
                }
                if (checkBox_Message.Checked)
                {
                    _daMessage.UpdateCommand = new OleDbCommandBuilder(_daMessage).GetUpdateCommand();
                }
                if (checkBox_Timer.Checked)
                {
                    _daTimer.UpdateCommand = new OleDbCommandBuilder(_daTimer).GetUpdateCommand();
                    _daPlaceTimer.UpdateCommand = new OleDbCommandBuilder(_daPlaceTimer).GetUpdateCommand();
                }
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка редактора. В базе данных у таблицы не найден Primary Key. " + _oleDb.MessErr + exception.Message;
            }
        }

        private void Table_RowChanged(object sender, DataRowChangeEventArgs e)
        {
            button_AcceptedChanges.Visible = true;
        }
        private void Table_RowDeleted(object sender, DataRowChangeEventArgs e)
        {
            button_AcceptedChanges.Visible = true;
        }
        private void button_AcceptedChanges_Click(object sender, EventArgs e)
        {
            try
            {
                Stopwatch watch = new Stopwatch();
                watch.Start();
                DataGridViewDataSourceClear();
                Cursor = Cursors.WaitCursor;
                if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                    !checkBox_Object.Checked && !checkBox_Timer.Checked)
                {
                    toolStripStatusLabel1.Text = @"Ничего не выбрано";
                    return;
                }
                toolStripStatusLabel1.Text = @" Запись в БД (";
                if (checkBox_OIP.Checked)
                {
                    _daPlaceOip.Update(Data.DataSet, "PlaceOIP");
                    toolStripStatusLabel1.Text += @" PlaceOIP ";
                    _daOip.Update(Data.DataSet, "OIP");
                    toolStripStatusLabel1.Text += @" OIP ";
                    if (Program.Settings.DB_version > 8)
                    {
                        _daOipType.Update(Data.DataSet, "OIPTYPE");
                        toolStripStatusLabel1.Text += @" OIPTYPE";
                    }
                }
                if (checkBox_SysObject.Checked)
                {
                    _daSysObject.Update(Data.DataSet, "SysObject");
                    toolStripStatusLabel1.Text += @" SysObject ";
                }

                if (checkBox_Object.Checked)
                {
                    _daObject.Update(Data.DataSet, "Object");
                    toolStripStatusLabel1.Text += @" Object ";
                }

                if (checkBox_Message.Checked)
                {
                    _daMessage.Update(Data.DataSet, "Message");
                    toolStripStatusLabel1.Text += @" Message ";
                }

                if (checkBox_Timer.Checked)
                {
                    _daPlaceTimer.Update(Data.DataSet, "PlaceTimer");
                    toolStripStatusLabel1.Text += @" PlaceTimer ";
                    _daTimer.Update(Data.DataSet, "Timer");
                    toolStripStatusLabel1.Text += @" Timer ";
                }
                Cursor = Cursors.Default;
                button_AcceptedChanges.Visible = false;
                DataGridViewDataSourceSet();
                if (checkBox_AutowriteSql.Checked)
                {
                    button_WriteSQLCommands_Click(null, null);
                    checkBox_bdEditor_CheckedChanged(null, null);
                    toolStripStatusLabel1.Text = "Пользовательский SQL:" + richTextBox_SQL_Results.Text + @" + Запись(автоматическая) в БД (успешно";
                }
                watch.Stop();
                toolStripStatusLabel1.Text += @") завершена [прошло " + watch.ElapsedMilliseconds + @"мс]";
            }
            catch (Exception exception)
            {
                DataGridViewDataSourceSet();
                Cursor = Cursors.Default;
                toolStripStatusLabel1.Text = @"Ошибка обновления БД " + _oleDb.MessErr + exception.Message;
                if (_oleDb.MessErr != null) toolStripStatusLabel1.Text += @" SQL: " + _oleDb.MessErr;
            }
        }

        #endregion

        #region Чтение CSV

        private void button_ReadFromCSV_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Stopwatch watch = new Stopwatch();
                watch.Start();
                if (!checkBox_OIP.Checked && !checkBox_SysObject.Checked && !checkBox_Message.Checked &&
                    !checkBox_Object.Checked && !checkBox_Timer.Checked)
                {
                    toolStripStatusLabel1.Text = @"Ничего не выбрано";
                    return;
                }

                button_Check.Enabled = true;
                button_ExportToCSV.Enabled = true;
                button_WriteToDB.Enabled = true;
                button_Check.BackColor = DefaultBackColor;
                button_Check.UseVisualStyleBackColor = true;
                if (!checkBox_bdEditor.Checked)
                {
                    Data.IninicializationDataSet();
                    DataGridViewInicialization();
                    Refresh();
                    toolStripStatusLabel1.Text = @"Чтение CSV";
                }
                else
                {
                   toolStripStatusLabel1.Text += @" --> добавлены данные из файлов CSV"; 
                }
                
                if (checkBox_OIP.Checked)
                {
                    ReadCsv("OIP");
                    ReadCsv("PlaceOIP");
                    if (Program.Settings.DB_version > 8) ReadCsv("OIPType");
                }
                if (checkBox_SysObject.Checked) ReadCsv("SysObject");
                if (checkBox_Message.Checked) ReadCsv("Message");
                if (checkBox_Object.Checked) ReadCsv("Object");
                if (checkBox_Timer.Checked)
                {
                    ReadCsv("Timer");
                    ReadCsv("PlaceTimer");
                }

                watch.Stop();
                toolStripStatusLabel1.Text += @" [прошло " + watch.ElapsedMilliseconds + @"мс]";
                Cursor = Cursors.Default;
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка чтения CSV [" + exception.Source + @"]: " + exception.Message;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        void ReadCsv(string sysName)
        {
            string pathimportDirectory = _pathimportDirectory;
            if (!Directory.Exists(pathimportDirectory)) Directory.CreateDirectory(pathimportDirectory);
            var reader = new StreamReader(File.OpenRead(pathimportDirectory + "\\" + sysName + ".csv"), Encoding.Default);
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (line == null) continue;
                string[] values = line.Split(';');
                if (values.Length == 0) continue;
                if (values[0].ToUpper().Contains("ID")) continue; //пропускаем первую строку
                for (int index = 0; index < values.Length; index++)
                {
                    values[index] = values[index].Replace("\"", "");
                }
                DataRow row = Data.DataSet.Tables[sysName].NewRow();
                try
                {
                    row.ItemArray = values;
                }
                catch (Exception)
                {
                    throw new Exception("Количество столбцов или тип данных не совпадают в " + sysName + " строка c ID " + values[0]);
                }
                Data.DataSet.Tables[sysName].Rows.Add(row);
            }
        }


        #endregion

        #region Интерфейс

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.Show(this);
        }

        private void button_Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_FileDialog_MSG_Click(object sender, EventArgs e)
        {
            openFileDialog_MSG.InitialDirectory = Application.StartupPath;
            openFileDialog_MSG.Multiselect = false;
            openFileDialog_MSG.Filter = @"Excel|*.xls;*.xlsx;*.xlsm";
            openFileDialog_MSG.ShowDialog(this);
        }

        private void button_FileDialog_IO_Click(object sender, EventArgs e)
        {
            openFileDialog_IO.InitialDirectory = Application.StartupPath;
            openFileDialog_IO.Multiselect = false;
            openFileDialog_IO.Filter = @"Excel|*.xls;*.xlsx;*.xlsm";
            openFileDialog_IO.ShowDialog(this);
        }

        private void openFileDialog_IO_FileOk(object sender, CancelEventArgs e)
        {
            textBox_path_IO.Text = openFileDialog_IO.FileName;
        }

        private void openFileDialog_MSG_FileOk(object sender, CancelEventArgs e)
        {
            textBox_path_MSG.Text = openFileDialog_MSG.FileName;
        }

        private void checkBox_Message_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Message.Checked) checkBox_SysObject.Checked = checkBox_Message.Checked;
            if (checkBox_Object.Checked || checkBox_Timer.Checked || checkBox_Message.Checked)
            {
                panel_typeMPSA.Enabled = true;
            }
            else
            {
                panel_typeMPSA.Enabled = false;
            }
        }

        private void checkBox_Object_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Object.Checked || checkBox_Timer.Checked || checkBox_Message.Checked)
            {
                panel_typeMPSA.Enabled = true;
            }
            else
            {
                panel_typeMPSA.Enabled = false;
            }
        }

        private void checkBox_Timer_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Object.Checked || checkBox_Timer.Checked || checkBox_Message.Checked)
            {
                panel_typeMPSA.Enabled = true;
            }
            else
            {
                panel_typeMPSA.Enabled = false;
            }
        }

        private void checkBox_SysObject_CheckedChanged(object sender, EventArgs e)
        {
            //if (!checkBox_SysObject.Checked)
            //{
            //    if (checkBox_Message.Checked || checkBox_Object.Checked)
            //        checkBox_SysObject.Checked = true;
            //}
        }

        private void checkBox_MNS_PNS_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_MNS_PNS.Checked && !checkBox_All.Checked)
            {
                //checkBox_SAR.Checked = !checkBox_MNS_PNS.Checked;
                checkBox_PT.Checked = !checkBox_MNS_PNS.Checked;
            }
            CheckForNullChoice();
        }

        private void CheckForNullChoice()
        {
            if (!checkBox_MNS_PNS.Checked && !checkBox_SAR.Checked && !checkBox_RP.Checked && !checkBox_PT.Checked)
                checkBox_MNS_PNS.Checked = true;
        }

        private void checkBox_SAR_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_SAR.Checked && !checkBox_All.Checked)
            {
                //checkBox_MNS_PNS.Checked = !checkBox_SAR.Checked;
                checkBox_RP.Checked = !checkBox_SAR.Checked;
                checkBox_PT.Checked = !checkBox_SAR.Checked;
            }
            CheckForNullChoice();
        }

        private void checkBox_RP_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_RP.Checked && !checkBox_All.Checked)
            {
                checkBox_SAR.Checked = !checkBox_RP.Checked;
                checkBox_PT.Checked = !checkBox_RP.Checked;
                checkBox_MNS_PNS.Checked = checkBox_RP.Checked;
            }
            CheckForNullChoice();
        }

        private void checkBox_PT_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_PT.Checked && !checkBox_All.Checked)
            {
                checkBox_MNS_PNS.Checked = !checkBox_PT.Checked;
                checkBox_RP.Checked = !checkBox_PT.Checked;
                checkBox_SAR.Checked = !checkBox_PT.Checked;
            }
            CheckForNullChoice();
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SettingsForm settingsForm = new SettingsForm(this);
            settingsForm.Show(this);
        }

        private void dataGridView_Message_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            ColorRow();
        }

        private void toolStripMenuItem_Delete_Click(object sender, EventArgs e)
        {
            TableWork(TableOperations.Delete);
        }

        private void ToolStripMenuItem_Insert_Click(object sender, EventArgs e)
        {
            TableWork(TableOperations.Insert);
        }

        private void toolStripMenuItem_Copy_Click(object sender, EventArgs e)
        {
            TableWork(TableOperations.Copy);
        }

        private void toolStripMenuItem_Paste_Click(object sender, EventArgs e)
        {
            TableWork(TableOperations.Paste);
        }

        private void TableWork(TableOperations operations)
        {
            try
            {
                switch (tabControl_Main.TabPages.IndexOf(tabControl_Main.SelectedTab))
                {
                    case 0:
                        foreach (DataGridViewRow row in dataGridView_OIP.SelectedRows)
                        {
                                    TableWorkOperations("OIP", operations, row);
                            }
                        break;
                    case 1:
                        foreach (DataGridViewRow row in dataGridView_PlaceOIP.SelectedRows)
                        {
                                    TableWorkOperations("PlaceOIP", operations, row);
                            }
                        break;
                    case 2:
                        foreach (DataGridViewRow row in dataGridView_SysObject.SelectedRows)
                        {
                                    TableWorkOperations("SysObject", operations, row);
                            }
                        break;
                    case 3:
                        foreach (DataGridViewRow row in dataGridView_Message.SelectedRows)
                        {
                                    TableWorkOperations("Message", operations, row);
                            }
                        break;
                    case 4:
                        foreach (DataGridViewRow row in dataGridView_Object.SelectedRows)
                        {
                                    TableWorkOperations("Object", operations, row);
                            }
                        break;
                    case 5:
                        foreach (DataGridViewRow row in dataGridView_Timer.SelectedRows)
                        {
                                    TableWorkOperations("Timer", operations, row);
                            }
                        break;
                    case 6:
                        foreach (DataGridViewRow row in dataGridView_PlaceTimer.SelectedRows)
                        {
                            TableWorkOperations("PlaceTimer", operations, row);
                        }
                        break;
                    case 8:
                        if (Program.Settings.DB_version > 8)
                        {
                            foreach (DataGridViewRow row in dataGridView_OIPType.SelectedRows)
                            {
                                TableWorkOperations("OIPTYPE", operations, row);
                            }
                        }
                        break;
                }
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка операции " + exception.Source + @" : " + exception.Message; 
            }
        }

        private static void TableWorkOperations(string tableName, TableOperations operations, DataGridViewRow row)
        {
            switch (operations)
            {
                case TableOperations.Delete:
                    Data.DataSet.Tables[tableName].Rows[row.Index].Delete();
                    break;
                case TableOperations.Insert:
                    DataRow dataRow = Data.DataSet.Tables[tableName].NewRow();
                    for (int i = 1; i < dataRow.ItemArray.Length; i++)
                    {
                        dataRow[i] = 0;
                    }
                    Data.DataSet.Tables[tableName].Rows.Add(dataRow);
                    break;
                case TableOperations.Copy:
                    StringBuilder sbldr = new StringBuilder();
                    foreach (DataGridViewCell item in row.Cells)
                    {
                        sbldr.Append(item.Value.ToString() + ';');
                    }
                    Clipboard.SetText(sbldr.ToString());
                    break;
                case TableOperations.Paste:
                    dataRow = Data.DataSet.Tables[tableName].NewRow();
                    string[] dataStr = Clipboard.GetText().Split(Convert.ToChar(";"));
                    for (int index = 1; index < row.Cells.Count; index++)
                    {
                        dataRow[index] = dataStr[index - 1];
                    }
                    Data.DataSet.Tables[tableName].Rows.Add(dataRow);
                    break;
            }
        }

        void ColorRow()
        {
            foreach (DataGridViewRow row in dataGridView_Message.Rows)
            {
                int value = Convert.ToInt32(row.Cells[10].Value);
                switch (value)
                {
                    case 4:
                        break;
                    case 1:
                        row.DefaultCellStyle.BackColor = Color.ForestGreen;
                        break;
                    case 3:
                        row.DefaultCellStyle.BackColor = Color.Red;
                        break;
                    case 2:
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        break;
                    case 5:
                        row.DefaultCellStyle.BackColor = Color.LightBlue;
                        break;
                    default:
                        row.DefaultCellStyle.BackColor = Color.Gray;
                        toolStripStatusLabel1.Text += @"В таблице Message(строка" + row.Index + @") неверный цвет!";
                        break;
                }
            }
        }

        private void label3_MouseMove(object sender, MouseEventArgs e)
        {
            label3.BorderStyle = BorderStyle.FixedSingle;
        }

        private void label3_MouseLeave(object sender, EventArgs e)
        {
            label3.BorderStyle = BorderStyle.None;
        }

        private void label3_DoubleClick(object sender, EventArgs e)
        {
            checkBox_OIP.Checked = !checkBox_OIP.Checked;
            checkBox_SysObject.Checked = !checkBox_SysObject.Checked;
            checkBox_Message.Checked = !checkBox_Message.Checked;
            checkBox_Object.Checked = !checkBox_Object.Checked;
            checkBox_Timer.Checked = !checkBox_Timer.Checked;
        }

        private void checkBox_All_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_All.Checked)
            {
                checkBox_MNS_PNS.Checked = true;
                checkBox_SIKN.Checked = true;
                checkBox_PT.Checked = true;
                checkBox_RP.Checked = true;
                checkBox_SAR.Checked = true;
            }
            else
            {
                checkBox_MNS_PNS.Checked = true;
                checkBox_SIKN.Checked = false;
                checkBox_PT.Checked = false;
                checkBox_RP.Checked = false;
                checkBox_SAR.Checked = false;
            }
            
        }
        private void button_WriteSQLCommands_Click(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox_SQL_commands.Lines.Length == 0)
                {
                    richTextBox_SQL_Results.Text = "Нет команд";
                    return;
                }
                _sqlDb = new SqlDb();

                _sqlDb.Open();
                _sqlDb.CreateDbCommand(richTextBox_SQL_commands.Lines);
                _sqlDb.ExecuteReader();

                _sqlDb.Reader.Read();
                // button_CheckDBversion.Text = Convert.ToString(_oleDb.Reader["Build"]);
                _sqlDb.Close();
                if (_sqlDb.MessArr == null) richTextBox_SQL_Results.Text = "Успешное выполнение";
                else richTextBox_SQL_Results.Lines = _sqlDb.MessArr;
            }
            catch (Exception exception)
            {
                toolStripStatusLabel1.Text = @"Ошибка выполнения SQL команд пользователя " + exception.Message;
                if (_sqlDb != null) toolStripStatusLabel1.Text += @" : " + _sqlDb.MessErr;
                richTextBox_SQL_Results.Text += Environment.NewLine + toolStripStatusLabel1.Text;
            }
            finally
            {
                if (_sqlDb != null && _sqlDb.SqlConnected()) _sqlDb.Close();
            }
        }
        #endregion



    }
}