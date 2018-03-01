using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DataGenerator
{
    /// <summary>
    /// "Библиотека вызова функций MS Excel" (Excel.dll)
    /// </summary>
    public class Excel
    {
        #region Константы и переменные
        public Application App;
        public Workbook Wbook;
        public Worksheet Wsheet;
        private static readonly object MissingObj = System.Reflection.Missing.Value;
        #endregion

        #region Конструктор
        
        // Создать экземпляр из шаблона
        public Excel()
        {
            Open();
        }

        // Создать экземпляр
        public void Open()
        {
            App = null;
            Wbook = null;
            Wsheet = null;
            try
            {
                App = new Application();
                Wbook = App.Workbooks.Add(MissingObj);
                Wsheet = (Worksheet)Wbook.Worksheets.Item[1];
            }
            catch (Exception)
            {
                App = null;
                Wbook = null;
                Wsheet = null;
            }
        }

        // Создать экземпляр из шаблона
        public Excel(string pathToTemplate, bool updateLinks, bool readOnly)
        {
            Open(pathToTemplate, updateLinks, readOnly);
        }

        // Создать экземпляр из шаблона
        public void Open(string pathToTemplate, bool updateLinks, bool readOnly)
        {
            App = null;
            Wbook = null;
            Wsheet = null;
            try
            {
                //object pathToTemplateObj = pathToTemplate;
                App = new Application();
                //Wbook = App.Workbooks.Add(pathToTemplateObj);
                //Wsheet = (Worksheet)Wbook.Worksheets.Item[1];
                Wbook = App.Workbooks.Open(pathToTemplate, updateLinks, readOnly);
                Wsheet = (Worksheet)Wbook.ActiveSheet;
            }
            catch (Exception)
            {
                App = null;
                Wbook = null;
                Wsheet = null;
            }
        }

        #endregion

        public bool ChangeBook(string bookName)
        {
            try
            {
                foreach (var book in App.Workbooks)
                {
                    var workbook = book as Workbook;
                    if (workbook != null)
                    {
                        if (workbook.Name == bookName)
                        {
                            workbook.Activate();
                        }
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Закрыть экземпляр
        /// </summary>
        public void Close()
        {
            try
            {
                Wbook.Close(false, MissingObj, MissingObj);
                App.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
                Wsheet = null;
                Wbook = null;
                App = null;
            }
            catch (Exception)
            {
                Wsheet = null;
                Wbook = null;
                App = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        // Сохранить
        public bool Save(string folder, string file)
        {
            ExceptionMsg = null;
            try
            {
                bool alerts = App.DisplayAlerts;
                App.DisplayAlerts = false;
                 Wbook.SaveAs(folder+file);
                App.DisplayAlerts = alerts;
                return true;
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
            return false;
        }

        public string ExceptionMsg { get; set; }

        /// <summary>
        /// Видимость
        /// </summary>
        public bool Visible
        {
            get
            {
                return App.Visible;
            }
            set
            {
                App.Visible = value;
            }
        }

        /// <summary>
        /// Получить список листов
        /// </summary>
        /// <returns></returns>
        public string[] GetSheetsNames()
        {
            try
            {
                var sheets = new string[Wbook.Sheets.Count];
                var i = 0;
                foreach (var sheet in Wbook.Sheets)
                {
                    var worksheet = sheet as Worksheet;
                    if (worksheet != null)
                    {
                        sheets[i] = worksheet.Name;
                        i++;
                    }
                }
                return sheets;
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return null;
            }
        }

        /// <summary>
        /// Получить лист
        /// </summary>
        /// <param name="name">Имя</param>
        /// <returns></returns>
        public Worksheet GetSheet(string name)
        {
            try
            {
                Worksheet ws;
                ws = Wbook.Sheets.OfType<Worksheet>().FirstOrDefault(worksheet => worksheet.Name.ToLower() == name.ToLower());
                if (ws != null)
                {
                    ws.AutoFilterMode = false;
                }
                return ws;
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return null;
            }
        }

        // Поиск номера столбца по значению в ячейке строки
        public Int64 GetColumnNumberByField(ref Worksheet wsheet, Int64 row, string textField, bool bTrimSpace = false)
        {
            try
            {
                if ((wsheet != null) && (row>0))
                {
                    // обрезать пробелы справа слева и поднять регистр
                    var textFieldTrimAndUpper = bTrimSpace ? textField.ToUpper().TrimEnd(' ').TrimStart(' ') : textField.ToUpper();
                    for (var num = 1; num < wsheet.UsedRange.Columns.Count; num++)
                    {
                        if (!string.IsNullOrEmpty(wsheet.Cells[row, num].Value) && !string.IsNullOrEmpty(textFieldTrimAndUpper))
                        {
                            var textFieldTrimAndUpperExcel = bTrimSpace
                                ? wsheet.Cells[row, num].Value.ToUpper().TrimEnd(' ').TrimStart(' ')
                                : wsheet.Cells[row, num].Value.ToUpper();
                            if (textFieldTrimAndUpperExcel == textFieldTrimAndUpper)
                            {
                                return num;
                            }
                        }
                    }
                }
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        // Поиск номера строки по значению в ячейке столбца
        public Int64 GetRowNumberByField(ref Worksheet wsheet, Int64 col, string textField, bool bTrimSpace = false)
        {
            try
            {
                if ((wsheet != null) && (col > 0))
                {
                    // обрезать пробелы справа слева и поднять регистр
                    var textFieldTrimAndUpper = bTrimSpace ? textField.ToUpper().TrimEnd(' ').TrimStart(' ') : textField.ToUpper();
                    for (var num = 1; num < wsheet.UsedRange.Rows.Count; num++)
                    {
                        if (!string.IsNullOrEmpty(wsheet.Cells[num, col].Value) && !string.IsNullOrEmpty(textFieldTrimAndUpper))
                        {
                            var textFieldTrimAndUpperExcel = bTrimSpace 
                                ? wsheet.Cells[num, col].Value.ToUpper().TrimEnd(' ').TrimStart(' ') 
                                : wsheet.Cells[num, col].Value.ToUpper();
                            if (textFieldTrimAndUpperExcel == textFieldTrimAndUpper)
                            {
                                return num;
                            }
                        }
                    }
                }
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }


        /// <summary>
        /// Получить значение ячейки в виде строки
        /// </summary>
        /// <param name="wsheet">Лист</param>
        /// <param name="row">строка</param>
        /// <param name="col">столбец</param>
        /// <returns></returns>
        public string GetCellValueAsString(ref Worksheet wsheet, Int64 row, Int64 col)
        {
            try
            {
                return Convert.ToString(wsheet.Cells[row, col].Value);
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return null;
            }
        }

        public string GetCellValueAsString(ref Worksheet wsheet, decimal row, decimal col)
        {
            try
            {
                return Convert.ToString(wsheet.Cells[row, col].Value);
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return null;
            }
        }

        public object GetCellValueAsObject(ref Worksheet wsheet, decimal row, decimal col)
        {
            try
            {
                return wsheet.Cells[row, col].Value;
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return null;
            }
        }

        // Получить значение ячейки в виде числа Int32
        public Int32 GetCellValueAsInt32(ref Worksheet wsheet, Int64 row, Int64 col)
        {
            try
            {
                if (string.IsNullOrEmpty(Convert.ToString(wsheet.Cells[row, col].Value)))
                    return 0;
                else
                    return Convert.ToInt32(wsheet.Cells[row, col].Value);
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return -1;
            }
        }

        // Получить значение ячейки в виде числа Int64
        public Int64 GetCellValueAsInt64(ref Worksheet wsheet, Int64 row, Int64 col)
        {
            try
            {
                if (string.IsNullOrEmpty(Convert.ToString(wsheet.Cells[row, col].Value)))
                    return 0;
                else
                    return Convert.ToInt64(wsheet.Cells[row, col].Value);
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
                return -1;
            }
        }

        // Вставка значения в ячейку
        public void SetCellValue(Int64 row, Int64 col, ref string cellValue)
        {
            try
            {
                if (cellValue == null) cellValue = "";
                Wsheet.Cells[row, col].NumberFormat = "@";
                Wsheet.Cells[row, col] = cellValue;
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Вставка значений в строку
        public void SetRowValue(Int64 row, Int64 colStart, Int64 colEnd, ref string[] values)
        {
            try
            {
                if (values != null)
                {
                    Wsheet.Range[Wsheet.Cells[row, colStart], Wsheet.Cells[row, colEnd]].NumberFormat = "@";
                    Wsheet.Range[Wsheet.Cells[row, colStart], Wsheet.Cells[row, colEnd]].Value = values;
                }
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Вставка значений в диапазон
        public void SetRangeValue(Int64 rowStart, Int64 colStart, Int64 rowEnd, Int64 colIndexEnd, ref string[,] range)
        {
            try
            {
                if (range != null)
                {
                    Wsheet.Range[Wsheet.Cells[rowStart, colStart], Wsheet.Cells[rowEnd, colIndexEnd]].NumberFormat = "@";
                    Wsheet.Range[Wsheet.Cells[rowStart, colStart], Wsheet.Cells[rowEnd, colIndexEnd]].Value = range;
                }
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Ширина столбца
        public void SetColumnWidth(Int64 col, int colWidth)
        {
            ((Range)Wsheet.Columns[col, Type.Missing]).EntireColumn.ColumnWidth = colWidth;
        }

        // Автоширина столбца
        public void SetColumnAutoWidth(Int64 col)
        {
            try
            {
                ((Range) Wsheet.Columns[col, Type.Missing]).AutoFit();
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Фиксирование заголовка
        public void SetHeaderFixed(bool fix)
        {
            try
            {
                if (fix)
                {
                    App.ActiveWindow.SplitColumn = 0;
                    App.ActiveWindow.SplitRow = 1;
                    App.ActiveWindow.FreezePanes = true;
                }
                else
                {
                    App.ActiveWindow.FreezePanes = false;
                }
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Выравнивание диапазона
        public void SetRangeAlign(Int64 rowStart, Int64 colStart, Int64 rowCount, Int64 colCount)
        {
            try
            {
                var range = Wsheet.Range[Wsheet.Cells[rowStart, colStart], Wsheet.Cells[rowCount, colCount]];
                range.HorizontalAlignment = Constants.xlCenter;
                range.VerticalAlignment = Constants.xlCenter;
                range.WrapText = true;
                range.Orientation = 0;
                range.AddIndent = false;
                //range.IndentLevel = 0;
                range.ShrinkToFit = false;
                range.MergeCells = false;
                range.ReadingOrder = Convert.ToInt32(Constants.xlContext);
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Выравнивание диапазона
        public void SetColumnsAlign(Int64 colStart, Int64 colCount)
        {
            try
            {
                for (var i = colStart; i <= colCount; i++)
                {
                    var range = ((Range) Wsheet.Columns[i, Type.Missing]);
                    range.HorizontalAlignment = Constants.xlCenter;
                    range.VerticalAlignment = Constants.xlCenter;
                    range.WrapText = true;
                    range.Orientation = 0;
                    range.AddIndent = false;
                    //range.IndentLevel = 0;
                    range.ShrinkToFit = false;
                    range.MergeCells = false;
                    range.ReadingOrder = Convert.ToInt32(Constants.xlContext);
                }
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Форматирование диапазона
        public void SetRangeFormat(Int64 rowStart, Int64 colStart, Int64 rowCount, Int64 colCount, string setFormat, int rowHeight = 0, int colWidth = 0)
        {
            try
            {

                if ((rowHeight != 0) && (colWidth != 0))
                {
                    var range =
                        Wsheet.Range[Wsheet.Cells[rowStart, colStart], Wsheet.Cells[rowCount, colCount]];
                    // стиль
                    switch (setFormat.ToLower())
                    {
                        case "xlrangeautoformataccounting1":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatAccounting1);
                            break;
                        case "xlrangeautoformataccounting2":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatAccounting2);
                            break;
                        case "xlrangeautoformataccounting3":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatAccounting3);
                            break;
                        case "xlrangeautoformataccounting4":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatAccounting4);
                            break;
                        case "xlrangeautoformatclassic1":
                            range.AutoFormat();
                            break;
                        case "xlrangeautoformatclassic2":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                            break;
                        case "xlrangeautoformatclassic3":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic3);
                            break;
                        case "xlrangeautoformatlist1":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatList1);
                            break;
                        case "xlrangeautoformatlist2":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatList2);
                            break;
                        case "xlrangeautoformatlist3":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatList3);
                            break;
                        case "xlrangeautoformatreport1":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport1);
                            break;
                        case "xlrangeautoformatreport2":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport2);
                            break;
                        case "xlrangeautoformatreport3":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport3);
                            break;
                        case "xlrangeautoformatreport4":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport4);
                            break;
                        case "xlrangeautoformatreport5":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport5);
                            break;
                        case "xlrangeautoformatreport6":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport6);
                            break;
                        case "xlrangeautoformatreport7":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport7);
                            break;
                        case "xlrangeautoformatreport8":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport8);
                            break;
                        case "xlrangeautoformatreport9":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport9);
                            break;
                        case "xlrangeautoformatreport10":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatReport10);
                            break;
                        case "xlrangeautoformattable1":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable1);
                            break;
                        case "xlrangeautoformattable2":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable2);
                            break;
                        case "xlrangeautoformattable3":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable3);
                            break;
                        case "xlrangeautoformattable4":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable4);
                            break;
                        case "xlrangeautoformattable5":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable5);
                            break;
                        case "xlrangeautoformattable6":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable6);
                            break;
                        case "xlrangeautoformattable7":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable7);
                            break;
                        case "xlrangeautoformattable8":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable8);
                            break;
                        case "xlrangeautoformattable9":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable9);
                            break;
                        case "xlrangeautoformattable10":
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatTable10);
                            break;
                        case "без форматирования":
                            break;
                        default:
                            range.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatNone);
                            break;
                    }
                    // высота строк
                    range.RowHeight = rowHeight;
                    // ширина столбцов
                    range.ColumnWidth = colWidth;
                }
                //range.Borders(xlDiagonalDown).LineStyle = xlNone
                //Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                //range.Select();
                //range.Select().Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous;
                //range.Select().ColorIndex = 0;
                //range.Select().TintAndShade = 0;
                //range.Select().Weight = XlBorderWeight.xlThin;
                //range.Select().Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous;
                //range.Select().ColorIndex = 0;
                //range.Select().TintAndShade = 0;
                //range.Select().Weight = XlBorderWeight.xlThin;
                //range.Select().Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
                //range.Select().ColorIndex = 0;
                //range.Select().TintAndShade = 0;
                //range.Select().Weight = XlBorderWeight.xlThin;
                //range.Select().Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous;
                //range.Select().ColorIndex = 0;
                //range.Select().TintAndShade = 0;
                //range.Select().Weight = XlBorderWeight.xlThin;
                //range.Select().Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous;
                //range.Select().ColorIndex = 0;
                //range.Select().TintAndShade = 0;
                //range.Select().Weight = XlBorderWeight.xlThin;
                //range.Select().Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous;
                //range.Select().ColorIndex = 0;
                //range.Select().TintAndShade = 0;
                //range.Select().Weight = XlBorderWeight.xlThin;
                //// Установить курсор в первую ячейку
                //range = _workSheet.Range[_workSheet.Cells[1, 1], _workSheet.Cells[1, 1]];
                //range.Select();
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
        }

        // Проверка версии

        private bool CheckVersion(ref Application objApp)
        {
            bool bVersion = false;
            // экземпляр приложения Excel
            // Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (objApp != null)
            {
                try
                {
                    System.Globalization.NumberFormatInfo provider = new System.Globalization.NumberFormatInfo { NumberDecimalSeparator = "." };
                    uint xlVersion = Convert.ToUInt16(Convert.ToDouble(objApp.Version, provider));
                    switch (xlVersion)
                    {
                        case 8:   // Excel 97
                            bVersion = true;
                            break;
                        case 10:  // Excel 2002
                            bVersion = true;
                            break;
                        case 11:  // Excel 2003
                            bVersion = true;
                            break;
                        case 12:  // Excel 2007
                            bVersion = true;
                            break;
                        case 14:  // Excel 2010
                            bVersion = true;
                            break;
                        case 15:  // Excel 2013
                            bVersion = true;
                            break;
                        //default:  // "Версия Excel не поддерживается (внутренний номер " + xlVersion.ToString() + ")";
                        //    break;
                    }
                }
                catch (Exception)
                {
                    bVersion = false;
                }
            }
            return bVersion;
        }

        // Путь к программе
        public string GetProgramPath()
        {
            return CheckVersion(ref App) ? App.Path : null;
        }

        // Проверка наличия Excel
        public bool CheckExcel()
        {
            return CheckVersion(ref App);
        }

        // Версия
        public string GetVersion()
        {
            if (CheckVersion(ref App))
            {
                try
                {
                    System.Globalization.NumberFormatInfo provider = new System.Globalization.NumberFormatInfo { NumberDecimalSeparator = "." };
                    uint xlVersion = Convert.ToUInt16(Convert.ToDouble(App.Version, provider));
                    string strVersion = "";
                    switch (xlVersion)
                    {
                        case 8:   // Excel 97
                            strVersion = "Excel 97";
                            break;
                        case 10:  // Excel 2002
                            strVersion = "Excel 2002";
                            break;
                        case 11:  // Excel 2003
                            strVersion = "Excel 2003";
                            break;
                        case 12:  // Excel 2007
                            strVersion = "Excel 2007";
                            break;
                        case 14:  // Excel 2010
                            strVersion = "Excel 2010";
                            break;
                        case 15:  // Excel 2013
                            strVersion = "Excel 2013";
                            break;
                        //default:  // "Версия Excel не поддерживается (внутренний номер " + xlVersion.ToString() + ")";
                        //    break;
                    }
                    return strVersion != ""
                        ? strVersion
                        : "Версия Excel не поддерживается (внутренний номер " + xlVersion.ToString() + ")";
                }
                catch (Exception err)
                {
                    return "Ошибка вызова Excel: " + err.Message;
                }
            }
            else
                return "Версия Excel не поддерживается!";
        }

        // Версия
        public string GetExtVersion()
        {
            if (CheckVersion(ref App))
            {
                try
                {
                    System.Globalization.NumberFormatInfo provider = new System.Globalization.NumberFormatInfo { NumberDecimalSeparator = "." };
                    uint xlVersion = Convert.ToUInt16(Convert.ToDouble(App.Version, provider));
                    string strVersion = "";
                    switch (xlVersion)
                    {
                        case 8:   // Excel 97
                            strVersion = "xls";
                            break;
                        case 10:  // Excel 2002
                            strVersion = "xls";
                            break;
                        case 11:  // Excel 2003
                            strVersion = "xls";
                            break;
                        case 12:  // Excel 2007
                            strVersion = "xlsx";
                            break;
                        case 14:  // Excel 2010
                            strVersion = "xlsx";
                            break;
                        case 15:  // Excel 2013
                            strVersion = "xlsx";
                            break;
                            //default:  // "Версия Excel не поддерживается (внутренний номер " + xlVersion.ToString() + ")";
                            //    break;
                    }
                    return strVersion != ""
                        ? strVersion
                        : "Версия Excel не поддерживается (внутренний номер " + xlVersion.ToString() + ")";
                }
                catch (Exception err)
                {
                    return "Ошибка вызова Excel: " + err.Message;
                }
            }
            else
                return "Версия Excel не поддерживается!";
        }

        /// <summary>
        /// Номер последней строки в которой что-то есть
        /// </summary>
        /// <param name="wsheet"></param>
        /// <returns></returns>
        public int LastRowCell(ref Worksheet wsheet)
        {
            int lastrow = wsheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            //for (int i = lastrow; i >= 1; i--)
            //{
            //    if (wsheet.Cells[i, 1].Value != null)
            //    {
            //        lastrow = i;
            //        break;
            //    }
            //}
            
            return lastrow;
        }

    }
}
