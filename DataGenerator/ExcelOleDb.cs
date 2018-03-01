using System.Data;
using System.Data.OleDb;

namespace DataGenerator
{
    public class ExcelOleDb
    {
        public ExcelOleDb(string pathExcelFile)
        {
            string strConnection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathExcelFile +
                    ";Extended Properties='Excel 12.0 XML;HDR=NO;IMEX=0';";
            OleDbConnection connection = new OleDbConnection(strConnection);
            connection.Open();

            DataSet = new DataSet();

            string sheet1 = @"ИП$";
            // Выбираем все данные с листа
            string select = $"SELECT * FROM [{sheet1}]";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(select, connection);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable); // Заполняем таблицу
            dataTable.TableName = sheet1.Substring(0, sheet1.Length - 1); // В конце от Экселя стоит символ '$'
            DataSet.Tables.Add(dataTable);

            connection.Close();
        }

        public DataSet DataSet {get ; set ;}

    }
}
