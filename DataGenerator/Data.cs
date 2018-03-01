using System;
using System.Data;

namespace DataGenerator
{
    public static class Data
    {
        public static DataSet DataSet { get; set; }
        public static DataSet DataSetSysId { get; set; }
        private static DataTable OipTable { get; set; }
        private static DataTable OipType { get; set; }
        private static DataTable PlaceOipTable { get; set; }
        private static DataTable SysObjectTable { get; set; }
        private static DataTable MessageTable { get; set; }
        private static DataTable ObjectTable { get; set; }
        private static DataTable TimerTable { get; set; }
        private static DataTable PlaceTimerTable { get; set; }
        private static DataTable SysIdTable { get; set; }

        public static void IninicializationDataSetSysId()
        {
            DataSetSysId = new DataSet();
            SysIdTable = new DataTable("SysID");
            CreateSysIdTableColumns();
            DataSetSysId.Tables.Add(SysIdTable);
        }

        public static void IninicializationDataSet()
        {
            DataSet = new DataSet();

            OipTable = new DataTable("OIP");
            if (Program.Settings.DB_version > 8) OipType = new DataTable("OIPType");
            PlaceOipTable = new DataTable("PlaceOIP");
            SysObjectTable = new DataTable("SysObject");
            MessageTable = new DataTable("Message");
            ObjectTable = new DataTable("Object");
            TimerTable = new DataTable("Timer");
            PlaceTimerTable = new DataTable("PlaceTimer");

            #region CreateColumns

            CreateOipColumns();
            if (Program.Settings.DB_version > 8)
            {
                CreateOipTypeColumns();
            }
            CreatePlaceOipColumns();
            CreateSysObjectColumns();
            CreateMessageColumns();
            CreateObjectColumns();
            CreateTimerColumns();
            CreatePlaceTimerColumns();
            #endregion

            #region DataSet.Add

            DataSet.Tables.Add(OipTable);
            if (Program.Settings.DB_version > 8) DataSet.Tables.Add(OipType);
            DataSet.Tables.Add(PlaceOipTable);
            DataSet.Tables.Add(SysObjectTable);
            DataSet.Tables.Add(MessageTable);
            DataSet.Tables.Add(ObjectTable);
            DataSet.Tables.Add(TimerTable);
            DataSet.Tables.Add(PlaceTimerTable);
            #endregion
        }

        public static void ClearDataSet()
        {
            DataSet.Clear();
            DataSet.Dispose();
        }

        private static void CreateSysIdTableColumns()
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.Int32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            SysIdTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnSysId = new DataColumn();
            columnSysId.DataType = Type.GetType("System.String");
            columnSysId.ColumnName = "SysID"; // имя столбца  
            columnSysId.AutoIncrement = false; //поле автоинкрементное
            columnSysId.Caption = "SysID"; // заголовок столбца
            columnSysId.ReadOnly = false; // изменение
            columnSysId.Unique = false; // уникальный
            SysIdTable.Columns.Add(columnSysId); //добавляем в таблицу

            DataColumn columnObject = new DataColumn();
            columnObject.DataType = Type.GetType("System.String");
            columnObject.ColumnName = "Object"; // имя столбца  
            columnObject.AutoIncrement = false; //поле автоинкрементное
            columnObject.Caption = "Object"; // заголовок столбца
            columnObject.ReadOnly = false; // изменение
            columnObject.Unique = false; // уникальный
            SysIdTable.Columns.Add(columnObject); //добавляем в таблицу

            DataColumn columnComment = new DataColumn();
            columnComment.DataType = Type.GetType("System.String");
            columnComment.ColumnName = "Comment"; // имя столбца  
            columnComment.AutoIncrement = false; //поле автоинкрементное
            columnComment.Caption = "Comment"; // заголовок столбца
            columnComment.ReadOnly = false; // изменение
            columnComment.Unique = false; // уникальный
            SysIdTable.Columns.Add(columnComment); //добавляем в таблицу

            DataColumn comment = new DataColumn();
            comment.DataType = Type.GetType("System.String");
            comment.ColumnName = "Comment2"; // имя столбца  
            comment.AutoIncrement = false; //поле автоинкрементное
            comment.Caption = "Comment2"; // заголовок столбца
            comment.ReadOnly = false; // изменение
            comment.Unique = false; // уникальный
            SysIdTable.Columns.Add(comment); //добавляем в таблицу

            DataColumn column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "Лист в IO"; // имя столбца  
            column.AutoIncrement = false; //поле автоинкрементное
            column.Caption = "Лист в IO"; // заголовок столбца
            column.ReadOnly = false; // изменение
            column.Unique = false; // уникальный
            SysIdTable.Columns.Add(column); //добавляем в таблицу

            DataColumn column6 = new DataColumn();
            column6.DataType = Type.GetType("System.String");
            column6.ColumnName = "Начальная Строка"; // имя столбца  
            column6.AutoIncrement = false; //поле автоинкрементное
            column6.Caption = "Начальная Строка"; // заголовок столбца
            column6.ReadOnly = false; // изменение
            column6.Unique = false; // уникальный
            SysIdTable.Columns.Add(column6); //добавляем в таблицу

            DataColumn column7 = new DataColumn();
            column7.DataType = Type.GetType("System.String");
            column7.ColumnName = "Столбец с названием"; // имя столбца  
            column7.AutoIncrement = false; //поле автоинкрементное
            column7.Caption = "Столбец с названием"; // заголовок столбца
            column7.ReadOnly = false; // изменение
            column7.Unique = false; // уникальный
            SysIdTable.Columns.Add(column7); //добавляем в таблицу

            DataColumn column8 = new DataColumn();
            column8.DataType = Type.GetType("System.String");
            column8.ColumnName = "Столбец с идентификатором"; // имя столбца  
            column8.AutoIncrement = false; //поле автоинкрементное
            column8.Caption = "Столбец с идентификатором"; // заголовок столбца
            column8.ReadOnly = false; // изменение
            column8.Unique = false; // уникальный
            SysIdTable.Columns.Add(column8); //добавляем в таблицу

            DataColumn column9 = new DataColumn();
            column9.DataType = Type.GetType("System.String");
            column9.ColumnName = "Столбец с SysNum"; // имя столбца  
            column9.AutoIncrement = false; //поле автоинкрементное
            column9.Caption = "Столбец с SysNum"; // заголовок столбца
            column9.ReadOnly = false; // изменение
            column9.Unique = false; // уникальный
            SysIdTable.Columns.Add(column9); //добавляем в таблицу

        }

        private static void CreateOipColumns()
        {
            DataColumn columnId = new DataColumn(); //1
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            OipTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnName = new DataColumn(); //2
            columnName.DataType = Type.GetType("System.String");
            columnName.ColumnName = "ParamName"; // имя столбца  
            columnName.AutoIncrement = false; //поле автоинкрементное
            columnName.Caption = "ParamName"; // заголовок столбца
            columnName.ReadOnly = false; // изменение
            columnName.Unique = false; // уникальный
            OipTable.Columns.Add(columnName); //добавляем в таблицу

            DataColumn columnParamId = new DataColumn(); //3
            columnParamId.DataType = Type.GetType("System.String");
            columnParamId.ColumnName = "ParamID"; // имя столбца  
            columnParamId.AutoIncrement = false; //поле автоинкрементное
            columnParamId.Caption = "ParamID"; // заголовок столбца
            columnParamId.ReadOnly = false; // изменение
            columnParamId.Unique = false; // уникальный
            OipTable.Columns.Add(columnParamId); //добавляем в таблицу

            DataColumn columnUnit = new DataColumn(); //4
            columnUnit.DataType = Type.GetType("System.String");
            columnUnit.ColumnName = "Unit"; // имя столбца  
            columnUnit.AutoIncrement = false; //поле автоинкрементное
            columnUnit.Caption = "Unit"; // заголовок столбца
            columnUnit.ReadOnly = false; // изменение
            columnUnit.Unique = false; // уникальный
            OipTable.Columns.Add(columnUnit); //добавляем в таблицу

            DataColumn columnScaleBeginning = new DataColumn(); //5
            columnScaleBeginning.DataType = Type.GetType("System.String");
            columnScaleBeginning.ColumnName = "ScaleBeginning"; // имя столбца  
            columnScaleBeginning.AutoIncrement = false; //поле автоинкрементное
            columnScaleBeginning.Caption = "ScaleBeginning"; // заголовок столбца
            columnScaleBeginning.ReadOnly = false; // изменение
            columnScaleBeginning.Unique = false; // уникальный
            //columnScaleBeginning.DefaultValue = 0;
            OipTable.Columns.Add(columnScaleBeginning); //добавляем в таблицу

            DataColumn columnUst0 = new DataColumn(); //6
            columnUst0.DataType = Type.GetType("System.String");
            columnUst0.ColumnName = "Ust0"; // имя столбца  
            columnUst0.AutoIncrement = false; //поле автоинкрементное
            columnUst0.Caption = "Ust0"; // заголовок столбца
            columnUst0.ReadOnly = false; // изменение
            columnUst0.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst0); //добавляем в таблицу

            DataColumn columnUst1 = new DataColumn(); //7
            columnUst1.DataType = Type.GetType("System.String");
            columnUst1.ColumnName = "Ust1"; // имя столбца  
            columnUst1.AutoIncrement = false; //поле автоинкрементное
            columnUst1.Caption = "Ust1"; // заголовок столбца
            columnUst1.ReadOnly = false; // изменение
            columnUst1.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst1); //добавляем в таблицу

            DataColumn columnUst2 = new DataColumn(); //8
            columnUst2.DataType = Type.GetType("System.String");
            columnUst2.ColumnName = "Ust2"; // имя столбца  
            columnUst2.AutoIncrement = false; //поле автоинкрементное
            columnUst2.Caption = "Ust2"; // заголовок столбца
            columnUst2.ReadOnly = false; // изменение
            columnUst2.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst2); //добавляем в таблицу

            DataColumn columnUst3 = new DataColumn(); //9
            columnUst3.DataType = Type.GetType("System.String");
            columnUst3.ColumnName = "Ust3"; // имя столбца  
            columnUst3.AutoIncrement = false; //поле автоинкрементное
            columnUst3.Caption = "Ust3"; // заголовок столбца
            columnUst3.ReadOnly = false; // изменение
            columnUst3.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst3); //добавляем в таблицу

            DataColumn columnUst4 = new DataColumn(); //10
            columnUst4.DataType = Type.GetType("System.String");
            columnUst4.ColumnName = "Ust4"; // имя столбца  
            columnUst4.AutoIncrement = false; //поле автоинкрементное
            columnUst4.Caption = "Ust4"; // заголовок столбца
            columnUst4.ReadOnly = false; // изменение
            columnUst4.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst4); //добавляем в таблицу

            DataColumn columnUst5 = new DataColumn(); //11
            columnUst5.DataType = Type.GetType("System.String");
            columnUst5.ColumnName = "Ust5"; // имя столбца  
            columnUst5.AutoIncrement = false; //поле автоинкрементное
            columnUst5.Caption = "Ust5"; // заголовок столбца
            columnUst5.ReadOnly = false; // изменение
            columnUst5.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst5); //добавляем в таблицу

            DataColumn columnUst6 = new DataColumn(); //12
            columnUst6.DataType = Type.GetType("System.String");
            columnUst6.ColumnName = "Ust6"; // имя столбца  
            columnUst6.AutoIncrement = false; //поле автоинкрементное
            columnUst6.Caption = "Ust0"; // заголовок столбца
            columnUst6.ReadOnly = false; // изменение
            columnUst6.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst6); //добавляем в таблицу

            DataColumn columnUst7 = new DataColumn(); //13
            columnUst7.DataType = Type.GetType("System.String");
            columnUst7.ColumnName = "Ust7"; // имя столбца  
            columnUst7.AutoIncrement = false; //поле автоинкрементное
            columnUst7.Caption = "Ust7"; // заголовок столбца
            columnUst7.ReadOnly = false; // изменение
            columnUst7.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst7); //добавляем в таблицу

            DataColumn columnUst8 = new DataColumn(); //14
            columnUst8.DataType = Type.GetType("System.String");
            columnUst8.ColumnName = "Ust8"; // имя столбца  
            columnUst8.AutoIncrement = false; //поле автоинкрементное
            columnUst8.Caption = "Ust8"; // заголовок столбца
            columnUst8.ReadOnly = false; // изменение
            columnUst8.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst8); //добавляем в таблицу

            DataColumn columnUst9 = new DataColumn(); //15
            columnUst9.DataType = Type.GetType("System.String");
            columnUst9.ColumnName = "Ust9"; // имя столбца  
            columnUst9.AutoIncrement = false; //поле автоинкрементное
            columnUst9.Caption = "Ust9"; // заголовок столбца
            columnUst9.ReadOnly = false; // изменение
            columnUst9.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst9); //добавляем в таблицу

            DataColumn columnUst10 = new DataColumn(); //16
            columnUst10.DataType = Type.GetType("System.String");
            columnUst10.ColumnName = "Ust10"; // имя столбца  
            columnUst10.AutoIncrement = false; //поле автоинкрементное
            columnUst10.Caption = "Ust10"; // заголовок столбца
            columnUst10.ReadOnly = false; // изменение
            columnUst10.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst10); //добавляем в таблицу

            DataColumn columnUst11 = new DataColumn(); //17
            columnUst11.DataType = Type.GetType("System.String");
            columnUst11.ColumnName = "Ust11"; // имя столбца  
            columnUst11.AutoIncrement = false; //поле автоинкрементное
            columnUst11.Caption = "Ust11"; // заголовок столбца
            columnUst11.ReadOnly = false; // изменение
            columnUst11.Unique = false; // уникальный
            OipTable.Columns.Add(columnUst11); //добавляем в таблицу

            DataColumn columnScaleEnd = new DataColumn(); //18
            columnScaleEnd.DataType = Type.GetType("System.String");
            columnScaleEnd.ColumnName = "ScaleEnd"; // имя столбца  
            columnScaleEnd.AutoIncrement = false; //поле автоинкрементное
            columnScaleEnd.Caption = "ScaleEnd"; // заголовок столбца
            columnScaleEnd.ReadOnly = false; // изменение
            columnScaleEnd.Unique = false; // уникальный
            //columnScaleEnd.DefaultValue = 10000;
            OipTable.Columns.Add(columnScaleEnd); //добавляем в таблицу

            DataColumn columnHist = new DataColumn(); //19
            columnHist.DataType = Type.GetType("System.String");
            columnHist.ColumnName = "Hist"; // имя столбца  
            columnHist.AutoIncrement = false; //поле автоинкрементное
            columnHist.Caption = "Hist"; // заголовок столбца
            columnHist.ReadOnly = false; // изменение
            columnHist.Unique = false; // уникальный
            OipTable.Columns.Add(columnHist); //добавляем в таблицу

            DataColumn columnDelta = new DataColumn(); //20
            columnDelta.DataType = Type.GetType("System.String");
            columnDelta.ColumnName = "Delta"; // имя столбца  
            columnDelta.AutoIncrement = false; //поле автоинкрементное
            columnDelta.Caption = "Delta"; // заголовок столбца
            columnDelta.ReadOnly = false; // изменение
            columnDelta.Unique = false; // уникальный
            OipTable.Columns.Add(columnDelta); //добавляем в таблицу

            DataColumn columnLimSpd = new DataColumn(); //21
            columnLimSpd.DataType = Type.GetType("System.String");
            columnLimSpd.ColumnName = "LimSpd"; // имя столбца  
            columnLimSpd.AutoIncrement = false; //поле автоинкрементное
            columnLimSpd.Caption = "LimSpd"; // заголовок столбца
            columnLimSpd.ReadOnly = false; // изменение
            columnLimSpd.Unique = false; // уникальный
            OipTable.Columns.Add(columnLimSpd); //добавляем в таблицу

            DataColumn columnAdressBeginning = new DataColumn(); //22
            columnAdressBeginning.DataType = Type.GetType("System.UInt32");
            columnAdressBeginning.ColumnName = "AdressBeginning"; // имя столбца  
            columnAdressBeginning.AutoIncrement = false; //поле автоинкрементное
            columnAdressBeginning.Caption = "AdressBeginning"; // заголовок столбца
            columnAdressBeginning.ReadOnly = false; // изменение
            columnAdressBeginning.Unique = false; // уникальный
            OipTable.Columns.Add(columnAdressBeginning); //добавляем в таблицу

            DataColumn columnPlace = new DataColumn(); //23
            columnPlace.DataType = Type.GetType("System.UInt32");
            columnPlace.ColumnName = "Place"; // имя столбца  
            columnPlace.AutoIncrement = false; //поле автоинкрементное
            columnPlace.Caption = "Place"; // заголовок столбца
            columnPlace.ReadOnly = false; // изменение
            columnPlace.Unique = false; // уникальный
            OipTable.Columns.Add(columnPlace); //добавляем в таблицу

            if (Program.Settings.DB_version < 7) return;

            DataColumn columnPlc = new DataColumn(); //24
            columnPlc.DataType = Type.GetType("System.UInt32");
            columnPlc.ColumnName = "PLC"; // имя столбца  
            columnPlc.AutoIncrement = false; //поле автоинкрементное
            columnPlc.Caption = "PLC"; // заголовок столбца
            columnPlc.ReadOnly = false; // изменение
            columnPlc.Unique = false; // уникальный
            columnPlc.DefaultValue = 1; 
            OipTable.Columns.Add(columnPlc); //добавляем в таблицу

            if (Program.Settings.DB_version < 8) return;

            DataColumn columnKf = new DataColumn //25
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "KF",
                AutoIncrement = false,
                Caption = "KF",
                ReadOnly = false,
                DefaultValue = "1"
            };
            OipTable.Columns.Add(columnKf); //добавляем в таблицу

            DataColumn columnNdiap = new DataColumn //26
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "Ndiap",
                AutoIncrement = false,
                Caption = "Ndiap",
                ReadOnly = false,
                DefaultValue = "1"
            };

            OipTable.Columns.Add(columnNdiap); //добавляем в таблицу
            DataColumn columnAdressBeginningKf = new DataColumn //27
            {
                DataType = Type.GetType("System.UInt32"),
                Unique = false,
                ColumnName = "AdressBeginningKF",
                AutoIncrement = false,
                Caption = "AdressBeginningKF",
                ReadOnly = false,
                DefaultValue = 11067
            };
            OipTable.Columns.Add(columnAdressBeginningKf); //добавляем в таблицу

            DataColumn columnCtrlUst = new DataColumn //28
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "CtrlUst",
                AutoIncrement = false,
                Caption = "CtrlUst",
                ReadOnly = false,
                DefaultValue = "1"
            };
            OipTable.Columns.Add(columnCtrlUst); //добавляем в таблицу

            DataColumn columnOpMask = new DataColumn //29
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "OpMask",
                AutoIncrement = false,
                Caption = "OpMask",
                ReadOnly = false,
                DefaultValue = "1"
            };
            OipTable.Columns.Add(columnOpMask); //добавляем в таблицу

            DataColumn columnSignMask = new DataColumn //30
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "SignMask",
                AutoIncrement = false,
                Caption = "SignMask",
                ReadOnly = false,
                DefaultValue = "1"
            };
            OipTable.Columns.Add(columnSignMask); //добавляем в таблицу

            if (Program.Settings.DB_version < 9) return;

            DataColumn columnOpcPath = new DataColumn //31
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "OpcPath",
                AutoIncrement = false,
                Caption = "OpcPath",
                ReadOnly = false,
                DefaultValue = "Path"
            };
            OipTable.Columns.Add(columnOpcPath);

            DataColumn columnPlaceName = new DataColumn //32
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "PlaceName",
                AutoIncrement = false,
                Caption = "PlaceName",
                ReadOnly = false,
                DefaultValue = "1"
            };
            OipTable.Columns.Add(columnPlaceName);

            DataColumn columnTypeName = new DataColumn //33
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "TypeName",
                AutoIncrement = false,
                Caption = "TypeName",
                ReadOnly = false,
                DefaultValue = "1"
            };
            OipTable.Columns.Add(columnTypeName);

            if (Program.Settings.DB_version < 10) return;
        }

        private static void CreateOipTypeColumns()
        {
            DataColumn columnId = new DataColumn(); //1
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            OipType.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnTypeNum = new DataColumn(); //2
            columnTypeNum.DataType = Type.GetType("System.UInt32");
            columnTypeNum.ColumnName = "TYPENUM"; // имя столбца  
            columnTypeNum.AutoIncrement = false; //поле автоинкрементное
            columnTypeNum.Caption = "TYPENUM"; // заголовок столбца
            columnTypeNum.ReadOnly = false; // изменение
            columnTypeNum.Unique = false; // уникальный
            OipType.Columns.Add(columnTypeNum); //добавляем в таблицу
            
            DataColumn columnTypeName1 = new DataColumn(); //3
            columnTypeName1.DataType = Type.GetType("System.String");
            columnTypeName1.ColumnName = "TypeName"; // имя столбца  
            columnTypeName1.AutoIncrement = false; //поле автоинкрементное
            columnTypeName1.Caption = "TypeName"; // заголовок столбца
            columnTypeName1.ReadOnly = false; // изменение
            columnTypeName1.Unique = false; // уникальный
            OipType.Columns.Add(columnTypeName1); //добавляем в таблицу
            
            DataColumn columnScaleBeginningName = new DataColumn(); //4
            columnScaleBeginningName.DataType = Type.GetType("System.String");
            columnScaleBeginningName.ColumnName = "SCALEBEGINNINGNAME"; // имя столбца  
            columnScaleBeginningName.AutoIncrement = false; //поле автоинкрементное
            columnScaleBeginningName.Caption = "SCALEBEGINNINGNAME"; // заголовок столбца
            columnScaleBeginningName.ReadOnly = false; // изменение
            columnScaleBeginningName.Unique = false; // уникальный
            OipType.Columns.Add(columnScaleBeginningName); //добавляем в таблицу

            // USTNAME0
            DataColumn columnUstName0 = new DataColumn(); //5
            columnUstName0.DataType = Type.GetType("System.String");
            columnUstName0.ColumnName = "USTNAME0"; // имя столбца  
            columnUstName0.AutoIncrement = false; //поле автоинкрементное
            columnUstName0.Caption = "USTNAME0"; // заголовок столбца
            columnUstName0.ReadOnly = false; // изменение
            columnUstName0.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName0); //добавляем в таблицу

            // USTNAME1
            DataColumn columnUstName1 = new DataColumn(); //6
            columnUstName1.DataType = Type.GetType("System.String");
            columnUstName1.ColumnName = "USTNAME1"; // имя столбца  
            columnUstName1.AutoIncrement = false; //поле автоинкрементное
            columnUstName1.Caption = "USTNAME1"; // заголовок столбца
            columnUstName1.ReadOnly = false; // изменение
            columnUstName1.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName1); //добавляем в таблицу

            // USTNAME2
            DataColumn columnUstName2 = new DataColumn(); //7
            columnUstName2.DataType = Type.GetType("System.String");
            columnUstName2.ColumnName = "USTNAME2"; // имя столбца  
            columnUstName2.AutoIncrement = false; //поле автоинкрементное
            columnUstName2.Caption = "USTNAME2"; // заголовок столбца
            columnUstName2.ReadOnly = false; // изменение
            columnUstName2.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName2); //добавляем в таблицу

            // USTNAME3
            DataColumn columnUstName3 = new DataColumn(); //8
            columnUstName3.DataType = Type.GetType("System.String");
            columnUstName3.ColumnName = "USTNAME3"; // имя столбца  
            columnUstName3.AutoIncrement = false; //поле автоинкрементное
            columnUstName3.Caption = "USTNAME3"; // заголовок столбца
            columnUstName3.ReadOnly = false; // изменение
            columnUstName3.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName3); //добавляем в таблицу

            // USTNAME4
            DataColumn columnUstName4 = new DataColumn(); //9
            columnUstName4.DataType = Type.GetType("System.String");
            columnUstName4.ColumnName = "USTNAME4"; // имя столбца  
            columnUstName4.AutoIncrement = false; //поле автоинкрементное
            columnUstName4.Caption = "USTNAME4"; // заголовок столбца
            columnUstName4.ReadOnly = false; // изменение
            columnUstName4.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName4); //добавляем в таблицу

            // USTNAME5
            DataColumn columnUstName5 = new DataColumn(); //10
            columnUstName5.DataType = Type.GetType("System.String");
            columnUstName5.ColumnName = "USTNAME5"; // имя столбца  
            columnUstName5.AutoIncrement = false; //поле автоинкрементное
            columnUstName5.Caption = "USTNAME5"; // заголовок столбца
            columnUstName5.ReadOnly = false; // изменение
            columnUstName5.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName5); //добавляем в таблицу

            // USTNAME6
            DataColumn columnUstName6 = new DataColumn(); //11
            columnUstName6.DataType = Type.GetType("System.String");
            columnUstName6.ColumnName = "USTNAME6"; // имя столбца  
            columnUstName6.AutoIncrement = false; //поле автоинкрементное
            columnUstName6.Caption = "USTNAME6"; // заголовок столбца
            columnUstName6.ReadOnly = false; // изменение
            columnUstName6.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName6); //добавляем в таблицу

            // USTNAME7
            DataColumn columnUstName7 = new DataColumn(); //12
            columnUstName7.DataType = Type.GetType("System.String");
            columnUstName7.ColumnName = "USTNAME7"; // имя столбца  
            columnUstName7.AutoIncrement = false; //поле автоинкрементное
            columnUstName7.Caption = "USTNAME7"; // заголовок столбца
            columnUstName7.ReadOnly = false; // изменение
            columnUstName7.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName7); //добавляем в таблицу

            // USTNAME8
            DataColumn columnUstName8 = new DataColumn(); //13
            columnUstName8.DataType = Type.GetType("System.String");
            columnUstName8.ColumnName = "USTNAME8"; // имя столбца  
            columnUstName8.AutoIncrement = false; //поле автоинкрементное
            columnUstName8.Caption = "USTNAME8"; // заголовок столбца
            columnUstName8.ReadOnly = false; // изменение
            columnUstName8.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName8); //добавляем в таблицу

            // USTNAME9
            DataColumn columnUstName9 = new DataColumn(); //14
            columnUstName9.DataType = Type.GetType("System.String");
            columnUstName9.ColumnName = "USTNAME9"; // имя столбца  
            columnUstName9.AutoIncrement = false; //поле автоинкрементное
            columnUstName9.Caption = "USTNAME9"; // заголовок столбца
            columnUstName9.ReadOnly = false; // изменение
            columnUstName9.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName9); //добавляем в таблицу

            // USTNAME10
            DataColumn columnUstName10 = new DataColumn(); //15
            columnUstName10.DataType = Type.GetType("System.String");
            columnUstName10.ColumnName = "USTNAME10"; // имя столбца  
            columnUstName10.AutoIncrement = false; //поле автоинкрементное
            columnUstName10.Caption = "USTNAME10"; // заголовок столбца
            columnUstName10.ReadOnly = false; // изменение
            columnUstName10.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName10); //добавляем в таблицу

            // USTNAME11
            DataColumn columnUstName11 = new DataColumn(); //16
            columnUstName11.DataType = Type.GetType("System.String");
            columnUstName11.ColumnName = "USTNAME11"; // имя столбца  
            columnUstName11.AutoIncrement = false; //поле автоинкрементное
            columnUstName11.Caption = "USTNAME11"; // заголовок столбца
            columnUstName11.ReadOnly = false; // изменение
            columnUstName11.Unique = false; // уникальный
            OipType.Columns.Add(columnUstName11); //добавляем в таблицу

            //SCALEENDNAME
            DataColumn columnScaleEndName = new DataColumn(); //17
            columnScaleEndName.DataType = Type.GetType("System.String");
            columnScaleEndName.ColumnName = "SCALEENDNAME"; // имя столбца  
            columnScaleEndName.AutoIncrement = false; //поле автоинкрементное
            columnScaleEndName.Caption = "SCALEENDNAME"; // заголовок столбца
            columnScaleEndName.ReadOnly = false; // изменение
            columnScaleEndName.Unique = false; // уникальный
            OipType.Columns.Add(columnScaleEndName); //добавляем в таблицу

            //.......................

            if (Program.Settings.DB_version < 10) return;
        }

        private static void CreatePlaceOipColumns()
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            PlaceOipTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnPlace = new DataColumn();
            columnPlace.DataType = Type.GetType("System.UInt32");
            columnPlace.ColumnName = "Place"; // имя столбца  
            columnPlace.AutoIncrement = false; //поле автоинкрементное
            columnPlace.Caption = "Place"; // заголовок столбца
            columnPlace.ReadOnly = false; // изменение
            columnPlace.Unique = true; // уникальный
            PlaceOipTable.Columns.Add(columnPlace); //добавляем в таблицу

            DataColumn columnPlaceName = new DataColumn();
            columnPlaceName.DataType = Type.GetType("System.String");
            columnPlaceName.ColumnName = "PlaceName"; // имя столбца  
            columnPlaceName.AutoIncrement = false; //поле автоинкрементное
            columnPlaceName.Caption = "PlaceName"; // заголовок столбца
            columnPlaceName.ReadOnly = false; // изменение
            columnPlaceName.Unique = false; // уникальный
            PlaceOipTable.Columns.Add(columnPlaceName); //добавляем в таблицу
        }

        private static void CreateSysObjectColumns() 
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            SysObjectTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnSysId = new DataColumn();
            columnSysId.DataType = Type.GetType("System.UInt32");
            columnSysId.ColumnName = "SysID"; // имя столбца  
            columnSysId.AutoIncrement = false; //поле автоинкрементное
            columnSysId.Caption = "SysID"; // заголовок столбца
            columnSysId.ReadOnly = false; // изменение
            columnSysId.Unique = true; // уникальный
            SysObjectTable.Columns.Add(columnSysId); //добавляем в таблицу

            DataColumn columnObject = new DataColumn();
            columnObject.DataType = Type.GetType("System.String");
            columnObject.ColumnName = "Object"; // имя столбца  
            columnObject.AutoIncrement = false; //поле автоинкрементное
            columnObject.Caption = "Object"; // заголовок столбца
            columnObject.ReadOnly = false; // изменение
            columnObject.Unique = false; // уникальный
            SysObjectTable.Columns.Add(columnObject); //добавляем в таблицу

            DataColumn columnComment = new DataColumn();
            columnComment.DataType = Type.GetType("System.String");
            columnComment.ColumnName = "Comment"; // имя столбца  
            columnComment.AutoIncrement = false; //поле автоинкрементное
            columnComment.Caption = "Comment"; // заголовок столбца
            columnComment.ReadOnly = false; // изменение
            columnComment.Unique = false; // уникальный
            SysObjectTable.Columns.Add(columnComment); //добавляем в таблицу
        } 
  
        private static void CreateMessageColumns()
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            MessageTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnSysId = new DataColumn();
            columnSysId.DataType = Type.GetType("System.UInt32");
            columnSysId.ColumnName = "SysID"; // имя столбца  
            columnSysId.AutoIncrement = false; //поле автоинкрементное
            columnSysId.Caption = "SysID"; // заголовок столбца
            columnSysId.ReadOnly = false; // изменение
            columnSysId.Unique = false; // уникальный
            MessageTable.Columns.Add(columnSysId); //добавляем в таблицу

            DataColumn columnMess = new DataColumn();
            columnMess.DataType = Type.GetType("System.UInt32");
            columnMess.ColumnName = "Mess"; // имя столбца  
            columnMess.AutoIncrement = false; //поле автоинкрементное
            columnMess.Caption = "Mess"; // заголовок столбца
            columnMess.ReadOnly = false; // изменение
            columnMess.Unique = false; // уникальный
            MessageTable.Columns.Add(columnMess); //добавляем в таблицу

            DataColumn columnMessage = new DataColumn();
            columnMessage.DataType = Type.GetType("System.String");
            columnMessage.ColumnName = "Message"; // имя столбца  
            columnMessage.AutoIncrement = false; //поле автоинкрементное
            columnMessage.Caption = "Message"; // заголовок столбца
            columnMessage.ReadOnly = false; // изменение
            columnMessage.Unique = false; // уникальный
            MessageTable.Columns.Add(columnMessage); //добавляем в таблицу

            DataColumn columnKind = new DataColumn();
            columnKind.DataType = Type.GetType("System.UInt32");
            columnKind.ColumnName = "Kind"; // имя столбца  
            columnKind.AutoIncrement = false; //поле автоинкрементное
            columnKind.Caption = "Kind"; // заголовок столбца
            columnKind.ReadOnly = false; // изменение
            columnKind.Unique = false; // уникальный
            MessageTable.Columns.Add(columnKind); //добавляем в таблицу

            DataColumn columnPriority = new DataColumn();
            columnPriority.DataType = Type.GetType("System.UInt32");
            columnPriority.ColumnName = "Priority"; // имя столбца  
            columnPriority.AutoIncrement = false; //поле автоинкрементное
            columnPriority.Caption = "Priority"; // заголовок столбца
            columnPriority.ReadOnly = false; // изменение
            columnPriority.Unique = false; // уникальный
            MessageTable.Columns.Add(columnPriority); //добавляем в таблицу

            DataColumn columnSound = new DataColumn();
            columnSound.DataType = Type.GetType("System.UInt32");
            columnSound.ColumnName = "Sound"; // имя столбца  
            columnSound.AutoIncrement = false; //поле автоинкрементное
            columnSound.Caption = "Sound"; // заголовок столбца
            columnSound.ReadOnly = false; // изменение
            columnSound.Unique = false; // уникальный
            MessageTable.Columns.Add(columnSound); //добавляем в таблицу

            DataColumn columnIdSound = new DataColumn();
            columnIdSound.DataType = Type.GetType("System.UInt32");
            columnIdSound.ColumnName = "IDSound"; // имя столбца  
            columnIdSound.AutoIncrement = false; //поле автоинкрементное
            columnIdSound.Caption = "IDSound"; // заголовок столбца
            columnIdSound.ReadOnly = false; // изменение
            columnIdSound.Unique = false; // уникальный
            MessageTable.Columns.Add(columnIdSound); //добавляем в таблицу

            DataColumn columnType = new DataColumn();
            columnType.DataType = Type.GetType("System.UInt32");
            columnType.ColumnName = "Type"; // имя столбца  
            columnType.AutoIncrement = false; //поле автоинкрементное
            columnType.Caption = "Type"; // заголовок столбца
            columnType.ReadOnly = false; // изменение
            columnType.Unique = false; // уникальный
            MessageTable.Columns.Add(columnType); //добавляем в таблицу

            DataColumn columnisAck = new DataColumn();
            columnisAck.DataType = Type.GetType("System.UInt32");
            columnisAck.ColumnName = "isAck"; // имя столбца  
            columnisAck.AutoIncrement = false; //поле автоинкрементное
            columnisAck.Caption = "isAck"; // заголовок столбца
            columnisAck.ReadOnly = false; // изменение
            columnisAck.Unique = false; // уникальный
            MessageTable.Columns.Add(columnisAck); //добавляем в таблицу

            DataColumn columnIdColor = new DataColumn();
            columnIdColor.DataType = Type.GetType("System.UInt32");
            columnIdColor.ColumnName = "IDColor"; // имя столбца  
            columnIdColor.AutoIncrement = false; //поле автоинкрементное
            columnIdColor.Caption = "IDColor"; // заголовок столбца
            columnIdColor.ReadOnly = false; // изменение
            columnIdColor.Unique = false; // уникальный
            MessageTable.Columns.Add(columnIdColor); //добавляем в таблицу
        }

        private static void CreateObjectColumns() //ID/SysID/Name/Sound1/Sound2/Sound3
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            ObjectTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnSysId = new DataColumn();
            columnSysId.DataType = Type.GetType("System.UInt32");
            columnSysId.ColumnName = "SysID"; // имя столбца  
            columnSysId.AutoIncrement = false; //поле автоинкрементное
            columnSysId.Caption = "SysID"; // заголовок столбца
            columnSysId.ReadOnly = false; // изменение
            columnSysId.Unique = false; // уникальный
            ObjectTable.Columns.Add(columnSysId); //добавляем в таблицу

            DataColumn columnSysNum = new DataColumn();
            columnSysNum.DataType = Type.GetType("System.UInt32");
            columnSysNum.ColumnName = "SysNum"; // имя столбца  
            columnSysNum.AutoIncrement = false; //поле автоинкрементное
            columnSysNum.Caption = "SysNum"; // заголовок столбца
            columnSysNum.ReadOnly = false; // изменение
            columnSysNum.Unique = false; // уникальный
            ObjectTable.Columns.Add(columnSysNum); //добавляем в таблицу

            DataColumn columnName = new DataColumn();
            columnName.DataType = Type.GetType("System.String");
            columnName.ColumnName = "Name"; // имя столбца  
            columnName.AutoIncrement = false; //поле автоинкрементное
            columnName.Caption = "Name"; // заголовок столбца
            columnName.ReadOnly = false; // изменение
            columnName.Unique = false; // уникальный
            ObjectTable.Columns.Add(columnName); //добавляем в таблицу

            DataColumn columnSound1 = new DataColumn();
            columnSound1.DataType = Type.GetType("System.Int32");
            columnSound1.ColumnName = "Sound1"; // имя столбца  
            columnSound1.AutoIncrement = false; //поле автоинкрементное
            columnSound1.Caption = "Sound1"; // заголовок столбца
            columnSound1.ReadOnly = false; // изменение
            columnSound1.Unique = false; // уникальный
            columnSound1.DefaultValue = -1;
            ObjectTable.Columns.Add(columnSound1); //добавляем в таблицу

            DataColumn columnSound2 = new DataColumn();
            columnSound2.DataType = Type.GetType("System.Int32");
            columnSound2.ColumnName = "Sound2"; // имя столбца  
            columnSound2.AutoIncrement = false; //поле автоинкрементное
            columnSound2.Caption = "Sound2"; // заголовок столбца
            columnSound2.ReadOnly = false; // изменение
            columnSound2.Unique = false; // уникальный
            columnSound2.DefaultValue = -1;
            ObjectTable.Columns.Add(columnSound2); //добавляем в таблицу

            DataColumn columnSound3 = new DataColumn();
            columnSound3.DataType = Type.GetType("System.Int32");
            columnSound3.ColumnName = "Sound3"; // имя столбца  
            columnSound3.AutoIncrement = false; //поле автоинкрементное
            columnSound3.Caption = "Sound3"; // заголовок столбца
            columnSound3.ReadOnly = false; // изменение
            columnSound3.Unique = false; // уникальный
            columnSound3.DefaultValue = -1;
            ObjectTable.Columns.Add(columnSound3); //добавляем в таблицу

            if (Program.Settings.DB_version < 7) return;
            
            DataColumn columnSoundMes = new DataColumn();
            columnSoundMes.DataType = Type.GetType("System.String");
            columnSoundMes.ColumnName = "SoundMES"; // имя столбца  
            columnSoundMes.AutoIncrement = false; //поле автоинкрементное
            columnSoundMes.Caption = "SoundMES"; // заголовок столбца
            columnSoundMes.ReadOnly = false; // изменение
            columnSoundMes.Unique = false; // уникальный
            columnSoundMes.DefaultValue = -1;
            ObjectTable.Columns.Add(columnSoundMes); //добавляем в таблицу

            DataColumn columnSoundId = new DataColumn();
            columnSoundId.DataType = Type.GetType("System.String");
            columnSoundId.ColumnName = "SoundID"; // имя столбца  
            columnSoundId.AutoIncrement = false; //поле автоинкрементное
            columnSoundId.Caption = "SoundID"; // заголовок столбца
            columnSoundId.ReadOnly = false; // изменение
            columnSoundId.Unique = false; // уникальный
            columnSoundId.DefaultValue = -1;
            ObjectTable.Columns.Add(columnSoundId); //добавляем в таблицу

            DataColumn columnSound = new DataColumn();
            columnSound.DataType = Type.GetType("System.String");
            columnSound.ColumnName = "Sound"; // имя столбца  
            columnSound.AutoIncrement = false; //поле автоинкрементное
            columnSound.Caption = "Sound"; // заголовок столбца
            columnSound.ReadOnly = false; // изменение
            columnSound.Unique = false; // уникальный
            columnSound.DefaultValue = -1;
            ObjectTable.Columns.Add(columnSound); //добавляем в таблицу
        }
        //ID/SysID/Name/Sound1/Sound2/Sound3/SoundMES/SoundID/Sound

        private static void CreateTimerColumns()
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            TimerTable.Columns.Add(columnId); //1

            DataColumn columnParamName = new DataColumn();
            columnParamName.DataType = Type.GetType("System.String");
            columnParamName.ColumnName = "ParamName"; // имя столбца  
            columnParamName.AutoIncrement = false; //поле автоинкрементное
            columnParamName.Caption = "ParamName"; // заголовок столбца
            columnParamName.ReadOnly = false; // изменение
            columnParamName.Unique = false; // уникальный
            TimerTable.Columns.Add(columnParamName); //2

            DataColumn columnParamID = new DataColumn();
            columnParamID.DataType = Type.GetType("System.String");
            columnParamID.ColumnName = "ParamID"; // имя столбца  
            columnParamID.AutoIncrement = false; //поле автоинкрементное
            columnParamID.Caption = "ParamID"; // заголовок столбца
            columnParamID.ReadOnly = false; // изменение
            columnParamID.Unique = false; // уникальный
            TimerTable.Columns.Add(columnParamID); //3

            DataColumn columnUnit = new DataColumn();
            columnUnit.DataType = Type.GetType("System.String");
            columnUnit.ColumnName = "Unit"; // имя столбца  
            columnUnit.AutoIncrement = false; //поле автоинкрементное
            columnUnit.Caption = "Unit"; // заголовок столбца
            columnUnit.ReadOnly = false; // изменение
            columnUnit.Unique = false; // уникальный
            TimerTable.Columns.Add(columnUnit); //4

            DataColumn columnValue = new DataColumn();
            columnValue.DataType = Type.GetType("System.String");
            columnValue.ColumnName = "Value"; // имя столбца  
            columnValue.AutoIncrement = false; //поле автоинкрементное
            columnValue.Caption = "Value"; // заголовок столбца
            columnValue.ReadOnly = false; // изменение
            columnValue.Unique = false; // уникальный
            TimerTable.Columns.Add(columnValue); //5

            DataColumn columnAdressBeginning = new DataColumn();
            columnAdressBeginning.DataType = Type.GetType("System.UInt32");
            columnAdressBeginning.ColumnName = "AdressBeginning"; // имя столбца  
            columnAdressBeginning.AutoIncrement = false; //поле автоинкрементное
            columnAdressBeginning.Caption = "AdressBeginning"; // заголовок столбца
            columnAdressBeginning.ReadOnly = false; // изменение
            columnAdressBeginning.Unique = false; // уникальный
            TimerTable.Columns.Add(columnAdressBeginning); //6

            DataColumn columnPlace = new DataColumn();
            columnPlace.DataType = Type.GetType("System.UInt32");
            columnPlace.ColumnName = "Place"; // имя столбца  
            columnPlace.AutoIncrement = false; //поле автоинкрементное
            columnPlace.Caption = "Place"; // заголовок столбца
            columnPlace.ReadOnly = false; // изменение
            columnPlace.Unique = false; // уникальный
            TimerTable.Columns.Add(columnPlace); //7

            if (Program.Settings.DB_version < 7) return;

            DataColumn columnPlc = new DataColumn();
            columnPlc.DataType = Type.GetType("System.UInt32");
            columnPlc.ColumnName = "PLC"; // имя столбца  
            columnPlc.AutoIncrement = false; //поле автоинкрементное
            columnPlc.Caption = "PLC"; // заголовок столбца
            columnPlc.ReadOnly = false; // изменение
            columnPlc.Unique = false; // уникальный
            TimerTable.Columns.Add(columnPlc); //8

            if (Program.Settings.DB_version < 9) return;

            DataColumn columnOpcPathTm = new DataColumn //9
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "OpcPath",
                AutoIncrement = false,
                Caption = "OpcPath",
                ReadOnly = false,
                DefaultValue = "Path"
            };
            TimerTable.Columns.Add(columnOpcPathTm); //добавляем в таблицу

            DataColumn columnPlaceNameTm = new DataColumn //10
            {
                DataType = Type.GetType("System.String"),
                Unique = false,
                ColumnName = "PlaceName",
                AutoIncrement = false,
                Caption = "PlaceName",
                ReadOnly = false
            };
            TimerTable.Columns.Add(columnPlaceNameTm); //добавляем в таблицу

            if (Program.Settings.DB_version < 10) return;

        }

        private static void CreatePlaceTimerColumns()
        {
            DataColumn columnId = new DataColumn();
            columnId.DataType = Type.GetType("System.UInt32");
            columnId.ColumnName = "ID"; // имя столбца  
            columnId.AutoIncrement = true; //поле автоинкрементное
            columnId.AutoIncrementSeed = 1;
            columnId.AutoIncrementStep = 1;
            columnId.Caption = "ID"; // заголовок столбца
            columnId.ReadOnly = true; // изменение
            columnId.Unique = true; // уникальный
            PlaceTimerTable.Columns.Add(columnId); //добавляем в таблицу

            DataColumn columnPlace = new DataColumn();
            columnPlace.DataType = Type.GetType("System.UInt32");
            columnPlace.ColumnName = "Place"; // имя столбца  
            columnPlace.AutoIncrement = false; //поле автоинкрементное
            columnPlace.Caption = "Place"; // заголовок столбца
            columnPlace.ReadOnly = false; // изменение
            columnPlace.Unique = true; // уникальный
            PlaceTimerTable.Columns.Add(columnPlace); //добавляем в таблицу

            DataColumn columnPlaceName = new DataColumn();
            columnPlaceName.DataType = Type.GetType("System.String");
            columnPlaceName.ColumnName = "PlaceName"; // имя столбца  
            columnPlaceName.AutoIncrement = false; //поле автоинкрементное
            columnPlaceName.Caption = "PlaceName"; // заголовок столбца
            columnPlaceName.ReadOnly = false; // изменение
            columnPlaceName.Unique = false; // уникальный
            PlaceTimerTable.Columns.Add(columnPlaceName); //добавляем в таблицу
        }

    }
}
