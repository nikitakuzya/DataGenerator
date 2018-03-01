using System;   
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using Microsoft.Win32;
using static DataGenerator.Properties.Settings;

namespace DataGenerator
{
    public class OleDb
    {

        #region Свойства

        /// <summary>
        /// false - не сохранять сведения безопасности, true - сохранять
        /// </summary>
        public bool PersistSecurityInfo { get; set; }

        /// <summary>
        /// false - идентификация Sql, true - Windows
        /// </summary>
        public bool IntegratedSecurityInfo { get; set; }

        /// <summary>
        /// Строка подключения 
        /// </summary>
        public string ConnectionString { get; set; } = "";

        public string Provider { get; set; }

        public string DataSource { get; set; }

        public string UserId { get; set; }

        public string[] MessArr { get; set; }

        /// <summary>
        /// Незашифрованный пароль
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Шифрованный пароль
        /// </summary>
        public string PasswordCrypt
        {
            get { return Crypt.SetPassword(Password); }
            set { Password = Crypt.GetPassword(value); }
        }

        public string InitialCatalog { get; set; }

        public string TimeOut { get; set; }

        public string TableName { get; set; }

        /// <summary>
        /// Команда
        /// </summary>
        public string CommandString { get; set; }

        public string MessErr { get; set; }

        /// <summary>
        /// Объект подключения
        /// </summary>
        public System.Data.OleDb.OleDbConnection Connection { get; private set; }

        /// <summary>
        /// Объект команд
        /// </summary>
        public OleDbCommand Command { get; private set; }

        /// <summary>
        /// Объект чтения
        /// </summary>
        public OleDbDataReader Reader { get; private set; }

        #endregion свойства
        
        public OleDb()
        {
            if (Program.Settings.DB_useRegSettings)
            {
                RegistryKey regKey32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Active\\NPS\\v1.0\\SettingDB");
                RegistryKey regKey64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Active\\NPS\\v1.0\\SettingDB");

                if (regKey64 != null)
                {
                    Provider = (string)regKey64.GetValue("Provider");
                    DataSource = (string)regKey64.GetValue("DataSource");
                    UserId = (string)regKey64.GetValue("UserId");
                    Password = Program.Settings.DB_usePasswordCrypt
                        ? Crypt.GetPassword((string) regKey64.GetValue("PasswordCrypt"))
                        : (string) regKey64.GetValue("Password");
                    InitialCatalog = (string)regKey64.GetValue("InitialCatalog");
                    TimeOut = (string)regKey64.GetValue("TimeOut");
                }
                else if (regKey32 != null)
                {
                    Provider = (string)regKey32.GetValue("Provider");
                    DataSource = (string)regKey32.GetValue("DataSource");
                    UserId = (string)regKey32.GetValue("UserId");
                    Password = Program.Settings.DB_usePasswordCrypt
                        ? Crypt.GetPassword((string) regKey32.GetValue("PasswordCrypt"))
                        : (string) regKey32.GetValue("Password");
                    InitialCatalog = (string)regKey32.GetValue("InitialCatalog");
                    TimeOut = (string)regKey32.GetValue("TimeOut");
                }
                else
                {
                    throw new Exception("Не найдены настройки соединения с базой данных в реестре");
                }
            }
        else
            {
                Provider = Program.Settings.DB_Provider;
                DataSource = Program.Settings.DB_DataSource;
                UserId = Program.Settings.DB_UserID;
                Password = Program.Settings.DB_usePasswordCrypt ? Crypt.GetPassword(Program.Settings.DB_Password) : Program.Settings.DB_Password;
                Password = Program.Settings.DB_Password;
                InitialCatalog = Program.Settings.DB_InitialCatalog;
                TimeOut = "2";
            }
        }

        public OleDb(string connectionstr)
        {
            ConnectionString = connectionstr;
        }
        
        /// <summary>
        /// Проверка подключения
        /// </summary>
        /// <returns></returns>
        public bool SqlConnected()
        {
            if (Connection == null)
            {
                return false;
            }
            switch (Connection.State)
            {
                case System.Data.ConnectionState.Broken:
                    return false;
                case System.Data.ConnectionState.Closed:
                    return false;
                case System.Data.ConnectionState.Connecting:
                    return true;
                case System.Data.ConnectionState.Executing:
                    return true;
                case System.Data.ConnectionState.Fetching:
                    return true;
                case System.Data.ConnectionState.Open:
                    return true;
                default:
                    return false;
            }
        }
        
        public bool CheckTable(string tableName, Dictionary<string, string> columns)
        {
            if (!SqlConnected()) return false;
            bool result = false;
            Dictionary<string, string> columnsExisting = new Dictionary<string, string>();
            CreateDbCommand($"SELECT * FROM {tableName}");
            ExecuteReader();
            for (int i = 0; i < Reader.FieldCount; i++)
            {
                columnsExisting.Add(Reader.GetName(i).ToUpper(), Reader.GetDataTypeName(i));
            }
            foreach (KeyValuePair<string, string> col in columns)
            {
                result = columnsExisting.ContainsKey(col.Key.ToUpper());
            }
            return result;
        }

        public void Open()
        {
            try
            {
                ConnectionString = "Provider=" + Provider + ";" +
                                   "Persist Security Info=" + Convert.ToString(PersistSecurityInfo) + ";" +
                                   "Integrated Security Info=" + Convert.ToString(IntegratedSecurityInfo) + ";" +
                                   "User ID=" + UserId + ";" +
                                   "Password=" + Password + ";" +
                                   "Data Source=" + DataSource + ";" +
                                   "Initial Catalog=" + InitialCatalog +
                                   ";Timeout=" + Convert.ToString(TimeOut);
                if (Connection != null) return;
                Connection = new System.Data.OleDb.OleDbConnection(ConnectionString);
                Connection.InfoMessage += new OleDbInfoMessageEventHandler(Connection_InfoMessage);
                Connection.Open();
            }
            catch (Exception exception)
            {
                MessErr = exception.Message;
                throw;
            }
        }

        private void Connection_InfoMessage(object sender, OleDbInfoMessageEventArgs args)
        {
            List<string> messList = new List<string>();
            messList.Add(args.Source + " : " + args.Message);
            messList.AddRange(from OleDbError err in args.Errors select err.Message);
            MessArr = messList.ToArray();
        }

        public void CreateDbCommand(string queryString)
        {
            try
            {
                Command = Connection.CreateCommand();
                Command.CommandText = queryString;
            }
            catch (Exception exception)
            {
                MessErr = exception.Message;
            }
        }

        public void CreateDbCommand(string[] queryString)
        {
            try
            {
                Command = Connection.CreateCommand();
                foreach (var str in queryString)
                {
                    Command.CommandText += str;
                    Command.CommandText += System.Environment.NewLine;
                }
         }
            catch (Exception exception)
            {
                MessErr = exception.Message;
            }
        }

        public void ExecuteReader()
        {
            try
            {
                Reader = Command.ExecuteReader();
        }
            catch (Exception exception)
            {
                MessErr = exception.Message;
            }
}

        public Exception ExecuteReader(string[] lines)
        {
            try
            {
                var Command = "";
                foreach (var line in lines)
                {
                    if (line.Trim(' ').ToUpper() != "GO")
                    {
                        Command += line;
                    }
                    else
                    {
                        CreateDbCommand(Command);
                        ExecuteReader();
                        Command = "";
                    }
                }
                return null;
            }
            catch (Exception exception)
            {
                return exception;
            }
        }

        public void Close()
        {
            try
            {
                Connection.Close();
            }
            catch (Exception exception)
            {
                MessErr = exception.Message;
            }
        }

        public static Dictionary<string, Dictionary<string, string>> GetUsingTablesStructure()
        {
            Dictionary<string, Dictionary<string, string>> resultList =
                new Dictionary<string, Dictionary<string, string>>();

            var structure = new Dictionary<string, string>
            {
                {"IDSound", "int"},
                {"PathSound", "string"}
            };
            //resultList.Add("IDSound", structure);

            structure = new Dictionary<string, string>
            {
                {"IDColor", "int"}
            };
            //resultList.Add("IDColors", structure);
            structure = new Dictionary<string, string>
            {
                {"UserLevel", "int"},
                {"UserName", "string"}
            };
            //resultList.Add("Registration", structure);

            if (Properties.Settings.Default.DB_version < 7) 
            {
                structure = new Dictionary<string, string>
                {
                    {"ID", "int"},
                    {"ParamName", "string"},
                    {"ParamID", "string"},
                    {"Unit", "string"},
                    {"ScaleBeginning", "string"},
                    {"Ust0", "string"},
                    {"Ust1", "string"},
                    {"Ust2", "string"},
                    {"Ust3", "string"},
                    {"Ust4", "string"},
                    {"Ust5", "string"},
                    {"Ust6", "string"},
                    {"Ust7", "string"},
                    {"Ust8", "string"},
                    {"Ust9", "string"},
                    {"Ust10", "string"},
                    {"Ust11", "string"},
                    {"ScaleEnd", "string"},
                    {"Hist", "string"},
                    {"Delta", "string"},
                    {"LimSpd", "string"},
                    {"AdressBeginning", "string"},
                    {"Place", "string"}
                };
            }
            else
            {
                structure = new Dictionary<string, string>
                {
                    {"ID", "int"},
                    {"ParamName", "string"},
                    {"ParamID", "string"},
                    {"Unit", "string"},
                    {"ScaleBeginning", "string"},
                    {"Ust0", "string"},
                    {"Ust1", "string"},
                    {"Ust2", "string"},
                    {"Ust3", "string"},
                    {"Ust4", "string"},
                    {"Ust5", "string"},
                    {"Ust6", "string"},
                    {"Ust7", "string"},
                    {"Ust8", "string"},
                    {"Ust9", "string"},
                    {"Ust10", "string"},
                    {"Ust11", "string"},
                    {"ScaleEnd", "string"},
                    {"Hist", "string"},
                    {"Delta", "string"},
                    {"LimSpd", "string"},
                    {"AdressBeginning", "string"},
                    {"Place", "string"},
                    {"PLC", "string"}
                };
            }
            resultList.Add("OIP", structure);

            structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"Place", "int"},
                {"PlaceName", "string"}
            };
            resultList.Add("PlaceOIP", structure);

            structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"SysID", "int"},
                {"Object", "string"},
                {"Comment", "string"}
            };
            resultList.Add("SysObject", structure);

            structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"IDmess", "int"},
                {"UserName", "string"},
                {"Place", "string"},
                {"DTime", "string"},
                {"DTimeAck", "string"},
                {"SysID", "int"},
                {"SysNum", "int"},
                {"Mess", "int"},
                {"Message", "string"},
                {"isAck", "int"},
                {"Priority", "int"},
                {"Value", "string"}
            };
            resultList.Add("PLCMessage", structure);

            structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"SysID", "int"},
                {"Mess", "int"},
                {"Message", "string"},
                {"Kind", "int"},
                {"Priority", "int"},
                {"Sound", "int"},
                {"IDSound", "int"},
                {"Type", "int"},
                {"IsAck", "int"},
                {"IDColor", "int"}
            };
            resultList.Add("Message", structure);

            if (Properties.Settings.Default.DB_version < 7)
                structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"ParamName", "string"},
                {"ParamID", "string"},
                {"Unit", "string"},
                {"Value", "string"},
                {"AdressBeginning", "int"},
                {"Place", "int"}
            };
            else
                structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"ParamName", "string"},
                {"ParamID", "string"},
                {"Unit", "string"},
                {"Value", "string"},
                {"AdressBeginning", "int"},
                {"Place", "int"},
                {"PLC", "int"}
            };
            resultList.Add("Timer", structure);

            structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"Place", "int"},
                {"PlaceName", "string"}
            };
            resultList.Add("PlaceTimer", structure);

            if (Properties.Settings.Default.DB_version < 7)
            {
                structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"SysID", "int"},
                {"SysNum", "int"},
                {"Name", "string"},
                {"Sound1", "int"},
                {"Sound2", "int"},
                {"Sound3", "int"}
            };
            }
            else
            {
                structure = new Dictionary<string, string>
            {
                {"ID", "int"},
                {"SysID", "int"},
                {"SysNum", "int"},
                {"Name", "string"},
                {"Sound1", "int"},
                {"Sound2", "int"},
                {"Sound3", "int"},
                {"SoundMES", "string"},
                {"setSoundID", "string"},
                {"Sound", "string"}
            };
            }
            resultList.Add("Object", structure);

            return resultList;
        }

        /// <summary>
        ///  Очистить выделенную таблицу БД
        /// </summary>
        /// <param name="table"></param>
        public void TruncateTable(string table)
        {
            switch (table)
            {
                case "PlaceOIP":
                case "SysObject":
                case "PlaceTimer":
                    CreateDbCommand("DELETE FROM " + table);
                    ExecuteReader();
                    CreateDbCommand("DBCC CHECKIDENT (" + table + ", RESEED, 0) ");
                    ExecuteReader();
                    break;
                default:
                    CreateDbCommand("truncate table " + table);
                    ExecuteReader();
                    break;
            }
        }

        #region Методы - Запись в БД

        /// <summary>
        /// Очистить ошибки
        /// </summary>
        public void ClearMessErr()
        {
            MessErr = null;
        }

        /// <summary>
        /// Запись в Accounts
        /// </summary>
        /// <param name="iUserId"></param>
        /// <param name="strLogin"></param>
        /// <param name="strPassword"></param>
        /// <param name="strUserName"></param>
        /// <param name="iGroupId"></param>
        /// <param name="strJob"></param>
        /// <param name="strFirm"></param>
        /// <param name="iIsActive"></param>
        /// <returns></returns>
        public bool WriteOperMess_Accounts(uint iUserId, string strLogin, string strPassword, string strUserName, uint iGroupId, string strJob, string strFirm, uint iIsActive)
        {
            CommandString = "";
            ClearMessErr();
            if ((strLogin != null) && (strPassword != null) && (strUserName != null) && (strJob != null) && (strFirm != null))
            {
                try
                {
                    CreateDbCommand("insert into Accounts (UserID, Login, Password, UserName, GroupID, Job, Firm, isActive) " + "values (?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("UserID", OleDbType.Integer).Value = iUserId;
                    Command.Parameters.Add("Login", OleDbType.VarWChar).Value = strLogin;
                    Command.Parameters.Add("Password", OleDbType.VarWChar).Value = strPassword;
                    Command.Parameters.Add("UserName", OleDbType.VarWChar).Value = strUserName;
                    Command.Parameters.Add("GroupID", OleDbType.Integer).Value = iGroupId;
                    Command.Parameters.Add("Job", OleDbType.VarWChar).Value = strJob;
                    Command.Parameters.Add("Firm", OleDbType.VarWChar).Value = strFirm;
                    Command.Parameters.Add("isActive", OleDbType.Integer).Value = iIsActive;
                    CommandString = "insert into Accounts (UserID, Login, Password, UserName, GroupID, Job, Firm, isActive) " + "values (" + Convert.ToString(iUserId) + ", '" + strLogin + "', '" + strPassword + "', '" + strUserName + "', " + Convert.ToString(iGroupId) + ", '" + strJob + "', '" + strFirm + "', " + Convert.ToString(iIsActive) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Accounts расширенная версия
        /// </summary>
        /// <param name="iUserId"></param>
        /// <param name="strLogin"></param>
        /// <param name="strPassword"></param>
        /// <param name="strUserName"></param>
        /// <param name="iGroupId"></param>
        /// <param name="strJob"></param>
        /// <param name="strFirm"></param>
        /// <param name="iIsActive"></param>
        /// <param name="strOldPassword"></param>
        /// <param name="strTimePasswordChange"></param>
        /// <param name="strPassword1"></param>
        /// <param name="strPassword2"></param>
        /// <param name="strPassword3"></param>
        /// <param name="strPassword4"></param>
        /// <returns></returns>
        public bool WriteOperMess_Accounts_2(uint iUserId, string strLogin, string strPassword, string strUserName, uint iGroupId, string strJob, string strFirm, uint iIsActive, string strOldPassword, string strTimePasswordChange, string strPassword1, string strPassword2, string strPassword3, string strPassword4)
        {
            CommandString = "";
            ClearMessErr();
            if ((strLogin != null) && (strPassword != null) && (strUserName != null) && (strJob != null) && (strFirm != null) && (strOldPassword != null) && (strTimePasswordChange != null) && (strPassword1 != null) && (strPassword2 != null) && (strPassword3 != null) && (strPassword4 != null))
            {
                try
                {
                    CreateDbCommand("insert into Accounts (UserID, Login, Password, UserName, GroupID, Job, Firm, isActive, OldPassword, TimePasswordChange, Password1, Password2, Password3, Password4) " + "values (?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("UserID", OleDbType.Integer).Value = iUserId;
                    Command.Parameters.Add("Login", OleDbType.VarWChar).Value = strLogin;
                    Command.Parameters.Add("Password", OleDbType.VarWChar).Value = strPassword != "" ? strPassword : null;
                    Command.Parameters.Add("UserName", OleDbType.VarWChar).Value = strUserName;
                    Command.Parameters.Add("GroupID", OleDbType.Integer).Value = iGroupId;
                    Command.Parameters.Add("Job", OleDbType.VarWChar).Value = strJob;
                    Command.Parameters.Add("Firm", OleDbType.VarWChar).Value = strFirm;
                    Command.Parameters.Add("isActive", OleDbType.Integer).Value = iIsActive;
                    Command.Parameters.Add("OldPassword", OleDbType.VarWChar).Value = strOldPassword != "" ? strOldPassword : null;
                    Command.Parameters.Add("TimePasswordChange", OleDbType.VarWChar).Value = strTimePasswordChange != "" ? strTimePasswordChange : null;
                    Command.Parameters.Add("Password1", OleDbType.VarWChar).Value = strPassword1 != "" ? strPassword1 : null;
                    Command.Parameters.Add("Password2", OleDbType.VarWChar).Value = strPassword2 != "" ? strPassword2 : null;
                    Command.Parameters.Add("Password3", OleDbType.VarWChar).Value = strPassword3 != "" ? strPassword3 : null;
                    Command.Parameters.Add("Password4", OleDbType.VarWChar).Value = strPassword4 != "" ? strPassword4 : null;
                    CommandString = "insert into Accounts (UserID, Login, Password, UserName, GroupID, Job, Firm, isActive, " + "OldPassword, TimePasswordChange, Password1, Password2, Password3, Password4) " + "values (" + Convert.ToString(iUserId) + ", '" + strLogin + "', '" + strPassword + "', '" + strUserName + "', " + Convert.ToString(iGroupId) + ", '" + strJob + "', '" + strFirm + "', " + Convert.ToString(iIsActive) + "', '" + strOldPassword + "', '" + strTimePasswordChange + "', '" + strPassword1 + "', '" + strPassword2 + "', '" + strPassword3 + "', '" + strPassword4 + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в ButtSetup
        /// </summary>
        /// <param name="iNButt"></param>
        /// <param name="iAdrButt"></param>
        /// <returns></returns>
        public bool WriteOperMess_ButtSetup(uint iNButt, uint iAdrButt)
        {
            CommandString = "";
            ClearMessErr();
            try
            {
                CreateDbCommand("insert into ButtSetup (NButt, AdrButt) values (?,?)");
                Command.Parameters.Add("NButt", OleDbType.Integer).Value = iNButt;
                Command.Parameters.Add("AdrButt", OleDbType.Integer).Value = iAdrButt;
                CommandString = "insert into ButtSetup (NButt, AdrButt) " + "values (" + Convert.ToString(iNButt) + ", " + Convert.ToString(iAdrButt) + ")";
                Command.ExecuteNonQuery();
                return true;
            }
            catch (Exception exception)
            {
                MessErr = exception.Message;
            }
            return false;
        }

        /// <summary>
        /// Запись в IDColors
        /// </summary>
        /// <param name="iIdColor"></param>
        /// <param name="strArgbColorBg"></param>
        /// <param name="strNameColorBg"></param>
        /// <param name="strArgbColorText"></param>
        /// <param name="strNameColorText"></param>
        /// <returns></returns>
        public bool WriteOperMess_IDColors(uint iIdColor, string strArgbColorBg, string strNameColorBg, string strArgbColorText, string strNameColorText)
        {
            CommandString = "";
            ClearMessErr();
            if ((strArgbColorBg != null) && (strNameColorBg != null) && (strArgbColorText != null) && (strNameColorText != null))
            {
                try
                {
                    CreateDbCommand("insert into IDColors (IDColor, ARGBColorBG, NameColorBG, ARGBColorText, NameColorText) " + "values (?,?,?,?,?)");
                    Command.Parameters.Add("IDColor", OleDbType.Integer).Value = iIdColor;
                    Command.Parameters.Add("ARGBColorBG", OleDbType.VarWChar).Value = strArgbColorBg;
                    Command.Parameters.Add("NameColorBG", OleDbType.VarWChar).Value = strNameColorBg;
                    Command.Parameters.Add("ARGBColorText", OleDbType.VarWChar).Value = strArgbColorText;
                    Command.Parameters.Add("NameColorText", OleDbType.VarWChar).Value = strNameColorText;
                    CommandString = "insert into IDColors (IDColor, ARGBColorBG, NameColorBG, ARGBColorText, NameColorText) " + "values (" + Convert.ToString(iIdColor) + ", '" + strArgbColorBg + "', '" + strNameColorBg + "', '" + strArgbColorText + "', '" + strNameColorText + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в IDSound
        /// </summary>
        /// <param name="iIdSound"></param>
        /// <param name="strPathSound"></param>
        /// <returns></returns>
        public bool WriteOperMess_IDSound(uint iIdSound, string strPathSound)
        {
            CommandString = "";
            ClearMessErr();
            if (strPathSound != null)
            {
                try
                {
                    CreateDbCommand("insert into IDSound (IDSound, PathSound) values (?,?)");
                    Command.Parameters.Add("IDSound", OleDbType.Integer).Value = iIdSound;
                    Command.Parameters.Add("PathSound", OleDbType.VarWChar).Value = strPathSound;
                    CommandString = "insert into IDColors IDSound (IDSound, PathSound) " + "values (" + Convert.ToString(iIdSound) + ", '" + strPathSound + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Kind
        /// </summary>
        /// <param name="iKind"></param>
        /// <param name="strName"></param>
        /// <returns></returns>
        public bool WriteOperMess_Kind(uint iKind, string strName)
        {
            CommandString = "";
            ClearMessErr();
            if (strName != null)
            {
                try
                {
                    CreateDbCommand("insert into Kind (Kind, Name) values (?,?)");
                    Command.Parameters.Add("Kind", OleDbType.Integer).Value = iKind;
                    Command.Parameters.Add("Name", OleDbType.VarWChar).Value = strName;
                    CommandString = "insert into Kind (Kind, Name) " + "values (" + Convert.ToString(iKind) + ", '" + strName + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Message
        /// </summary>
        /// <param name="iSysId"></param>
        /// <param name="iMess"></param>
        /// <param name="strMessage"></param>
        /// <param name="iKind"></param>
        /// <param name="iPriority"></param>
        /// <param name="iSound"></param>
        /// <param name="iIdSound"></param>
        /// <param name="iType"></param>
        /// <param name="iIsAck"></param>
        /// <param name="iIdColor"></param>
        /// <returns></returns>
        public bool WriteOperMess_Message(uint iSysId, uint iMess, string strMessage, uint iKind, uint iPriority, uint iSound, uint iIdSound, uint iType, uint iIsAck, uint iIdColor)
        {
            CommandString = "";
            ClearMessErr();
            if (strMessage != null)
            {
                try
                {
                    CreateDbCommand("insert into Message (SysID, Mess, Message, Kind, Priority, Sound, IDSound, Type, isAck, IDColor) " + "values (?,?,?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("SysID", OleDbType.Integer).Value = iSysId;
                    Command.Parameters.Add("Mess", OleDbType.Integer).Value = iMess;
                    Command.Parameters.Add("Message", OleDbType.VarWChar).Value = strMessage;
                    Command.Parameters.Add("Kind", OleDbType.Integer).Value = iKind;
                    Command.Parameters.Add("Priority", OleDbType.Integer).Value = iPriority;
                    Command.Parameters.Add("Sound", OleDbType.Integer).Value = iSound;
                    Command.Parameters.Add("IDSound", OleDbType.Integer).Value = iIdSound;
                    Command.Parameters.Add("Type", OleDbType.Integer).Value = iType;
                    Command.Parameters.Add("IsAck", OleDbType.Integer).Value = iIsAck;
                    Command.Parameters.Add("IDColor", OleDbType.Integer).Value = iIdColor;
                    CommandString = "insert into Message (SysID, Mess, Message, Kind, Priority, Sound, IDSound, Type, isAck, IDColor) " + "values (" + Convert.ToString(iSysId) + ", '" + Convert.ToString(iMess) + "', '" + strMessage + "', '" + Convert.ToString(iKind) + ", " + Convert.ToString(iPriority) + ", " + Convert.ToString(iSound) + ", " + Convert.ToString(iIdSound) + ", " + Convert.ToString(iType) + ", " + Convert.ToString(iIsAck) + ", " + Convert.ToString(iIdColor) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Object версии 6
        /// </summary>
        /// <param name="iSysId"></param>
        /// <param name="iSysNum"></param>
        /// <param name="strName"></param>
        /// <param name="iSound1"></param>
        /// <param name="iSound2"></param>
        /// <param name="iSound3"></param>
        /// <returns></returns>
        public bool WriteOperMess_Object(uint iSysId, uint iSysNum, string strName, int iSound1, int iSound2, int iSound3)
        {
            CommandString = "";
            ClearMessErr();
            if (strName != null)
            {
                try
                {
                    CreateDbCommand("insert into Object (SysID, SysNum, Name, Sound1, Sound2, Sound3) values (?,?,?,?,?,?)");
                    Command.Parameters.Add("SysID", OleDbType.Integer).Value = iSysId;
                    Command.Parameters.Add("SysNum", OleDbType.Integer).Value = iSysNum;
                    Command.Parameters.Add("Name", OleDbType.VarWChar).Value = strName;
                    Command.Parameters.Add("Sound1", OleDbType.Integer).Value = iSound1;
                    Command.Parameters.Add("Sound2", OleDbType.Integer).Value = iSound2;
                    Command.Parameters.Add("Sound3", OleDbType.Integer).Value = iSound2;
                    CommandString = "insert into Object (SysID, SysNum, Name, Sound1, Sound2, Sound3) " + "values (" + Convert.ToString(iSysId) + ", " + Convert.ToString(iSysNum) + ", '" + strName + "', '" + Convert.ToString(iSound1) + ", " + Convert.ToString(iSound2) + ", " + Convert.ToString(iSound3) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Object версии 7
        /// </summary>
        /// <param name="iSysId"></param>
        /// <param name="iSysNum"></param>
        /// <param name="strName"></param>
        /// <param name="iSound1"></param>
        /// <param name="iSound2"></param>
        /// <param name="iSound3"></param>
        /// <param name="iSoundMes"></param>
        /// <param name="iSoundId"></param>
        /// <param name="iSound"></param>
        /// <returns></returns>
        public bool WriteOperMess_Object(uint iSysId, uint iSysNum, string strName, int iSound1, int iSound2, int iSound3, string iSoundMes, string iSoundId, string iSound)
        {
            CommandString = "";
            ClearMessErr();
            if (strName != null)
            {
                try
                {
                    CreateDbCommand("insert into Object (SysID, SysNum, Name, Sound1, Sound2, Sound3, SoundMES, SoundID, Sound) values (?,?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("SysID", OleDbType.Integer).Value = iSysId;
                    Command.Parameters.Add("SysNum", OleDbType.Integer).Value = iSysNum;
                    Command.Parameters.Add("Name", OleDbType.VarWChar).Value = strName;
                    Command.Parameters.Add("Sound1", OleDbType.Integer).Value = iSound1;
                    Command.Parameters.Add("Sound2", OleDbType.Integer).Value = iSound2;
                    Command.Parameters.Add("Sound3", OleDbType.Integer).Value = iSound2;
                    Command.Parameters.Add("SoundMES", OleDbType.VarWChar).Value = iSoundMes;
                    Command.Parameters.Add("SoundID", OleDbType.VarWChar).Value = iSoundId;
                    Command.Parameters.Add("Sound", OleDbType.VarWChar).Value = iSound;
                    CommandString = "insert into Object (SysID, SysNum, Name, Sound1, Sound2, Sound3, SoundMES, SoundID, Sound) " + "values (" + Convert.ToString(iSysId) + ", " + Convert.ToString(iSysNum) + ", '" + strName + "', '" + Convert.ToString(iSound1) + ", " + Convert.ToString(iSound2) + ", " + Convert.ToString(iSound3)+ ", '" + iSoundMes + "', '" + iSoundId + "', '" + iSound + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в OIP версии 6
        /// </summary>
        /// <param name="strParamName"></param>
        /// <param name="strParamId"></param>
        /// <param name="strUnit"></param>
        /// <param name="strScaleBeginning"></param>
        /// <param name="strUst0"></param>
        /// <param name="strUst1"></param>
        /// <param name="strUst2"></param>
        /// <param name="strUst3"></param>
        /// <param name="strUst4"></param>
        /// <param name="strUst5"></param>
        /// <param name="strUst6"></param>
        /// <param name="strUst7"></param>
        /// <param name="strUst8"></param>
        /// <param name="strUst9"></param>
        /// <param name="strUst10"></param>
        /// <param name="strUst11"></param>
        /// <param name="strScaleEnd"></param>
        /// <param name="strHist"></param>
        /// <param name="strDelta"></param>
        /// <param name="strLimSpd"></param>
        /// <param name="iAdressBeginning"></param>
        /// <param name="iPlace"></param>
        /// <returns></returns>
        public bool WriteOperMess_OIP(string strParamName, string strParamId, string strUnit, string strScaleBeginning, string strUst0, string strUst1, string strUst2, string strUst3, string strUst4, string strUst5, string strUst6, string strUst7, string strUst8, string strUst9, string strUst10, string strUst11, string strScaleEnd, string strHist, string strDelta, string strLimSpd, uint iAdressBeginning, uint iPlace)
        {
            CommandString = "";
            ClearMessErr();
            if ((strParamName != null) && (strParamId != null) && (strUnit != null) && (strScaleBeginning != null) && 
                (strUst0 != null) && (strUst1 != null) && (strUst2 != null) && (strUst3 != null) && (strUst4 != null) && 
                (strUst5 != null) && (strUst6 != null) && (strUst7 != null) && (strUst8 != null) && (strUst9 != null) && 
                (strUst10 != null) && (strUst11 != null) && (strScaleEnd != null) && (strHist != null) && (strDelta != null) && 
                (strLimSpd != null))
            {
                try
                {
                    CreateDbCommand("insert into OIP (ParamName, ParamID, Unit, ScaleBeginning, Ust0, Ust1, Ust2, Ust3, Ust4, Ust5, Ust6, Ust7, Ust8, Ust9, Ust10, Ust11, ScaleEnd, Hist, Delta, LimSpd, AdressBeginning, Place) " + "values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("ParamName", OleDbType.VarWChar).Value = strParamName;
                    Command.Parameters.Add("ParamID", OleDbType.VarWChar).Value = strParamId;
                    Command.Parameters.Add("Unit", OleDbType.VarWChar).Value = strUnit;
                    Command.Parameters.Add("ScaleBeginning", OleDbType.VarWChar).Value = strScaleBeginning;
                    Command.Parameters.Add("Ust0", OleDbType.VarWChar).Value = strUst0;
                    Command.Parameters.Add("Ust1", OleDbType.VarWChar).Value = strUst1;
                    Command.Parameters.Add("Ust2", OleDbType.VarWChar).Value = strUst2;
                    Command.Parameters.Add("Ust3", OleDbType.VarWChar).Value = strUst3;
                    Command.Parameters.Add("Ust4", OleDbType.VarWChar).Value = strUst4;
                    Command.Parameters.Add("Ust5", OleDbType.VarWChar).Value = strUst5;
                    Command.Parameters.Add("Ust6", OleDbType.VarWChar).Value = strUst6;
                    Command.Parameters.Add("Ust7", OleDbType.VarWChar).Value = strUst7;
                    Command.Parameters.Add("Ust8", OleDbType.VarWChar).Value = strUst8;
                    Command.Parameters.Add("Ust9", OleDbType.VarWChar).Value = strUst9;
                    Command.Parameters.Add("Ust10", OleDbType.VarWChar).Value = strUst10;
                    Command.Parameters.Add("Ust11", OleDbType.VarWChar).Value = strUst11;
                    Command.Parameters.Add("ScaleEnd", OleDbType.VarWChar).Value = strScaleEnd;
                    Command.Parameters.Add("Hist", OleDbType.VarWChar).Value = strHist;
                    Command.Parameters.Add("Delta", OleDbType.VarWChar).Value = strDelta;
                    Command.Parameters.Add("LimSpd", OleDbType.VarWChar).Value = strLimSpd;
                    Command.Parameters.Add("AdressBeginning", OleDbType.Integer).Value = iAdressBeginning;
                    Command.Parameters.Add("Place", OleDbType.Integer).Value = iPlace;
                    CommandString = "insert into OIP (ParamName, ParamID, Unit, ScaleBeginning, Ust0, Ust1, Ust2, Ust3, Ust4, Ust5, Ust6, Ust7, Ust8, Ust9, Ust10, Ust11, ScaleEnd, Hist, Delta, LimSpd, AdressBeginning, Place) " + "values ('" + strParamName + "', '" + strParamId + "', '" + strUnit + "', '" + strScaleBeginning + "', '" + strUst0 + "', '" + strUst1 + "', '" + strUst2 + "', '" + strUst3 + "', '" + strUst4 + "', '" + strUst5 + "', '" + strUst6 + "', '" + strUst7 + "', '" + strUst8 + "', '" + strUst9 + "', '" + strUst10 + "', '" + strUst11 + "', '" + strScaleEnd + "', '" + strHist + "', '" + strDelta + "', '" + strLimSpd + "', " + Convert.ToString(iAdressBeginning) + ", " + Convert.ToString(iPlace) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в OIP версии 7
        /// </summary>
        /// <param name="strParamName"></param>
        /// <param name="strParamId"></param>
        /// <param name="strUnit"></param>
        /// <param name="strScaleBeginning"></param>
        /// <param name="strUst0"></param>
        /// <param name="strUst1"></param>
        /// <param name="strUst2"></param>
        /// <param name="strUst3"></param>
        /// <param name="strUst4"></param>
        /// <param name="strUst5"></param>
        /// <param name="strUst6"></param>
        /// <param name="strUst7"></param>
        /// <param name="strUst8"></param>
        /// <param name="strUst9"></param>
        /// <param name="strUst10"></param>
        /// <param name="strUst11"></param>
        /// <param name="strScaleEnd"></param>
        /// <param name="strHist"></param>
        /// <param name="strDelta"></param>
        /// <param name="strLimSpd"></param>
        /// <param name="iAdressBeginning"></param>
        /// <param name="iPlace"></param>
        /// <returns></returns>
        public bool WriteOperMess_OIP(string strParamName, string strParamId, string strUnit, string strScaleBeginning, string strUst0, string strUst1, string strUst2, string strUst3, string strUst4, string strUst5, string strUst6, string strUst7, string strUst8, string strUst9, string strUst10, string strUst11, string strScaleEnd, string strHist, string strDelta, string strLimSpd, uint iAdressBeginning, uint iPlace, uint iPlc)
        {
            CommandString = "";
            ClearMessErr();
            if ((strParamName != null) && (strParamId != null) && (strUnit != null) && (strScaleBeginning != null) && (strUst0 != null) && (strUst1 != null) && (strUst2 != null) && (strUst3 != null) && (strUst4 != null) && (strUst5 != null) && (strUst6 != null) && (strUst7 != null) && (strUst8 != null) && (strUst9 != null) && (strUst10 != null) && (strUst11 != null) && (strScaleEnd != null) && (strHist != null) && (strDelta != null) && (strLimSpd != null))
            {
                try
                {
                    CreateDbCommand("insert into OIP (ParamName, ParamID, Unit, ScaleBeginning, Ust0, Ust1, Ust2, Ust3, Ust4, Ust5, Ust6, Ust7, Ust8, Ust9, Ust10, Ust11, ScaleEnd, Hist, Delta, LimSpd, AdressBeginning, Place, PLC) " + "values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("ParamName", OleDbType.VarWChar).Value = strParamName;
                    Command.Parameters.Add("ParamID", OleDbType.VarWChar).Value = strParamId;
                    Command.Parameters.Add("Unit", OleDbType.VarWChar).Value = strUnit;
                    Command.Parameters.Add("ScaleBeginning", OleDbType.VarWChar).Value = strScaleBeginning;
                    Command.Parameters.Add("Ust0", OleDbType.VarWChar).Value = strUst0;
                    Command.Parameters.Add("Ust1", OleDbType.VarWChar).Value = strUst1;
                    Command.Parameters.Add("Ust2", OleDbType.VarWChar).Value = strUst2;
                    Command.Parameters.Add("Ust3", OleDbType.VarWChar).Value = strUst3;
                    Command.Parameters.Add("Ust4", OleDbType.VarWChar).Value = strUst4;
                    Command.Parameters.Add("Ust5", OleDbType.VarWChar).Value = strUst5;
                    Command.Parameters.Add("Ust6", OleDbType.VarWChar).Value = strUst6;
                    Command.Parameters.Add("Ust7", OleDbType.VarWChar).Value = strUst7;
                    Command.Parameters.Add("Ust8", OleDbType.VarWChar).Value = strUst8;
                    Command.Parameters.Add("Ust9", OleDbType.VarWChar).Value = strUst9;
                    Command.Parameters.Add("Ust10", OleDbType.VarWChar).Value = strUst10;
                    Command.Parameters.Add("Ust11", OleDbType.VarWChar).Value = strUst11;
                    Command.Parameters.Add("ScaleEnd", OleDbType.VarWChar).Value = strScaleEnd;
                    Command.Parameters.Add("Hist", OleDbType.VarWChar).Value = strHist;
                    Command.Parameters.Add("Delta", OleDbType.VarWChar).Value = strDelta;
                    Command.Parameters.Add("LimSpd", OleDbType.VarWChar).Value = strLimSpd;
                    Command.Parameters.Add("AdressBeginning", OleDbType.Integer).Value = iAdressBeginning;
                    Command.Parameters.Add("Place", OleDbType.Integer).Value = iPlace;
                    Command.Parameters.Add("PLC", OleDbType.Integer).Value = iPlc;
                    CommandString = "insert into OIP (ParamName, ParamID, Unit, ScaleBeginning, Ust0, Ust1, Ust2, Ust3, Ust4, Ust5, Ust6, Ust7, Ust8, Ust9, Ust10, Ust11, ScaleEnd, Hist, Delta, LimSpd, AdressBeginning, Place, PLC) " + "values ('" + strParamName + "', '" + strParamId + "', '" + strUnit + "', '" + 
                        strScaleBeginning + "', '" + strUst0 + "', '" + strUst1 + "', '" + strUst2 + "', '" + 
                        strUst3 + "', '" + strUst4 + "', '" + strUst5 + "', '" + strUst6 + "', '" + 
                        strUst7 + "', '" + strUst8 + "', '" + strUst9 + "', '" + strUst10 + "', '" + 
                        strUst11 + "', '" + strScaleEnd + "', '" + strHist + "', '" + strDelta + "', '" + 
                        strLimSpd + "', " + Convert.ToString(iAdressBeginning) + ", " + Convert.ToString(iPlace) + ", " + Convert.ToString(iPlc) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в PlaceOIP
        /// </summary>
        /// <param name="iPlace"></param>
        /// <param name="strPlaceName"></param>
        /// <returns></returns>
        public bool WriteOperMess_PlaceOIP(uint iPlace, string strPlaceName)
        {
            CommandString = "";
            ClearMessErr();
            if (strPlaceName != null)
            {
                try
                {
                    CreateDbCommand("insert into PlaceOIP (Place, PlaceName) values (?,?)");
                    Command.Parameters.Add("Place", OleDbType.Integer).Value = iPlace;
                    Command.Parameters.Add("PlaceName", OleDbType.VarWChar).Value = strPlaceName;
                    CommandString = "insert into PlaceOIP (Place, PlaceName) " + "values (" + Convert.ToString(iPlace) + ", '" + strPlaceName + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в PlaceTimer
        /// </summary>
        /// <param name="iPlace"></param>
        /// <param name="strPlaceName"></param>
        /// <returns></returns>
        public bool WriteOperMess_PlaceTimer(uint iPlace, string strPlaceName)
        {
            CommandString = "";
            ClearMessErr();
            if (strPlaceName != null)
            {
                try
                {
                    CreateDbCommand("insert into PlaceTimer (Place, PlaceName) values (?,?)");
                    Command.Parameters.Add("Place", OleDbType.Integer).Value = iPlace;
                    Command.Parameters.Add("PlaceName", OleDbType.VarWChar).Value = strPlaceName;
                    CommandString = "insert into PlaceTimer (Place, PlaceName) " + "values (" + Convert.ToString(iPlace) + ", '" + strPlaceName + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        private static string GetValue(string input)
        {
            if (input.Length > 0)
            {
                // замена двойных кавычек на одинарные
                if ((input[0] == '"') && (input[input.Length - 1] == '"'))
                    return "'" + input.Substring(1, input.Length - 2) + "'";
            }
            else
                return "null";
            return input;
        }

        /// <summary>
        /// Запись в таблицу
        /// </summary>
        /// <param name="input"></param>
        /// <param name="rowWithId"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public bool WriteOperMess_Table(string input, bool rowWithId, string table)
        {
            var word = input.Split(';');
            var i = 0;
            if (rowWithId) i = 1;  // пропускать поле ID
            CommandString = "";
            ClearMessErr();
            var fields = "";
            try
            {
                // подготовка строки
                CommandString =
                    "DECLARE @TABLE_NAME NVARCHAR(MAX) = '" + table + "' " +
                    "DECLARE @COLUMN_COUNT INT = (SELECT COUNT(*) FROM SYS.ALL_COLUMNS WHERE OBJECT_ID = (SELECT OBJECT_ID FROM SYS.OBJECTS WHERE NAME = @TABLE_NAME)) " +
                    "DECLARE @COLUMN_ID INT = 1 " +
                    "DECLARE @STR_Command NVARCHAR(MAX) = 'INSERT INTO ' + @TABLE_NAME + ' (' " +
                    "WHILE (@COLUMN_ID <= @COLUMN_COUNT) BEGIN " +
                    "  IF ((SELECT NAME FROM SYS.ALL_COLUMNS WHERE(OBJECT_ID = (SELECT OBJECT_ID FROM SYS.OBJECTS WHERE NAME = @TABLE_NAME)) AND(COLUMN_ID = @COLUMN_ID)) <> 'ID') " +
                    "  IF(@COLUMN_ID < @COLUMN_COUNT) " +
                    "    SET @STR_Command = @STR_Command + (SELECT NAME FROM SYS.ALL_COLUMNS WHERE (OBJECT_ID = (SELECT OBJECT_ID FROM SYS.OBJECTS WHERE NAME = @TABLE_NAME)) AND(COLUMN_ID = @COLUMN_ID)) +', ' " +
                    "  ELSE " +
                    "    SET @STR_Command = @STR_Command + (SELECT NAME FROM SYS.ALL_COLUMNS WHERE (OBJECT_ID = (SELECT OBJECT_ID FROM SYS.OBJECTS WHERE NAME = @TABLE_NAME)) AND(COLUMN_ID = @COLUMN_ID)) " +
                    "  SET @COLUMN_ID = @COLUMN_ID + 1 " +
                    "END " +
                    "SET @STR_Command = @STR_Command + ') ' " +
                    "SELECT @STR_Command 'Command'";
                CreateDbCommand(CommandString);
                ExecuteReader();
                if (Reader.Read())
                    CommandString = Convert.ToString(Reader["Command"]);
                Reader.Close();
                // подготовка значений
                while (i < word.Length)
                {
                    if (!string.IsNullOrEmpty(word[i]))
                        if (i < word.Length - 1)
                            fields += GetValue(word[i]) + ",";
                        else
                            fields += GetValue(word[i]);
                    i++;
                }
                CommandString = CommandString + " VALUES (" + fields + ")";
                // выполнение вставки значений
                CreateDbCommand(CommandString);
                Command.ExecuteNonQuery();
                return true;
            }
            catch (Exception exception)
            {
                MessErr = exception.Message;
            }
            return false;
        }

        /// <summary>
        /// Запись в PLCMessage
        /// </summary>
        /// <param name="iIdMess"></param>
        /// <param name="strUserName"></param>
        /// <param name="strPlace"></param>
        /// <param name="strDTime"></param>
        /// <param name="strDTimeAck"></param>
        /// <param name="iSysId"></param>
        /// <param name="iSysNum"></param>
        /// <param name="iMess"></param>
        /// <param name="strMessage"></param>
        /// <param name="iIsAck"></param>
        /// <param name="iPriority"></param>
        /// <param name="strValue"></param>
        /// <returns></returns>
        public bool WriteOperMess_PLCMessage(int iIdMess, string strUserName, string strPlace, string strDTime, string strDTimeAck, uint iSysId, uint iSysNum, uint iMess, string strMessage, uint iIsAck, uint iPriority, string strValue)
        {
            CommandString = "";
            ClearMessErr();
            if ((strPlace != null) && (strDTime != null) && (strDTimeAck != null) && (strMessage != null) && (strValue != null))
            {
                try
                {
                    if (!string.IsNullOrEmpty(strUserName))
                    {
                        CreateDbCommand("insert into PLCMessage (IDmess, UserName, Place, DTime, DTimeAck, SysID, SysNum, Mess, Message, IsAck, Priority, Value) " + "values (?,?,?,?,?,?,?,?,?,?,?,?)");
                        Command.Parameters.Add("IDmess", OleDbType.Integer).Value = iIdMess;
                        Command.Parameters.Add("UserName", OleDbType.VarWChar).Value = strUserName;
                        Command.Parameters.Add("Place", OleDbType.VarWChar).Value = strPlace;
                        Command.Parameters.Add("DTime", OleDbType.VarWChar).Value = strDTime;
                        Command.Parameters.Add("DTimeAck", OleDbType.VarWChar).Value = strDTimeAck;
                        Command.Parameters.Add("SysID", OleDbType.Integer).Value = iSysId;
                        Command.Parameters.Add("SysNum", OleDbType.Integer).Value = iSysNum;
                        Command.Parameters.Add("Mess", OleDbType.Integer).Value = iMess;
                        Command.Parameters.Add("Message", OleDbType.VarWChar).Value = strMessage;
                        Command.Parameters.Add("IsAck", OleDbType.Integer).Value = iIsAck;
                        Command.Parameters.Add("Priority", OleDbType.Integer).Value = iPriority;
                        Command.Parameters.Add("Value", OleDbType.VarWChar).Value = strValue;
                        CommandString = "insert into PLCMessage (IDmess, UserName, Place, DTime, DTimeAck, SysID, SysNum, Mess, Message, IsAck, Priority, Value) " + "values (" + Convert.ToString(iIdMess) + ", '" + strUserName + "', '" + strPlace + "', '" + strDTime + "', '" + strDTimeAck + "', " + Convert.ToString(iSysId) + ", " + Convert.ToString(iSysNum) + ", " + Convert.ToString(iMess) + ", '" + strMessage + "', " + Convert.ToString(iIsAck) + ", " + Convert.ToString(iPriority) + ", '" + strValue + "')";
                    }
                    else
                    {
                        CreateDbCommand("insert into PLCMessage (IDmess, Place, DTime, DTimeAck, SysID, SysNum, Mess, Message, IsAck, Priority, Value) " + "values (?,?,?,?,?,?,?,?,?,?,?)");
                        Command.Parameters.Add("IDmess", OleDbType.Integer).Value = iIdMess;
                        Command.Parameters.Add("Place", OleDbType.VarWChar).Value = strPlace;
                        Command.Parameters.Add("DTime", OleDbType.VarWChar).Value = strDTime;
                        Command.Parameters.Add("DTimeAck", OleDbType.VarWChar).Value = strDTimeAck;
                        Command.Parameters.Add("SysID", OleDbType.Integer).Value = iSysId;
                        Command.Parameters.Add("SysNum", OleDbType.Integer).Value = iSysNum;
                        Command.Parameters.Add("Mess", OleDbType.Integer).Value = iMess;
                        Command.Parameters.Add("Message", OleDbType.VarWChar).Value = strMessage;
                        Command.Parameters.Add("IsAck", OleDbType.Integer).Value = iIsAck;
                        Command.Parameters.Add("Priority", OleDbType.Integer).Value = iPriority;
                        Command.Parameters.Add("Value", OleDbType.VarWChar).Value = strValue;
                        CommandString = "insert into PLCMessage (IDmess, Place, DTime, DTimeAck, SysID, SysNum, Mess, Message, IsAck, Priority, Value) " + "values (" + Convert.ToString(iIdMess) + "', '" + strPlace + "', '" + strDTime + "', '" + strDTimeAck + "', " + Convert.ToString(iSysId) + ", " + Convert.ToString(iSysNum) + ", " + Convert.ToString(iMess) + ", '" + strMessage + "', " + Convert.ToString(iIsAck) + ", " + Convert.ToString(iPriority) + ", '" + strValue + "')";
                    }
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Priority
        /// </summary>
        /// <param name="iPriority"></param>
        /// <param name="strNamePriority"></param>
        /// <returns></returns>
        public bool WriteOperMess_Priority(uint iPriority, string strNamePriority)
        {
            CommandString = "";
            ClearMessErr();
            if (strNamePriority != null)
            {
                try
                {
                    CreateDbCommand("insert into Priority (Priority, NamePriority) values (?,?)");
                    Command.Parameters.Add("Priority", OleDbType.Integer).Value = iPriority;
                    Command.Parameters.Add("NamePriority", OleDbType.VarWChar).Value = strNamePriority;
                    CommandString = "insert into Priority (Priority, NamePriority) " + "values (" + Convert.ToString(iPriority) + ", '" + strNamePriority + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Registration
        /// </summary>
        /// <param name="strLogin"></param>
        /// <param name="strUserName"></param>
        /// <param name="strUserGroup"></param>
        /// <param name="iUserLevel"></param>
        /// <returns></returns>
        public bool WriteOperMess_Registration(string strLogin, string strUserName, string strUserGroup, uint iUserLevel)
        {
            CommandString = "";
            ClearMessErr();
            if ((strLogin != null) && (strUserName != null) && (strUserGroup != null))
            {
                try
                {
                    CreateDbCommand("insert into Registration (Login, UserName, UserGroup, UserLevel) values (?,?,?,?)");
                    Command.Parameters.Add("Login", OleDbType.VarWChar).Value = strLogin;
                    Command.Parameters.Add("UserName", OleDbType.VarWChar).Value = strUserName;
                    Command.Parameters.Add("UserGroup", OleDbType.VarWChar).Value = strUserGroup;
                    Command.Parameters.Add("UserLevel", OleDbType.Integer).Value = iUserLevel;
                    CommandString = "insert into Registration (Login, UserName, UserGroup, UserLevel) " + "values ('" + strLogin + "', '" + strUserName + "', '" + strUserGroup + "', " + Convert.ToString(iUserLevel) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Sound
        /// </summary>
        /// <param name="iSound"></param>
        /// <param name="strName"></param>
        /// <returns></returns>
        public bool WriteOperMess_Sound(uint iSound, string strName)
        {
            CommandString = "";
            ClearMessErr();
            if (strName != null)
            {
                try
                {
                    CreateDbCommand("insert into Sound (Sound, Name) values (?,?)");
                    Command.Parameters.Add("Sound", OleDbType.Integer).Value = iSound;
                    Command.Parameters.Add("Name", OleDbType.VarWChar).Value = strName;
                    CommandString = "insert into Sound (Sound, Name) " + "values (" + Convert.ToString(iSound) + ", '" + strName + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в SysObject
        /// </summary>
        /// <param name="iSysId"></param>
        /// <param name="strObject"></param>
        /// <param name="strComment"></param>
        /// <returns></returns>
        public bool WriteOperMess_SysObject(uint iSysId, string strObject, string strComment)
        {
            CommandString = "";
            ClearMessErr();
            if ((strObject != null) && (strComment != null))
            {
                try
                {
                    CreateDbCommand("insert into SysObject (SysID, Object, Comment) values (?,?,?)");
                    Command.Parameters.Add("SysID", OleDbType.Integer).Value = iSysId;
                    Command.Parameters.Add("Object", OleDbType.VarWChar).Value = strObject;
                    Command.Parameters.Add("Comment", OleDbType.VarWChar).Value = strComment;
                    CommandString = "insert into SysObject (SysID, Object, Comment) " + "values (" + Convert.ToString(iSysId) + ", '" + strObject + "', '" + strComment + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Timer
        /// </summary>
        /// <param name="strParamName"></param>
        /// <param name="strParamId"></param>
        /// <param name="strUnit"></param>
        /// <param name="strValue"></param>
        /// <param name="iAdressBeginning"></param>
        /// <param name="iPlace"></param>
        /// <returns></returns>
        public bool WriteOperMess_Timer(string strParamName, string strParamId, string strUnit, string strValue, uint iAdressBeginning, uint iPlace)
        {
            CommandString = "";
            ClearMessErr();
            if ((strParamName != null) && (strParamId != null) && (strUnit != null) && (strValue != null))
            {
                try
                {
                    CreateDbCommand("insert into Timer (ParamName, ParamID, Unit, Value, AdressBeginning, Place) values (?,?,?,?,?,?)");
                    Command.Parameters.Add("ParamName", OleDbType.VarWChar).Value = strParamName;
                    Command.Parameters.Add("ParamID", OleDbType.VarWChar).Value = strParamId;
                    Command.Parameters.Add("Unit", OleDbType.VarWChar).Value = strUnit;
                    Command.Parameters.Add("Value", OleDbType.VarWChar).Value = strValue;
                    Command.Parameters.Add("AdressBeginning", OleDbType.Integer).Value = iAdressBeginning;
                    Command.Parameters.Add("Place", OleDbType.Integer).Value = iPlace;
                    CommandString = "insert into Timer (ParamName, ParamID, Unit, Value, AdressBeginning, Place) " + "values ('" + strParamName + "', '" + strParamId + "', '" + strUnit + "', '" + strValue + "', " + Convert.ToString(iAdressBeginning) + ", " + Convert.ToString(iPlace) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Timer
        /// </summary>
        /// <param name="strParamName"></param>
        /// <param name="strParamId"></param>
        /// <param name="strUnit"></param>
        /// <param name="strValue"></param>
        /// <param name="iAdressBeginning"></param>
        /// <param name="iPlace"></param>
        /// <param name="iPlc"></param>
        /// <returns></returns>
        public bool WriteOperMess_Timer(string strParamName, string strParamId, string strUnit, string strValue, uint iAdressBeginning, uint iPlace, uint iPlc)
        {
            CommandString = "";
            ClearMessErr();
            if ((strParamName != null) && (strParamId != null) && (strUnit != null) && (strValue != null))
            {
                try
                {
                    CreateDbCommand("insert into Timer (ParamName, ParamID, Unit, Value, AdressBeginning, Place, PLC) values (?,?,?,?,?,?,?)");
                    Command.Parameters.Add("ParamName", OleDbType.VarWChar).Value = strParamName;
                    Command.Parameters.Add("ParamID", OleDbType.VarWChar).Value = strParamId;
                    Command.Parameters.Add("Unit", OleDbType.VarWChar).Value = strUnit;
                    Command.Parameters.Add("Value", OleDbType.VarWChar).Value = strValue;
                    Command.Parameters.Add("AdressBeginning", OleDbType.Integer).Value = iAdressBeginning;
                    Command.Parameters.Add("Place", OleDbType.Integer).Value = iPlace;
                    Command.Parameters.Add("PLC", OleDbType.Integer).Value = iPlc;
                    CommandString = "insert into Timer (ParamName, ParamID, Unit, Value, AdressBeginning, Place, PLC) " + "values ('" + strParamName + "', '" + strParamId + "', '" + strUnit + "', '" + strValue + "', " + Convert.ToString(iAdressBeginning) + ", " + Convert.ToString(iPlace) + ", " + Convert.ToString(iPlc) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в Type
        /// </summary>
        /// <param name="iType"></param>
        /// <param name="strName"></param>
        /// <returns></returns>
        public bool WriteOperMess_Type(uint iType, string strName)
        {
            CommandString = "";
            ClearMessErr();
            if (strName != null)
            {
                try
                {
                    CreateDbCommand("insert into Type (Type, Name) values (?,?)");
                    Command.Parameters.Add("Type", OleDbType.Integer).Value = iType;
                    Command.Parameters.Add("Name", OleDbType.VarWChar).Value = strName;
                    CommandString = "insert into Type (Type, Name) " + "values (" + Convert.ToString(iType) + ", '" + strName + "')";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Запись в UserGroups
        /// </summary>
        /// <param name="iGroupId"></param>
        /// <param name="strUserGroup"></param>
        /// <param name="iUserLevel"></param>
        /// <param name="iGroupTime"></param>
        /// <param name="iIsActive"></param>
        /// <returns></returns>
        public bool WriteOperMess_UserGroups(uint iGroupId, string strUserGroup, uint iUserLevel, uint iGroupTime, uint iIsActive)
        {
            CommandString = "";
            ClearMessErr();
            if (strUserGroup != null)
            {
                try
                {
                    CreateDbCommand("insert into UserGroups (GroupID, UserGroup, UserLevel, GroupTime, isActive) values (?,?,?,?,?)");
                    Command.Parameters.Add("GroupID", OleDbType.Integer).Value = iGroupId;
                    Command.Parameters.Add("UserGroup", OleDbType.VarWChar).Value = strUserGroup;
                    Command.Parameters.Add("UserLevel", OleDbType.Integer).Value = iUserLevel;
                    Command.Parameters.Add("GroupTime", OleDbType.Integer).Value = iGroupTime;
                    Command.Parameters.Add("isActive", OleDbType.Integer).Value = iIsActive;
                    CommandString = "insert into UserGroups (GroupID, UserGroup, UserLevel, GroupTime, isActive) " + "values (" + Convert.ToString(iGroupId) + ", '" + strUserGroup + "', " + Convert.ToString(iUserLevel) + ", " + Convert.ToString(iGroupTime) + ", " + Convert.ToString(iIsActive) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        /// <summary>
        /// Расширенная запись в UserGroups
        /// </summary>
        /// <param name="iGroupId"></param>
        /// <param name="strUserGroup"></param>
        /// <param name="iUserLevel"></param>
        /// <param name="iGroupTime"></param>
        /// <param name="iIsActive"></param>
        /// <param name="iPasswordTime"></param>
        /// <param name="iMinChars"></param>
        /// <param name="iDayLife"></param>
        /// <returns></returns>
        public bool WriteOperMess_UserGroups_2(uint iGroupId, string strUserGroup, uint iUserLevel, uint iGroupTime, uint iIsActive, uint iPasswordTime, uint iMinChars, uint iDayLife)
        {
            CommandString = "";
            ClearMessErr();
            if (strUserGroup != null)
            {
                try
                {
                    CreateDbCommand("insert into UserGroups (GroupID, UserGroup, UserLevel, GroupTime, isActive, PasswordTime, MinChars, DayLife) values (?,?,?,?,?,?,?,?)");
                    Command.Parameters.Add("GroupID", OleDbType.Integer).Value = iGroupId;
                    Command.Parameters.Add("UserGroup", OleDbType.VarWChar).Value = strUserGroup;
                    Command.Parameters.Add("UserLevel", OleDbType.Integer).Value = iUserLevel;
                    Command.Parameters.Add("GroupTime", OleDbType.Integer).Value = iGroupTime;
                    Command.Parameters.Add("isActive", OleDbType.Integer).Value = iIsActive;
                    Command.Parameters.Add("PasswordTime", OleDbType.Integer).Value = iPasswordTime;
                    Command.Parameters.Add("MinChars", OleDbType.Integer).Value = iMinChars;
                    Command.Parameters.Add("DayLife", OleDbType.Integer).Value = iDayLife;
                    CommandString = "insert into UserGroups (GroupID, UserGroup, UserLevel, GroupTime, isActive, PasswordTime, MinChars, DayLife) " + "values (" + Convert.ToString(iGroupId) + ", '" + strUserGroup + "', " + Convert.ToString(iUserLevel) + ", " + Convert.ToString(iGroupTime) + ", " + Convert.ToString(iIsActive) + ", " + Convert.ToString(iPasswordTime) + ", " + Convert.ToString(iMinChars) + ", " + Convert.ToString(iDayLife) + ")";
                    Command.ExecuteNonQuery();
                    return true;
                }
                catch (Exception exception)
                {
                    MessErr = exception.Message;
                }
            }
            return false;
        }

        #endregion
     }
}
