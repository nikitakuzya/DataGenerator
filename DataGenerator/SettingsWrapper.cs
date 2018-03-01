using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Serialization;
using Microsoft.Win32;

namespace DataGenerator
{
    [XmlRoot, Serializable]
    public class SettingsWrapper : INotifyPropertyChanged
    {
        public SettingsWrapper()
        {
        }
        
        [XmlIgnore]
        public static string FullSettingsPath;
        [XmlIgnore]
        public static Exception LastException;

        public void SettingsPath(string programPath = "", string fullSettingsPath = "")
        {
            var defaultpath = "C:\\Active\\DataGenerator\\Settings.xml";
            try
            {
                if (FullSettingsPath != "")
                {
                    FullSettingsPath = fullSettingsPath;
                }
                else
                {
                    string activePath;
                    RegistryKey regKey64 =
                        RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                            .OpenSubKey("SOFTWARE\\Active\\NPS\\v1.0\\Paths");
                    if (regKey64 != null) activePath = (string) regKey64.GetValue("Project");
                    else activePath = "C:\\Active\\";
                    if (programPath != "")
                    {
                        FullSettingsPath = activePath + programPath + "\\Settings.xml";
                    }
                }
                if (!Directory.Exists(FullSettingsPath)) FullSettingsPath = defaultpath;           
            }
            catch (Exception)
            {
                FullSettingsPath = defaultpath;
            }
        }

        private SettingsWrapper NewSettings()
        {
            return new SettingsWrapper
            {
                OIPNumSettings = "3; 2; 3; 47; 16; 17; 18; 19; 20; 13300; 41; 4; 5; 6; 7; 8; 9; 10; 11; 12; 13; 14; 15; 100; 50; 11064; 55;",
                TimerAdressSettings = "8000;8096;9632;9760;9792;10112;10144;10400;9792;9792;",
                TimerListSettings = "Станц Защ;Защ НА;Защ Рез;twList4;twNA;twCommTimes;twZDV;twVS;twASUPTzoneFoam;twASUPTzoneWater;",
                TimerStrSettings = "4;4;4;2;2;2;2;2;2;2;",
                Gen_OIP_NoGap = false,
                DB_version = 9,
                DB_useRegSettings = true,
                DB_usePasswordCrypt = true,
                DB_Provider = "SQLOLEDB",
                DB_DataSource = @"localhost\SQLEXPRESS",
                DB_InitialCatalog = "OperMess",
                DB_UserID = "sa",
                DB_Password = "sa",
                Gen_OIP_NoGapPlaceValue = 1,
                startRow_ZDV = 4,
                startRow_VS =3,
                startRow_StZash =4,
                startRow_ZNA =4,
                Quntity_twCommTimes = 20,
                UseAdrBeginning = true

            };
        }

        #region Настройки

        private string _OIPNumSettings;

        [XmlElement]
        public string OIPNumSettings
        {
            get { return _OIPNumSettings; }
            set
            {
                if (_OIPNumSettings == value) return;
                _OIPNumSettings = value;
            }
        }

        private string _TimerAdressSettings;

        [XmlElement]
        public string TimerAdressSettings
        {
            get { return _TimerAdressSettings; }
            set
            {
                if (_TimerAdressSettings == value) return;
                _TimerAdressSettings = value;
            }
        }

        private string _TimerListSettings;

        [XmlElement]
        public string TimerListSettings
        {
            get { return _TimerListSettings; }
            set
            {
                if (_TimerListSettings == value) return;
                _TimerListSettings = value;
            }
        }

        
        private string _TimerStrSettings;

        [XmlElement]
        public string TimerStrSettings
        {
            get { return _TimerStrSettings; }
            set
            {
                if (_TimerStrSettings == value) return;
                _TimerStrSettings = value;
            }
        }
        
        private bool _Gen_OIP_NoGap;

        [XmlElement]
        public bool Gen_OIP_NoGap
        {
            get { return _Gen_OIP_NoGap; }
            set
            {
                if (_Gen_OIP_NoGap == value) return;
                _Gen_OIP_NoGap = value;
            }
        }
        
        private int _DB_version;

        [XmlElement]
        public int DB_version
        {
            get { return _DB_version; }
            set
            {
                if (_DB_version == value) return;
                _DB_version = value;
            }
        }


        private bool _DB_useRegSettings;

        [XmlElement]
        public bool DB_useRegSettings
        {
            get { return _DB_useRegSettings; }
            set
            {
                if (_DB_useRegSettings == value) return;
                _DB_useRegSettings = value;
            }
        }

        private bool _DB_usePasswordCrypt;

        [XmlElement]
        public bool DB_usePasswordCrypt
        {
            get { return _DB_usePasswordCrypt; }
            set
            {
                if (_DB_usePasswordCrypt == value) return;
                _DB_usePasswordCrypt = value;
            }
        }

        private string _DB_Provider;

        [XmlElement]
        public string DB_Provider
        {
            get { return _DB_Provider; }
            set
            {
                if (_DB_Provider == value) return;
                _DB_Provider = value;
            }
        }

        private string _DB_DataSource;

        [XmlElement]
        public string DB_DataSource
        {
            get { return _DB_DataSource; }
            set
            {
                if (_DB_DataSource == value) return;
                _DB_DataSource = value;
            }
        }

        private string _DB_InitialCatalog;

        [XmlElement]
        public string DB_InitialCatalog
        {
            get { return _DB_InitialCatalog; }
            set
            {
                if (_DB_InitialCatalog == value) return;
                _DB_InitialCatalog = value;
            }
        }

        private string _DB_UserID;

        [XmlElement]
        public string DB_UserID
        {
            get { return _DB_UserID; }
            set
            {
                if (_DB_UserID == value) return;
                _DB_UserID = value;
            }
        }

        private string _DB_Password;

        [XmlElement]
        public string DB_Password
        {
            get { return _DB_Password; }
            set
            {
                if (_DB_Password == value) return;
                _DB_Password = value;
            }
        }

        private int _Gen_OIP_NoGapPlaceValue;

        [XmlElement]
        public int Gen_OIP_NoGapPlaceValue
        {
            get { return _Gen_OIP_NoGapPlaceValue; }
            set
            {
                if (_Gen_OIP_NoGapPlaceValue == value) return;
                _Gen_OIP_NoGapPlaceValue = value;
            }
        }

        private int _startRow_ZDV;

        [XmlElement]
        public int startRow_ZDV
        {
            get { return _startRow_ZDV; }
            set
            {
                if (_startRow_ZDV == value) return;
                _startRow_ZDV = value;
            }
        }
        private int _startRow_VS;

        [XmlElement]
        public int startRow_VS
        {
            get { return _startRow_VS; }
            set
            {
                if (_startRow_VS == value) return;
                _startRow_VS = value;
            }
        }
        private int _startRow_StZash;

        [XmlElement]
        public int startRow_StZash
        {
            get { return _startRow_StZash; }
            set
            {
                if (_startRow_StZash == value) return;
                _startRow_StZash = value;
            }
        }
       private int _startRow_ZNA;

        [XmlElement]
        public int startRow_ZNA
        {
            get { return _startRow_ZNA; }
            set
            {
                if (_startRow_ZNA == value) return;
                _startRow_ZNA = value;
            }
        }

        private int _Quntity_twCommTimes;
        [XmlElement]
        public int Quntity_twCommTimes
        {
            get { return _Quntity_twCommTimes; }
            set
            {
                if (_Quntity_twCommTimes == value) return;
                _Quntity_twCommTimes = value;
            }
        }
        private bool _UseAdrBeginning;
        [XmlElement]
        public bool UseAdrBeginning
        {
            get { return _UseAdrBeginning; }
            set
            {
                if (_UseAdrBeginning == value) return;
                _UseAdrBeginning = value;
            }
        }


        //-----------------------------------------------
        #endregion Настройки

        #region Действия с настройками

        public bool Save()
        {
            try
            {
                //var xs = new XmlSerializer(typeof (SettingsWrapper));
                var xs = XmlSerializer.FromTypes(new[] { typeof(SettingsWrapper) })[0];
                using (var writer = new StreamWriter(FullSettingsPath))
                    xs.Serialize(writer, this);
                return true;
            }
            catch (Exception exception)
            {
                LastException = exception;
                return false;
            }
        }

        public SettingsWrapper Load()
        {
            try
            {           
                //var xs = new XmlSerializer(typeof (SettingsWrapper));
                var xs = XmlSerializer.FromTypes(new[] { typeof(SettingsWrapper) })[0];
                using (var fileStream = new StreamReader(FullSettingsPath))
                    return (SettingsWrapper) xs.Deserialize(fileStream);
            }
            catch (Exception exception)
            {
                var newsettings = NewSettings();
                newsettings.Save();
                LastException = exception;
                return newsettings;
            }
        }

        public bool Export(string exportPath)
        {
            try
            {
                var xs = new XmlSerializer(typeof (SettingsWrapper));
                using (var writer = new StreamWriter(exportPath))
                    xs.Serialize(writer, this);
                return true;
            }
            catch (Exception exception)
            {
                LastException = exception;
                return false;
            }

        }

        public bool Import(string importPath)
        {
            try
            {
                var xs = new XmlSerializer(typeof (SettingsWrapper));
                using (var fileStream = new StreamReader(importPath))
                {
                    var newsettings = (SettingsWrapper) xs.Deserialize(fileStream);
                    newsettings.Save();
                }
                return true;
            }
            catch (Exception exception)
            {
                LastException = exception;
                return false;
            }
        }

        public bool RestoreDefaults()
        {
            try
            {
                var newsettings = NewSettings();
                newsettings.Save();
                return true;
            }
            catch (Exception exception)
            {
                LastException = exception;
                return false;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged( string propertyName = null)
        //private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

    }
}
