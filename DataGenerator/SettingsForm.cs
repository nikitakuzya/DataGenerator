using System;
using System.Windows.Forms;
using DataGenerator.Properties;
using Microsoft.Win32;
using static DataGenerator.Properties.Settings;
using static DataGenerator.SettingsOip;
using static DataGenerator.SettingsTimer;

namespace DataGenerator
{
    public partial class SettingsForm : Form
    {
        readonly MainForm _mainForm;
        OleDb _oleDb;
        public SettingsForm(MainForm mainForm)
        {
            Program.Settings.Load();
            InitializeComponent();

            _mainForm = mainForm;

            numericUpDown_dbVersion.Value = Program.Settings.DB_version;
            checkBox_OIP_NoGap.Checked = Program.Settings.Gen_OIP_NoGap;
            radioButton_db_UseRegSettings.Checked = Program.Settings.DB_useRegSettings;
            radioButton_db_UseManualSettings.Checked = !radioButton_db_UseRegSettings.Checked;
            numericUpDown_Gen_OIP_NoGapPlaceValue.Value = Program.Settings.Gen_OIP_NoGapPlaceValue;

            textBox_db_Provider.Text = Program.Settings.DB_Provider;
            textBox_db_IniCatalog.Text = Program.Settings.DB_InitialCatalog;
            textBox_db_DataSource.Text = Program.Settings.DB_DataSource;
            textBox_db_UserID.Text = Program.Settings.DB_UserID;
            textBox_db_Password.Text = Program.Settings.DB_Password;
            checkBox_db_UsePasswordCrypt.Checked = Program.Settings.DB_usePasswordCrypt;
            panel_db_manualSettings.Enabled = radioButton_db_UseManualSettings.Checked;

            numericUpDown1.Value = Program.Settings.startRow_ZDV;
            numericUpDown2.Value = Program.Settings.startRow_VS;
            numericUpDown3.Value = Program.Settings.startRow_StZash;
            numericUpDown4.Value = Program.Settings.startRow_ZNA;
            numericUpDown_Quntity_twCommTimes.Value = Program.Settings.Quntity_twCommTimes;
            checkBox_UseAdrBeg.Checked = Program.Settings.UseAdrBeginning;

            ReadOipNumSetting();
            ReadTimerSetting();
            

            numericUpDown_Oip_StartRow.Value = OipStartRow;
            numericUpDown_Oip_ParamName.Value = OipParamName;
            numericUpDown_Oip_ParamID.Value = OipParamId;
            numericUpDown_Oip_Unit.Value = OipUnit;
            numericUpDown_Oip_ScaleBeginning.Value = OipScaleBeginning;
            numericUpDown_Oip_ScaleEnd.Value = OipScaleEnd;
            numericUpDown_Oip_Hist.Value = OipHist;
            numericUpDown_Oip_Delta.Value = OipDelta;
            numericUpDown_Oip_LimSpd.Value = OipLimSpd;
            numericUpDown_Oip_AdressBeginning.Value = OipAdressBeginning;
            numericUpDown_Oip_PlaceOld.Value = OipPlaceOld;
            numericUpDown_Oip_Place.Value = OipPlace;
            numericUpDown_Oip_Ust0.Value = OipUst0;
            numericUpDown_Oip_Ust1.Value = OipUst1;
            numericUpDown_Oip_Ust2.Value = OipUst2;
            numericUpDown_Oip_Ust3.Value = OipUst3;
            numericUpDown_Oip_Ust4.Value = OipUst4;
            numericUpDown_Oip_Ust5.Value = OipUst5;
            numericUpDown_Oip_Ust6.Value = OipUst6;
            numericUpDown_Oip_Ust7.Value = OipUst7;
            numericUpDown_Oip_Ust8.Value = OipUst8;
            numericUpDown_Oip_Ust9.Value = OipUst9;
            numericUpDown_Oip_Ust10.Value = OipUst10;
            numericUpDown_Oip_Ust11.Value = OipUst11;
            numericUpDown_Oip_PLC.Value = OipPlc;
            numericUpDown_OIPType.Value = OipType;
            numericUpDown_Oip_AdressBeginningKF.Value = OipAdressBeginningKf;

            numericUpDown_twZashNPS.Value = TwZashNps;
            numericUpDown_twZashNA.Value = TwZashNa;
            numericUpDown_twZashRP.Value = TwZashRp;
            numericUpDown_twList4.Value = TwList4;
            numericUpDown_twNA.Value = TwNa;
            numericUpDown_twCommTimes.Value = TwCommTimes;
            numericUpDown_twZDV.Value = TwZdv;
            numericUpDown_twVS.Value = TwVs;
            numericUpDown_twASUPTzoneF.Value = TwAsuptZoneF;
            numericUpDown_twASUPTzoneW.Value = TwAsuptZoneW;

            textBox_List_twZashNPS.Text = TwZashNpsList;
            textBox_List_twZashNA.Text = TwZashNaList;
            textBox_List_twZashRP.Text = TwZashRpList;
            textBox_List_twList4.Text = TwList4List;
            textBox_List_twNA.Text = TwNaList;
            textBox_List_twCommTimes.Text = TwCommTimesList;
            textBox_List_twZDV.Text = TwZdvList;
            textBox_List_twVS.Text = TwVsList;
            textBox_List_twASUPTzoneF.Text = TwAsuptZoneListF;
            textBox_List_twASUPTzoneW.Text = TwAsuptZoneListW;

            numericUpDown_Start_twZashNPS.Value = TwZashNpsStr;
            numericUpDown_Start_twZashNA.Value = TwZashNaStr;
            numericUpDown_Start_twZashRP.Value = TwZashRpStr;
            numericUpDown_Start_twList4.Value = TwList4Str;
            numericUpDown_Start_twNA.Value = TwNaStr;
            numericUpDown_Start_twCommTimes.Value = TwCommTimesStr;
            numericUpDown_Start_twZDV.Value = TwZdvStr;
            numericUpDown_Start_twVS.Value = TwVsStr;
            numericUpDown_Start_twASUPTzoneF.Value = TwAsuptZoneStrF;
            numericUpDown_Start_twASUPTzoneW.Value = TwAsuptZoneStrW;
        }

        private void button_Save_Click(object sender, EventArgs e)
        {
            Program.Settings.DB_version = (int) numericUpDown_dbVersion.Value;
            Program.Settings.Gen_OIP_NoGap = checkBox_OIP_NoGap.Checked;
            Program.Settings.DB_useRegSettings = radioButton_db_UseRegSettings.Checked;
            Program.Settings.DB_Provider = textBox_db_Provider.Text;
            Program.Settings.DB_InitialCatalog = textBox_db_IniCatalog.Text;
            Program.Settings.DB_DataSource = textBox_db_DataSource.Text;
            Program.Settings.DB_UserID = textBox_db_UserID.Text;
            Program.Settings.DB_Password = textBox_db_Password.Text;
            Program.Settings.DB_usePasswordCrypt = checkBox_db_UsePasswordCrypt.Checked;
            Program.Settings.Gen_OIP_NoGapPlaceValue = (int) numericUpDown_Gen_OIP_NoGapPlaceValue.Value;
            Program.Settings.UseAdrBeginning = checkBox_UseAdrBeg.Checked;

            OipStartRow = (int) numericUpDown_Oip_StartRow.Value;
            OipParamName = (int) numericUpDown_Oip_ParamName.Value;
            OipParamId = (int) numericUpDown_Oip_ParamID.Value;
            OipUnit = (int) numericUpDown_Oip_Unit.Value;
            OipScaleBeginning = (int) numericUpDown_Oip_ScaleBeginning.Value;
            OipScaleEnd = (int) numericUpDown_Oip_ScaleEnd.Value;
            OipHist = (int)numericUpDown_Oip_Hist.Value;
            OipDelta = (int)numericUpDown_Oip_Delta.Value;
            OipLimSpd = (int)numericUpDown_Oip_LimSpd.Value;
            OipAdressBeginning = (int)numericUpDown_Oip_AdressBeginning.Value;
            OipPlaceOld = (int)numericUpDown_Oip_PlaceOld.Value;
            OipPlace = (int)numericUpDown_Oip_Place.Value;
            OipUst0 = (int)numericUpDown_Oip_Ust0.Value;
            OipUst1 = (int)numericUpDown_Oip_Ust1.Value;
            OipUst2 = (int)numericUpDown_Oip_Ust2.Value;
            OipUst3 = (int)numericUpDown_Oip_Ust3.Value;
            OipUst4 = (int)numericUpDown_Oip_Ust4.Value;
            OipUst5 = (int)numericUpDown_Oip_Ust5.Value;
            OipUst6 = (int)numericUpDown_Oip_Ust6.Value;
            OipUst7 = (int)numericUpDown_Oip_Ust7.Value;
            OipUst8 = (int)numericUpDown_Oip_Ust8.Value;
            OipUst9 = (int)numericUpDown_Oip_Ust9.Value;
            OipUst10 = (int)numericUpDown_Oip_Ust10.Value;
            OipUst11 = (int)numericUpDown_Oip_Ust11.Value;
            OipPlc = (int)numericUpDown_Oip_PLC.Value;
            OipType = (int)numericUpDown_OIPType.Value;
            OipAdressBeginningKf = (int) numericUpDown_Oip_AdressBeginningKF.Value;

            WriteOipNumSetting();

            TwZashNps = (int)numericUpDown_twZashNPS.Value;
            TwZashNa = (int)numericUpDown_twZashNA.Value;
            TwZashRp = (int)numericUpDown_twZashRP.Value;
            TwList4 = (int)numericUpDown_twList4.Value;
            TwNa = (int)numericUpDown_twNA.Value;
            TwCommTimes = (int)numericUpDown_twCommTimes.Value;
            TwZdv = (int)numericUpDown_twZDV.Value;
            TwVs = (int)numericUpDown_twVS.Value;
            TwAsuptZoneF = (int)numericUpDown_twASUPTzoneF.Value;
            TwAsuptZoneW = (int)numericUpDown_twASUPTzoneW.Value;

            TwZashNpsList = textBox_List_twZashNPS.Text;
            TwZashNaList = textBox_List_twZashNA.Text;
            TwZashRpList = textBox_List_twZashRP.Text;
            TwList4List = textBox_List_twList4.Text;
            TwNaList = textBox_List_twNA.Text;
            TwCommTimesList = textBox_List_twCommTimes.Text;
            TwZdvList = textBox_List_twZDV.Text;
            TwVsList = textBox_List_twVS.Text;
            TwAsuptZoneListF = textBox_List_twASUPTzoneF.Text;
            TwAsuptZoneListW = textBox_List_twASUPTzoneW.Text;

            TwZashNpsStr = (int)numericUpDown_Start_twZashNPS.Value;
            TwZashNaStr = (int)numericUpDown_Start_twZashNA.Value;
            TwZashRpStr = (int)numericUpDown_Start_twZashRP.Value;
            TwList4Str = (int)numericUpDown_Start_twList4.Value;
            TwNaStr = (int)numericUpDown_Start_twNA.Value;
            TwCommTimesStr = (int)numericUpDown_Start_twCommTimes.Value;
            TwZdvStr = (int)numericUpDown_Start_twZDV.Value;
            TwVsStr = (int)numericUpDown_Start_twVS.Value;
            TwAsuptZoneStrF = (int)numericUpDown_Start_twASUPTzoneF.Value;
            TwAsuptZoneStrW = (int)numericUpDown_Start_twASUPTzoneW.Value;

            WriteTimerSetting();
   
            Program.Settings.startRow_ZDV = (int) numericUpDown1.Value;
            Program.Settings.startRow_VS = (int)numericUpDown2.Value;
            Program.Settings.startRow_StZash = (int)numericUpDown3.Value;
            Program.Settings.startRow_ZNA = (int)numericUpDown4.Value;
            Program.Settings.Quntity_twCommTimes = (int)numericUpDown_Quntity_twCommTimes.Value;

            Default.Save();
            Program.Settings.Save();
            Close();
        }

        private void button_Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_CheckDBversion_Click(object sender, EventArgs e)
        {
            try
            {
                _oleDb = new OleDb();
                _oleDb.Open();
                _oleDb.CreateDbCommand("SELECT TOP(1)* FROM [DbVersion] ORDER BY ID DESC");
                _oleDb.ExecuteReader();
                _oleDb.Reader.Read();
                button_CheckDBversion.Text = Convert.ToString(_oleDb.Reader["Build"]);
                _oleDb.Close();
            }
            catch (Exception exception)
            {
                button_CheckDBversion.Text = @"Ошибка";
                _mainForm.toolStripStatusLabel1.Text = @"Ошибка чтения версии базы " + exception.Message;
                if (_oleDb != null) _mainForm.toolStripStatusLabel1.Text += @" SQL: " + _oleDb.MessErr;
            }
            finally
            {
                if (_oleDb != null && _oleDb.SqlConnected()) _oleDb.Close();
            }
        }

        private void radioButton_db_UseRegSettings_CheckedChanged(object sender, EventArgs e)
        {
            panel_db_manualSettings.Enabled = radioButton_db_UseManualSettings.Checked;
        }

        private void radioButton_db_UseManualSettings_CheckedChanged(object sender, EventArgs e)
        {
            panel_db_manualSettings.Enabled = radioButton_db_UseManualSettings.Checked;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Program.Settings.startRow_ZDV = Convert.ToInt32(numericUpDown1.Value);
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            Program.Settings.startRow_ZDV = Convert.ToInt32(numericUpDown2.Value);
        }
        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            Program.Settings.startRow_StZash = Convert.ToInt32(numericUpDown3.Value);
        }
        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            Program.Settings.startRow_ZNA = Convert.ToInt32(numericUpDown4.Value);
        }


        private void numericUpDown_Quntity_twCommTimes_ValueChanged(object sender, EventArgs e)
        {
            Program.Settings.Quntity_twCommTimes = Convert.ToInt32(numericUpDown_Quntity_twCommTimes.Value);
        }

        private void button_Read_from_Reg_Click(object sender, EventArgs e)
        {
            
                RegistryKey regKey32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Active\\NPS\\v1.0\\SettingDB");
                RegistryKey regKey64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Active\\NPS\\v1.0\\SettingDB");

                if (regKey64 != null)
                {
                    textBox_db_Provider.Text = (string)regKey64.GetValue("Provider");
                    textBox_db_IniCatalog.Text = (string)regKey64.GetValue("InitialCatalog");
                    textBox_db_DataSource.Text = (string)regKey64.GetValue("DataSource");
                    textBox_db_UserID.Text = (string)regKey64.GetValue("UserId");
                    textBox_db_Password.Text = "";
                    //textBox_db_Password.Text  = Program.Settings.DB_usePasswordCrypt
                    //    ? Crypt.GetPassword((string)regKey64.GetValue("PasswordCrypt"))
                    //    : (string)regKey64.GetValue("Password");
                    checkBox_db_UsePasswordCrypt.Checked = Program.Settings.DB_usePasswordCrypt;
                    panel_db_manualSettings.Enabled = radioButton_db_UseManualSettings.Checked;
                   
                }
                else if (regKey32 != null)
                {
                textBox_db_Provider.Text = (string)regKey32.GetValue("Provider");
                textBox_db_IniCatalog.Text = (string)regKey32.GetValue("InitialCatalog");
                textBox_db_DataSource.Text = (string)regKey32.GetValue("DataSource");
                textBox_db_UserID.Text = (string)regKey32.GetValue("UserId");
                textBox_db_Password.Text = "";
                //textBox_db_Password.Text = Program.Settings.DB_usePasswordCrypt
                //    ? Crypt.GetPassword((string)regKey32.GetValue("PasswordCrypt"))
                //    : (string)regKey32.GetValue("Password");
                checkBox_db_UsePasswordCrypt.Checked = Program.Settings.DB_usePasswordCrypt;
                panel_db_manualSettings.Enabled = radioButton_db_UseManualSettings.Checked;

            }
            else
                {
                textBox_db_Provider.Text = "SQLOLEDB";
                textBox_db_IniCatalog.Text = "OperMess";
                textBox_db_DataSource.Text = @"localhost\SQLEXPRESS";
                textBox_db_UserID.Text = "sa";
                textBox_db_Password.Text = "";
                checkBox_db_UsePasswordCrypt.Checked = false;
                panel_db_manualSettings.Enabled = true;
                throw new Exception("Не найдены настройки соединения с базой данных в реестре");
                }

        }

      
    }
}
