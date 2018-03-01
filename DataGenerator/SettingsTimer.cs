using System;

namespace DataGenerator
{
    public static class SettingsTimer
    {
        public static void ReadTimerSetting()
        {
            string[] adressTimer = Properties.Settings.Default.TimerAdressSettings.Split(';');
            TwZashNps = Convert.ToInt32(adressTimer[0]);
            TwZashNa = Convert.ToInt32(adressTimer[1]);
            TwZashRp = Convert.ToInt32(adressTimer[2]);
            TwList4 = Convert.ToInt32(adressTimer[3]);
            TwNa = Convert.ToInt32(adressTimer[4]);
            TwCommTimes = Convert.ToInt32(adressTimer[5]);
            TwZdv = Convert.ToInt32(adressTimer[6]);
            TwVs = Convert.ToInt32(adressTimer[7]);
            TwAsuptZoneF = Convert.ToInt32(adressTimer[8]);
            TwAsuptZoneW = Convert.ToInt32(adressTimer[9]);

            string[] listTimer = Properties.Settings.Default.TimerListSettings.Split(';');
            TwZashNpsList = listTimer[0];
            TwZashNaList = listTimer[1];
            TwZashRpList = listTimer[2];
            TwList4List = listTimer[3];
            TwNaList = listTimer[4];
            TwCommTimesList = listTimer[5];
            TwZdvList = listTimer[6];
            TwVsList = listTimer[7];
            TwAsuptZoneListF = listTimer[8];
            TwAsuptZoneListW = listTimer[9];

            string[] strTimer = Properties.Settings.Default.TimerStrSettings.Split(';');
            TwZashNpsStr = Convert.ToInt32(strTimer[0]);
            TwZashNaStr = Convert.ToInt32(strTimer[1]);
            TwZashRpStr = Convert.ToInt32(strTimer[2]);
            TwList4Str = Convert.ToInt32(strTimer[3]);
            TwNaStr = Convert.ToInt32(strTimer[4]);
            TwCommTimesStr = Convert.ToInt32(strTimer[5]);
            TwZdvStr = Convert.ToInt32(strTimer[6]);
            TwVsStr = Convert.ToInt32(strTimer[7]);
            TwAsuptZoneStrF = Convert.ToInt32(strTimer[8]);
            TwAsuptZoneStrW = Convert.ToInt32(strTimer[9]);

         
        }

        public static void WriteTimerSetting()
        {
            string adressTimer = TwZashNps + ";" + TwZashNa + ";"+ TwZashRp + ";" + TwList4 + ";" + TwNa + ";" 
                + TwCommTimes + ";" + TwZdv + ";" + TwVs + ";" + TwAsuptZoneF + ";" + TwAsuptZoneW + ";";
            string listTimer = TwZashNpsList + ";" + TwZashNaList + ";" + TwZashRpList + ";" + TwList4List + ";" + TwNaList + ";"
                + TwCommTimesList + ";" + TwZdvList + ";" + TwVsList + ";" + TwAsuptZoneListF + ";" + TwAsuptZoneListW + ";";
            string strTimer = TwZashNpsStr + ";" + TwZashNaStr + ";" + TwZashRpStr + ";" + TwList4Str + ";" + TwNaStr + ";"
                + TwCommTimesStr + ";" + TwZdvStr + ";" + TwVsStr + ";" + TwAsuptZoneStrF + ";" + TwAsuptZoneStrW + ";";

            Properties.Settings.Default.TimerAdressSettings = adressTimer;
            Properties.Settings.Default.TimerListSettings = listTimer;
            Properties.Settings.Default.TimerStrSettings = strTimer;

        }
        //адреса
        public static int TwZashNps { get; set; }
        public static int TwZashNa { get; set; }
        public static int TwZashRp { get; set; }
        public static int TwList4 { get; set; }
        public static int TwNa { get; set; }
        public static int TwCommTimes { get; set; }
        public static int TwZdv { get; set; }
        public static int TwVs { get; set; }
        public static int TwAsuptZoneF { get; set; }
        public static int TwAsuptZoneW { get; set; }

        //листы
        public static string TwZashNpsList { get; set; }
        public static string TwZashNaList { get; set; }
        public static string TwZashRpList { get; set; }
        public static string TwList4List { get; set; }
        public static string TwNaList { get; set; }
        public static string TwCommTimesList { get; set; }
        public static string TwZdvList { get; set; }
        public static string TwVsList { get; set; }
        public static string TwAsuptZoneListF { get; set; }
        public static string TwAsuptZoneListW { get; set; }

        //нач строки
        public static int TwZashNpsStr { get; set; }
        public static int TwZashNaStr { get; set; }
        public static int TwZashRpStr { get; set; }
        public static int TwList4Str { get; set; }
        public static int TwNaStr { get; set; }
        public static int TwCommTimesStr { get; set; }
        public static int TwZdvStr { get; set; }
        public static int TwVsStr { get; set; }
        public static int TwAsuptZoneStrF { get; set; }
        public static int TwAsuptZoneStrW { get; set; }

        //кол-во элементов(защит и т.д.)
        public static int Quntity_twCommTimes { get; set; }
    }
}
