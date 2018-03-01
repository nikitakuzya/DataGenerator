using System;

namespace DataGenerator
{
    public static class SettingsOip
    {

        public static void ReadOipNumSetting()
        {
            string[] numOip = Properties.Settings.Default.OIPNumSettings.Split(';');
             OipStartRow  = Convert.ToInt32(numOip[0]);
             OipParamName  = Convert.ToInt32(numOip[1]);
             OipParamId  = Convert.ToInt32(numOip[2]);
             OipUnit  = Convert.ToInt32(numOip[3]);
             OipScaleBeginning  = Convert.ToInt32(numOip[4]);
             OipScaleEnd  = Convert.ToInt32(numOip[5]);
             OipHist  = Convert.ToInt32(numOip[6]);
             OipDelta  = Convert.ToInt32(numOip[7]);
             OipLimSpd  = Convert.ToInt32(numOip[8]);
             OipAdressBeginning  = Convert.ToInt32(numOip[9]);
             OipPlaceOld = Convert.ToInt32(numOip[10]);
             OipUst0  = Convert.ToInt32(numOip[11]);
             OipUst1  = Convert.ToInt32(numOip[12]);
             OipUst2  = Convert.ToInt32(numOip[13]);
             OipUst3  = Convert.ToInt32(numOip[14]);
             OipUst4  = Convert.ToInt32(numOip[15]);
             OipUst5  = Convert.ToInt32(numOip[16]);
             OipUst6  = Convert.ToInt32(numOip[17]);
             OipUst7  = Convert.ToInt32(numOip[18]);
             OipUst8  = Convert.ToInt32(numOip[19]);
             OipUst9  = Convert.ToInt32(numOip[20]);
             OipUst10  = Convert.ToInt32(numOip[21]);
             OipUst11  = Convert.ToInt32(numOip[22]);
             OipPlc  = Convert.ToInt32(numOip[23]);
             OipPlace = Convert.ToInt32(numOip[24]);
             OipAdressBeginningKf = Convert.ToInt32(numOip[25]);
             OipType = Convert.ToInt32(numOip[26]);
        }

        public static void WriteOipNumSetting()
        {
            string numOip = OipStartRow + ";" + OipParamName + ";" + OipParamId + ";" + OipUnit + ";" + OipScaleBeginning + ";" + 
                            OipScaleEnd + ";" + OipHist + ";" + OipDelta + ";" + OipLimSpd + ";" + OipAdressBeginning + ";" +
                            OipPlaceOld + ";" + OipUst0 + ";" + OipUst1 + ";" + OipUst2 + ";" + OipUst3 + ";" + OipUst4 + ";" +
                            OipUst5 + ";" + OipUst6 + ";" + OipUst7 + ";" + OipUst8 + ";" + OipUst9 + ";" + OipUst10 + ";" + 
                            OipUst11 + ";" + OipPlc + ";" + OipPlace + ";" + OipAdressBeginningKf + ";" + OipType + ";"; 
            Properties.Settings.Default.OIPNumSettings = numOip;
        }

        public static int OipStartRow { get; set; }
        public static int OipAdressBeginning { get; set; }
        public static int OipParamName { get; set; }
        public static int OipHist { get; set; }
        public static int OipDelta { get; set; }
        public static int OipParamId { get; set; }
        public static int OipUnit { get; set; }
        public static int OipScaleBeginning { get; set; }
        public static int OipUst0 { get; set; }
        public static int OipUst1 { get; set; }
        public static int OipUst2 { get; set; }
        public static int OipUst3 { get; set; }
        public static int OipUst4 { get; set; }
        public static int OipUst5 { get; set; }
        public static int OipUst6 { get; set; }
        public static int OipUst7 { get; set; }
        public static int OipUst8 { get; set; }
        public static int OipUst9 { get; set; }
        public static int OipUst10 { get; set; }
        public static int OipUst11 { get; set; }
        public static int OipScaleEnd { get; set; }
        public static int OipLimSpd { get; set; }
        public static int OipPlace { get; set; }
        public static int OipPlaceOld { get; set; }
        public static int OipPlc { get; set; }
        public static int OipAdressBeginningKf { get; set; }
        public static int OipType { get; set; }


    }
}
