using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataGenerator
{
    
    static class Program
    {
        public static SettingsWrapper Settings; //создание ссылки на настройки

        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Settings = new SettingsWrapper(); //загрузка настроек
            SettingsWrapper.FullSettingsPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\Settings.xml";
            Settings = Settings.Load(); //загрузка настроек
            Settings.Save();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
