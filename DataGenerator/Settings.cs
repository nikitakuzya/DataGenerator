using System;
using System.IO;
using System.Text;

namespace DataGenerator
{
    /// <summary>
    /// Работа с параметрами приложения из INI файла
    /// </summary>
    class Settings
    {
        readonly string _path; //Имя файла.

        /// <summary>
        /// Файл настроек
        /// </summary>
        /// <param name="iniPath">Путь</param>
        public Settings(string iniPath)
        {
            _path = new FileInfo(iniPath).FullName;
        }

        /// <summary>
        /// Читаем ini-файл и возвращаем значение указного ключа из заданной секции
        /// </summary>
        /// <param name="section">Cекция</param>
        /// <param name="key">Ключ</param>
        /// <returns></returns>
        public string ReadIni(string section, string key)
        {
            var retVal = new StringBuilder(255);
            NativeMethods.GetPrivateProfileString(section, key, "error", retVal, 255, _path);
            if (retVal.ToString() == "error") throw new Exception("Ошибка чтения файла настроек в [" + section + "] " + key);
            return retVal.ToString();
        }

        /// <summary>
        /// Записываем в ini-файл. Запись происходит в выбранную секцию в выбранный ключ.
        /// </summary>
        /// <param name="section">Cекция</param>
        /// <param name="key">Ключ</param>
        /// <param name="value">Значение</param>
        public void Write(string section, string key, string value)
        {
            NativeMethods.WritePrivateProfileString(section, key, value, _path);
        }


        /// <summary>
        /// Удаляем ключ из выбранной секции.
        /// </summary>
        /// <param name="key">Ключ</param>
        /// <param name="section">Cекция</param>
        public void DeleteKey(string key, string section = null)
        {
            Write(section, key, null);
        }

        /// <summary>
        /// Удаляем выбранную секцию
        /// </summary>
        /// <param name="section">Cекция</param>
        public void DeleteSection(string section = null)
        {
            Write(section, null, null);
        }


        /// <summary>
        /// Проверяем, есть ли такой ключ, в этой секции
        /// </summary>
        /// <param name="key">Ключ</param>
        /// <param name="section">Cекция</param>
        /// <returns></returns>
        public bool KeyExists(string key, string section = null)
        {
            return ReadIni(section, key).Length > 0;
        }
    }
}
