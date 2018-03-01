using System.Runtime.InteropServices;
using System.Text;

namespace DataGenerator
{
    public class NativeMethods
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode)] // Подключаем kernel32.dll и описываем его функцию WritePrivateProfilesString
        public static extern int WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)] // Еще раз подключаем kernel32.dll, а теперь описываем функцию GetPrivateProfileString
        public static extern int GetPrivateProfileString(string section, string key, string Default, StringBuilder retVal, int size, string filePath);

    }
}