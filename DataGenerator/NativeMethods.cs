using System.Runtime.InteropServices;
using System.Text;

namespace DataGenerator
{
    public class NativeMethods
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode)] // ���������� kernel32.dll � ��������� ��� ������� WritePrivateProfilesString
        public static extern int WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)] // ��� ��� ���������� kernel32.dll, � ������ ��������� ������� GetPrivateProfileString
        public static extern int GetPrivateProfileString(string section, string key, string Default, StringBuilder retVal, int size, string filePath);

    }
}