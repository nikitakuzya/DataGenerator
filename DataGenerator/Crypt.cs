using System;
using System.Windows.Forms;

namespace DataGenerator
{
    /// <summary>
    /// "Библиотека шифрования" (Crypt.dll)
    /// </summary>
    public static class Crypt
    {
        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public static string ExceptionMsg;

        // https://msdn.microsoft.com/en-us/library/bb311038.aspx
        /// <summary>
        /// строка как HEX
        /// </summary>
        /// <param name="strInput">входная строка</param>
        /// <returns></returns>
        public static string StringAsHex(string strInput)
        {
            ExceptionMsg = null;
            string strOutput = null;
            try
            {
                if (strInput != null)
                {
                    if (!string.IsNullOrWhiteSpace(strInput))
                    {
                        char[] values = strInput.ToCharArray();
                        // ReSharper disable once LoopCanBeConvertedToQuery
                        foreach (char letter in values)
                        {
                            int value = Convert.ToInt32(letter);
                            strOutput += $"{value:X}";
                        }
                    }
                    else
                        strOutput = "";
                }
                else
                    ExceptionMsg = "Входное значение должно отличаться от null";
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
            return strOutput;
        }

        // https://msdn.microsoft.com/en-us/library/bb311038.aspx
        /// <summary>
        /// строка HEX как строка
        /// </summary>
        /// <param name="strInput">входная строка</param>
        /// <returns></returns>
        public static string HexAsString(string strInput)
        {
            ExceptionMsg = null;
            var strOutput = "";
            try
            {
                if (strInput != null)
                {
                    if (!string.IsNullOrWhiteSpace(strInput))
                    {
                        for (var i = 0; i < strInput.Length; i += 2)
                        {
                            var iValue =
                                Convert.ToInt32(Convert.ToString((strInput[i]) + Convert.ToString((strInput[i + 1]))),
                                    16);
                            // string stringValue = Char.ConvertFromUtf32(value);
                            // char charValue = (char)value;
                            // Console.WriteLine("hexadecimal value = {0}, int value = {1}, char value = {2} or {3}", hex, value, stringValue, charValue);
                            strOutput += (char) iValue;
                        }
                    }
                    else
                        strOutput = "";
                }
                else
                    ExceptionMsg = "Входное значение должно отличаться от null";
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
            return strOutput;
        }

        /// <summary>
        /// Получить пароль
        /// </summary>
        /// <param name="strInput">входная строка</param>
        /// <returns></returns>
        public static string GetPassword(string strInput)
        {
            return SimpleWithOffset(HexAsString(strInput), false, 20);
        }

        /// <summary>
        /// Задать пароль
        /// </summary>
        /// <param name="strInput">входная строка</param>
        /// <returns></returns>
        public static string SetPassword(string strInput)
        {
            return StringAsHex(SimpleWithOffset(strInput, true, 20));
        }

        /// <summary>
        /// Шифровать/расшифровать текст с ключом со смещением
        /// </summary>
        /// <param name="strInput">входная строка</param>
        /// <param name="bCrypt">шифровать/расшифровать</param>
        /// <param name="key">ключ</param>
        /// <returns></returns>
        public static string SimpleWithOffset(string strInput, bool bCrypt, int key)
        {
            ExceptionMsg = null;
            var strResult = "";
            try
            {
                if ((key >= 0) && (key < 255))
                {
                    int code;
                    if (bCrypt) // дешифрование
                    {
                        foreach (var t in strInput)
                        {
                            code = Convert.ToByte(t); // код символа
                            code = code - key; // расшифрование
                            if (code < 1)
                            {
                                code = code + 255; // если итоговый код выходит за пределы таблицы ASCII
                            }
                            strResult = strResult + Convert.ToChar(code); // итоговый текст
                            if (key > 0) // смещение
                            {
                                key = key - 1;
                            }
                        }
                    }
                    else // шифрование
                    {
                        foreach (var t in strInput)
                        {
                            code = Convert.ToByte(t); // код символа
                            code = code + key; // шифрование
                            if (code > 255)
                            {
                                code = code - 255; // если итоговый код выходит за пределы таблицы ASCII
                            }
                            strResult = strResult + Convert.ToChar(code); // итоговый текст
                            if (key > 0) // смещение
                            {
                                key = key - 1;
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
            return strResult;
        }

        /// <summary>
        /// Шифровать/расшифровать текст с ключом без смещения
        /// </summary>
        /// <param name="strInput">входная строка</param>
        /// <param name="bCrypt">шифровать/расшифровать</param>
        /// <param name="key">ключ</param>
        /// <returns></returns>
        public static string SimpleWithoutOffset(string strInput, bool bCrypt, int key)
        {
            ExceptionMsg = "";
            string strResult = "";
            try
            {
                if ((key >= 0) && (key < 255))
                {
                    int code;
                    if (bCrypt) // дешифрование
                    {
                        foreach (char t in strInput)
                        {
                            code = Convert.ToByte(t); // код символа
                            code = code - key; // расшифрование
                            if (code < 1)
                            {
                                code = code + 255; // если итоговый код выходит за пределы таблицы ASCII
                            }
                            strResult = strResult + Convert.ToChar(code); // итоговый текст
                        }
                    }
                    else // шифрование
                    {
                        foreach (char t in strInput)
                        {
                            code = Convert.ToByte(t); // код символа
                            code = code + key; // шифрование
                            if (code > 255)
                            {
                                code = code - 255; // если итоговый код выходит за пределы таблицы ASCII
                            }
                            strResult = strResult + Convert.ToChar(code); // итоговый текст
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                ExceptionMsg = exception.Message;
            }
            return strResult;
        }

        /// <summary>
        /// Шифровать/расшифровать текст с ключом = 20 (пароль)
        /// </summary>
        /// <param name="bCrypt">шифровать/расшифровать</param>
        /// <param name="richTextBox1">входное поле с текстом</param>
        /// <param name="richTextBox2">выходное поле с текстом</param>
        public static void SetPasswordWrapp(bool bCrypt, ref RichTextBox richTextBox1, ref RichTextBox richTextBox2)
        {
            ExceptionMsg = "";
            if (bCrypt)
            {
                try
                {
                    richTextBox2.Text = SetPassword(richTextBox1.Text);
                }
                catch (Exception exception)
                {
                    richTextBox2.Text = "";
                    MessageBox.Show(@"Ошибка шифрования!" + Environment.NewLine + Environment.NewLine + exception.Message);
                    ExceptionMsg = exception.Message;
                }
            }
            else
            {
                try
                {
                    richTextBox1.Text = GetPassword(richTextBox2.Text);
                }
                catch (Exception exception)
                {
                    richTextBox1.Text = "";
                    MessageBox.Show(@"Ошибка расшифрования!" + Environment.NewLine + Environment.NewLine + exception.Message);
                    ExceptionMsg = exception.Message;
                }
            }
        }

        /// <summary>
        /// Шифровать/расшифровать текст с ключом
        /// </summary>
        /// <param name="bCrypt">шифровать/расшифровать</param>
        /// <param name="bKey"></param>
        /// <param name="keySource"></param>
        /// <param name="richTextBox1"></param>
        /// <param name="richTextBox2"></param>
        /// <param name="bOffset"></param>
        public static void SimpleWrapp(bool bCrypt, bool bKey, decimal keySource, ref RichTextBox richTextBox1, ref RichTextBox richTextBox2, bool bOffset)
        {
            ExceptionMsg = "";
            if (bCrypt)
            {
                if (bKey)
                {
                    Int32 key = 0;
                    try
                    {
                        key = Convert.ToInt32(keySource);
                    }
                    catch (Exception exception)
                    {
                        richTextBox2.Text = "";
                        MessageBox.Show(@"Ошибка ключа!" + Environment.NewLine + Environment.NewLine + @"Должно быть число!");
                        ExceptionMsg = exception.Message;
                    }
                    try
                    {
                        richTextBox2.Text = StringAsHex(bOffset 
                            ? SimpleWithOffset(richTextBox1.Text, true, key) 
                            : SimpleWithoutOffset(richTextBox1.Text, true, key));
                    }
                    catch (Exception exception)
                    {
                        richTextBox2.Text = "";
                        MessageBox.Show(@"Ошибка шифрования!" + Environment.NewLine + Environment.NewLine + exception.Message);
                        ExceptionMsg = exception.Message;
                    }
                }
                else
                {
                    try
                    {
                        richTextBox2.Text = StringAsHex(bOffset 
                            ? SimpleWithOffset(richTextBox1.Text, true, 0) 
                            : SimpleWithoutOffset(richTextBox1.Text, true, 0));
                    }
                    catch (Exception exception)
                    {
                        richTextBox2.Text = "";
                        MessageBox.Show(@"Ошибка шифрования!" + Environment.NewLine + Environment.NewLine + exception.Message);
                        ExceptionMsg = exception.Message;
                    }
                }
            }
            else
            {
                if (bKey)
                {
                    Int32 key = 0;
                    try
                    {
                        key = Convert.ToInt32(keySource);
                    }
                    catch (Exception exception)
                    {
                        richTextBox1.Text = "";
                        MessageBox.Show(@"Ошибка ключа!" + Environment.NewLine + Environment.NewLine + @"Должно быть число!");
                        ExceptionMsg = exception.Message;
                    }
                    try
                    {
                        richTextBox1.Text = bOffset 
                            ? SimpleWithOffset(HexAsString(richTextBox2.Text), false, key) 
                            : SimpleWithoutOffset(HexAsString(richTextBox2.Text), false, key);
                    }
                    catch (Exception exception)
                    {
                        richTextBox1.Text = "";
                        MessageBox.Show(@"Ошибка расшифрования!" + Environment.NewLine + Environment.NewLine + exception.Message);
                        ExceptionMsg = exception.Message;
                    }
                }
                else
                {
                    try
                    {
                        richTextBox1.Text = bOffset 
                            ? SimpleWithOffset(HexAsString(richTextBox2.Text), false, 0) 
                            : SimpleWithoutOffset(HexAsString(richTextBox2.Text), false, 0);
                    }
                    catch (Exception exception)
                    {
                        richTextBox1.Text = "";
                        MessageBox.Show(@"Ошибка расшифрования!" + Environment.NewLine + Environment.NewLine + exception.Message);
                        ExceptionMsg = exception.Message;
                    }
                }
            }
        }

        /// <summary>
        /// Преобразовать в HEX
        /// </summary>
        /// <param name="bConvert"></param>
        /// <param name="richTextBox1"></param>
        /// <param name="richTextBox2"></param>
        public static void ConvertToHexWrapp(bool bConvert, ref RichTextBox richTextBox1, ref RichTextBox richTextBox2)
        {
            ExceptionMsg = "";
            if (bConvert)
            {
                try
                {
                    richTextBox2.Text = StringAsHex(richTextBox1.Text);
                }
                catch (Exception exception)
                {
                    richTextBox2.Text = "";
                    MessageBox.Show(exception.Message);
                    ExceptionMsg = exception.Message;
                }
            }
            else
            {
                try
                {
                    richTextBox1.Text = HexAsString(richTextBox2.Text);
                }
                catch (Exception exception)
                {
                    richTextBox1.Text = "";
                    MessageBox.Show(exception.Message);
                    ExceptionMsg = exception.Message;
                }
            }
        }

        /// <summary>
        /// Шифрование методом Виженера
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public static string CodeVizner(string password)
        {
            try
            {
                string key = "vtyzpjdenbujhm";
                //все символы, которые могут быть использованы при вводе пароля
                const string all = @"`1234567890-=~!@#$%^&*()_+qwertyuiop[]QWERTYUIOP{}asdfghjkl'\ASDFGHJKL:""|ZXCVBNM<>?zxcvbnm,./№ёЁйцукенгшщзхъЙЦУКЕНГШЩЗХЪфывапролджэФЫВАПРОЛДЖЭячсмитьбюЯЧСМИТЬБЮ";
                string cPass = "";
                if (key.Length > password.Length)               //если длина строки пароля (ключа для входа в программу и для шифрования)>длины строки пароля (какого-либо сайта и т.д.),
                {
                    key = key.Substring(0, password.Length);    //то переменная key обрежется и станет равной длинне пароля 
                }
                else                                            // Иначе повторять ключ (ключключключклю), пока не станет равным длинне пароля
                    for (int i = 0; key.Length < password.Length; i++)
                    {
                        key = key + key.Substring(i, 1);
                    }
                // основной цикл шифрования
                for (int i = 0; i < password.Length; i++)
                {  //находим центр строки all (центр - это будущий первый символ строки со сдвигом)
                    var center = all.IndexOf(key.Substring(i, 1), StringComparison.Ordinal);
                    var leftSlice = all.Substring(center);  // левый срез
                    var rightSlice = all.Substring(0, center);  // правый срез
                    var st = leftSlice + rightSlice; // строка со сдвигом
                    center = all.IndexOf(password.Substring(i, 1), StringComparison.Ordinal);// теперь в переменную center запишем индекс очередного символа шифруемой строки
                    cPass += st.Substring(center, 1);    //поскольку индексы символа из строки со сдвигом и из обычной строки совпадают, то нужный нам символ берется по такому же индексу
                }

                return cPass;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Расшифровка
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public static string DecodeVizner(string password)
        {
            try
            {
                string key = "vtyzpjdenbujhm";
                // строка all содержит все символы, которые можно вводить с русской и англ раскладки клавиатуры
                const string all = @"`1234567890-=~!@#$%^&*()_+qwertyuiop[]QWERTYUIOP{}asdfghjkl'\ASDFGHJKL:""|ZXCVBNM<>?zxcvbnm,./№ёЁйцукенгшщзхъЙЦУКЕНГШЩЗХЪфывапролджэФЫВАПРОЛДЖЭячсмитьбюЯЧСМИТЬБЮ";
                // строка st со сдвигом по ключу (в качестве ключа используем наш пароль для входа)
                string cPass = "";
                // если пароль короче ключа - обрезаем ключ
                if (key.Length > password.Length)
                {
                    key = key.Substring(0, password.Length);
                }
                // Иначе повторяем ключ, пока он не примет длинну пароля
                else
                    for (int i = 0; key.Length < password.Length; i++)
                    {
                        key = key + key.Substring(i, 1);
                    }
                // основной цикл расшифрования
                // ReSharper disable once LoopCanBeConvertedToQuery
                for (int i = 0; i < password.Length; i++)
                {
                    //находим центр строки all (центр - это будущий первый символ строки со сдвигом)
                    var center = all.IndexOf(key.Substring(i, 1), StringComparison.Ordinal);
                    var leftSlice = all.Substring(center);  // левый срез
                    var rightSlice = all.Substring(0, center);  // правый срез
                    var st = leftSlice + rightSlice; // строка со сдвигом
                    center = st.IndexOf(password.Substring(i, 1), StringComparison.Ordinal); // теперь в переменную center запишем индекс очередного символа расшифроввываемой строки
                    cPass += all.Substring(center, 1); //поскольку индексы символа из строки со сдвигом и из обычной строки совпадают, то нужный нам символ берется по такому же индексу
                }
                return cPass; //возвращаем расшифрованный пароль.
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

    }
}
