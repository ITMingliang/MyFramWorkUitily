using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.Format
{
    public static class Serialization
    {
        #region json
        public static T JsonToObject<T>(string jsonText)
        {
            DataContractJsonSerializer dataContractJsonSerializer = new DataContractJsonSerializer(typeof(T));
            MemoryStream memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonText));
            T result = (T)((object)dataContractJsonSerializer.ReadObject(memoryStream));
            memoryStream.Dispose();
            return result;
        }
        public static string ObjectToJSON<T>(T obj)
        {
            DataContractJsonSerializer dataContractJsonSerializer = new DataContractJsonSerializer(typeof(T));
            string result = string.Empty;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                dataContractJsonSerializer.WriteObject(memoryStream, obj);
                memoryStream.Position = 0L;
                using (StreamReader streamReader = new StreamReader(memoryStream))
                {
                    result = streamReader.ReadToEnd();
                }
            }
            return result;
        }
        #endregion

        #region byte操作
        public static string ByteToHex(byte[] Bytes)
        {
            string str = string.Empty;
            foreach (byte Byte in Bytes)
            {
                str += String.Format("{0:X2}", Byte) + " ";
            }
            return str.Trim();
        }

        public static byte[] strToToHexByte(string hexString)
        {
            try
            {
                hexString = hexString.Replace(" ", "");
                if ((hexString.Length % 2) != 0)
                    hexString += " ";
                byte[] returnBytes = new byte[hexString.Length / 2];
                for (int i = 0; i < returnBytes.Length; i++)
                    returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
                return returnBytes;
            }
            catch
            {
                return null;
            }

        }
        public static byte[] strToHexByte(string hexString)
        {
            try
            {
                hexString = hexString.Replace(" ", "");
                if ((hexString.Length % 2) != 0)
                    hexString += " ";
                byte[] returnBytes = new byte[hexString.Length / 2];
                for (int i = 0; i < returnBytes.Length; i++)
                {
                    returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);

                }
                return returnBytes;


            }
            catch 
            {
                return null;
            }

        }
        #endregion
    }
}
