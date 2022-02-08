using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.FileOperates
{
	public class IniFile
	{
		private string m_FileName;

		public string FileName
		{
			get
			{
				return this.m_FileName;
			}
			set
			{
				this.m_FileName = value;
			}
		}

		[DllImport("kernel32.dll")]
		private static extern int GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

		[DllImport("kernel32.dll")]
		private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

		[DllImport("kernel32.dll")]
		private static extern int WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

		public IniFile(string aFileName)
		{
			this.m_FileName = aFileName;
		}

		public IniFile()
		{
		}

		public int ReadInt(string section, string name, int def)
		{
			return IniFile.GetPrivateProfileInt(section, name, def, this.m_FileName);
		}

		public string ReadString(string section, string name, string def)
		{
			StringBuilder stringBuilder = new StringBuilder(2048);
			IniFile.GetPrivateProfileString(section, name, def, stringBuilder, 2048, this.m_FileName);
			return stringBuilder.ToString();
		}

		public void WriteInt(string section, string name, int Ival)
		{
			IniFile.WritePrivateProfileString(section, name, Ival.ToString(), this.m_FileName);
		}

		public void WriteString(string section, string name, string strVal)
		{
			IniFile.WritePrivateProfileString(section, name, strVal, this.m_FileName);
		}

		public void DeleteSection(string section)
		{
			IniFile.WritePrivateProfileString(section, null, null, this.m_FileName);
		}

		public void DeleteAllSection()
		{
			IniFile.WritePrivateProfileString(null, null, null, this.m_FileName);
		}

		public string IniReadValue(string section, string name)
		{
			StringBuilder stringBuilder = new StringBuilder(256);
			IniFile.GetPrivateProfileString(section, name, "", stringBuilder, 256, this.m_FileName);
			return stringBuilder.ToString();
		}

		public void IniWriteValue(string section, string name, string value)
		{
			IniFile.WritePrivateProfileString(section, name, value, this.m_FileName);
		}
	}
}
