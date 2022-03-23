using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.FileOperates
{
    public class FileInfoHelper
    {
        public static string GetProgramVersion(string filePath )
        {
            FileVersionInfo fileVersionInfo =  FileVersionInfo.GetVersionInfo(filePath);

            return fileVersionInfo.FileVersion;
        }
    }
}
