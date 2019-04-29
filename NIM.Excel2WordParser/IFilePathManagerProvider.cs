using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator
{
    public static class IFilePathManagerProvider
    {

        public static FilePathManager PathProvider
        {
            get; set;

        }

        public static string GetWordResultPath(this FilePathManager filePathManager, string excelFileFullName)
        {
            var str = filePathManager.WordTemplateConfPath;
            var excelFileDirectory = new System.IO.DirectoryInfo(excelFileFullName);
            var wordResultPath = excelFileDirectory.Parent.FullName;
            if (!System.IO.Directory.Exists(wordResultPath))
                System.IO.Directory.CreateDirectory(wordResultPath);
            return wordResultPath;



        }
    }
}
