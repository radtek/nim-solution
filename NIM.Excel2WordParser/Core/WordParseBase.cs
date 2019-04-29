using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.Core
{
    public abstract class WordParseBase
    {
        protected string mOriginalExcelFileName;
        protected FilePathManager FilePathManager;
        protected string ExcelDataFileFullName { get; private set; }

        public WordParseBase(string originalExcelFileFullName, string copyedExcelFileFullName, FilePathManager filePathManager)
        {
            this.mOriginalExcelFileName = originalExcelFileFullName;
            this.ExcelDataFileFullName = copyedExcelFileFullName;
            this.FilePathManager = filePathManager;
        }


        public abstract string GeneraterFile();

    }
}
