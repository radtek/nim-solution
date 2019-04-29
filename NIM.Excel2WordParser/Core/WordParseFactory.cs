using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NIM.Utilty;

namespace NIM.CertificationGenerator.Core
{
    public class WordParseFactory
    {
        public static WordParseBase GetWordParse(string excelResultFileFullName)
        {
            var originalExcelFileFullName = excelResultFileFullName;
            //由于当前的EXCEL还在处于打开的状态，我们无法通过OPENXML打开这个EXCEL文件
            //所以我们将EXCEL复制到临时文件夹下面
            var filePathManager = IFilePathManagerProvider.PathProvider;
            var excelDataFileFullName = Path.Combine(filePathManager.TemplateFilesPath, Guid.NewGuid().ToString() + ".xlsx");
            File.Copy(excelResultFileFullName, excelDataFileFullName);

            var productAliasName = ""; //ProductName将始终来自于第一个sheet表中的B2列
            using (SpreadsheetDocument document =
              SpreadsheetDocument.Open(excelDataFileFullName, false))
            {
                var addressName = "B2";
                WorkbookPart wbPart = document.WorkbookPart;
                //查找对应的sheet ,它的规范是 从记录 或者记录1中的B2中去获取



                var sheets = wbPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList();


                var sheet = sheets.Where(t => t.Name == "记录").FirstOrDefault();
                if (sheet == null)
                    sheet = sheets.Where(t => t.Name == "记录1").FirstOrDefault();
                if (sheet == null)
                    throw new Exception("EXCEL中，找不到以`记录` 或者 `记录1`命名的sheet ");
                productAliasName = document.GetCellValue(sheet.Name, addressName);
                if (string.IsNullOrEmpty(productAliasName))
                    productAliasName = document.GetCellValue(sheet.Name, "C2");
            }

            var productName = ProductAliasName.GetProductName(productAliasName);

            if (productName == "辐射温度计")
                return new RadiationThermomater.WordParser(originalExcelFileFullName, excelDataFileFullName, filePathManager);
            else if (productName == "红外热像仪")
                return new ThermalImager.WordParse(originalExcelFileFullName, excelDataFileFullName, filePathManager);
            else if (productName == "黑体辐射源")
                return new BlackBodySingleRadiationiation.WordParser(originalExcelFileFullName, excelDataFileFullName, filePathManager);
            else if (productName == "黑体辐射源_")
                return new BlackBodyDoubleRadiation.WordParser(originalExcelFileFullName, excelDataFileFullName, filePathManager);
            else if (productName == "面辐射源")
                return new BlackBodySurfaceRadiation.WordParser(originalExcelFileFullName, excelDataFileFullName, filePathManager);

            throw new Exception("根据产品名称，找不到 WordParseBase (" + productName + ").");

        }

    }
}
