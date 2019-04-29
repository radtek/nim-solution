using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NIM.CertificationGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NIM.Utilty;

namespace NIM.CertificationGenerator.ThermalImager
{
    public class WordParse : WordParseBase
    {
        public WordParse(string originalExcelFileFullName, string copyedExcelFileFullName, FilePathManager filePathManager) :
            base(originalExcelFileFullName, copyedExcelFileFullName, filePathManager)
        {

        }


        private string ResultWordFileFullName { get; set; }
        private TestingProcessResultDto Result { get; set; }

        public Core.ITestingResultDto TestingResult
        {
            get
            {
                if (this.Result == null)
                {
                    this.InitResult();
                }
                return this.Result;
            }
        }
        public override string GeneraterFile()
        {
            try
            {
                if (this.Result == null)
                {
                    //根据传入的EXCEL，得到一个DTO对象
                    this.InitResult();
                }


                //对得到的EXCEL DTO进行一些预处理，以满足生成的要求

                this.PrepareResult();

                //根据配置文件，得到一个生成的WORD证书
                this.SetWordFileFullName();
                //将WORD证书中的值全部替换过来
                this.ParseCore();

                return this.ResultWordFileFullName;
            }
            finally
            {
                if (System.IO.File.Exists(this.ExcelDataFileFullName))
                    System.IO.File.Delete(this.ExcelDataFileFullName);

            }

        }

        private void SetWordFileFullName()
        {
            var productCode = this.Result.ProductName;

            //根据产品名称，取得配置文件中要求的产品名称
            productCode = ProductAliasName.GetProductName(productCode);
            if (this.Result.FileType == FileType.Range1)
                productCode += "_单一量程";
            else if (this.Result.FileType == FileType.Range2)
                productCode += "_二量程";
            else if (this.Result.FileType == FileType.Range3)
                productCode += "_三量程";
            else if (this.Result.FileType == FileType.Range4)
                productCode += "_四量程";
            else
                throw new Exception("不被支持的量程类型." + this.Result.FileType.ToString());


            var wordTemplateConfPath = WordTemplateConf.GetCertificateTemplateConfFileFullName(productCode, this.ExcelDataFileFullName);


            var wordResultPath = this.FilePathManager.GetWordResultPath(this.mOriginalExcelFileName);

            var excelFileShortName = System.IO.Path.GetFileNameWithoutExtension(this.mOriginalExcelFileName);


            var wordResultFileName = System.IO.Path.Combine(wordResultPath, excelFileShortName + ".docx");
            if (System.IO.File.Exists(wordResultFileName))
                System.IO.File.Delete(wordResultFileName);

            System.IO.File.Copy(wordTemplateConfPath, wordResultFileName);
            this.ResultWordFileFullName = wordResultFileName;



        }
        private void PrepareResult()
        {
            this.Result.TemperatureRanges = this.Result.Ranges.OrderBy(t => t.StandardTemperature).Select(t => t.DisplayStandardTemperature).Distinct().ToList();


        }

        private void InitResult()
        {
            using (SpreadsheetDocument document =
           SpreadsheetDocument.Open(this.ExcelDataFileFullName, false))
            {
                var resultProvider = new TestingResultProvider();
                this.Result = resultProvider.GetResult(document);
            }

        }

        private void ParseCore()
        {
            var list = new List<(string replacedString, string replacingString)>();
            using (WordprocessingDocument document = WordprocessingDocument.Open(this.ResultWordFileFullName, true))
            {
                this.SimpleLabelReplacement(this.Result, document);

                this.SetStatistcisTable(this.Result, document);

                this.SetFinallyNotes(this.Result, document);
                document.Save();
            }
        }


        private void SimpleLabelReplacement(TestingProcessResultDto result, WordprocessingDocument document)
        {


            var list = WordKeyBulider<TestingProcessResultDto>.GetRepalcedValues(result);

            list.AddRange(this._GetStatisticsRangeDescription(result));
            list.AddRange(this._GetPointValues(result));

            document.SearchAndReplaces(list);

        }

        private List<(string replacedString, string replacingString)> _GetPointValues(TestingProcessResultDto result)
        {
            var list = new List<(string, string)>();
            for (var i = 0; i < result.TestingPointValues.Count(); i++)
            {
                var obj = result.TestingPointValues[i];
                var key = "{{pointvalue" + obj.Location + "}}";
                var value = obj.Value;
                list.Add((key, value));


            }

            return list;
        }

        private List<(string replacedString, string replacingString)> _GetStatisticsRangeDescription(TestingProcessResultDto result)
        {
            var range1 = "";
            var range2 = "";
            var range3 = "";
            var range4 = "";

            var obj = this.Result.Ranges.Where(t => t.Type == RangetType.范围1).FirstOrDefault();
            if (obj != null)
                range1 = obj.Description;
            obj = this.Result.Ranges.Where(t => t.Type == RangetType.范围2).FirstOrDefault();
            if (obj != null)
                range2 = obj.Description;
            obj = this.Result.Ranges.Where(t => t.Type == RangetType.范围3).FirstOrDefault();
            if (obj != null)
                range3 = obj.Description;
            obj = this.Result.Ranges.Where(t => t.Type == RangetType.范围4).FirstOrDefault();
            if (obj != null)
                range4 = obj.Description;

            var list = new List<(string, string)>();

            list.Add(("{{range1}}", range1));
            list.Add(("{{range2}}", range2));
            list.Add(("{{range3}}", range3));
            list.Add(("{{range4}}", range4));
            return list;
        }

        private void SetStatistcisTable(TestingProcessResultDto result, WordprocessingDocument document)
        {
            var tables = document.MainDocumentPart.Document.Body.Elements<Table>().ToList();

            Table table = null;
            //statstics-data

            IEnumerable<TableProperties> tableProperties = document.MainDocumentPart.Document.Body.Descendants<TableProperties>().Where(tp => tp.TableCaption != null);
            foreach (TableProperties tProp in tableProperties)
            {
                if (tProp.TableCaption.Val.InnerText.Equals("statstics-data"))
                {
                    // do something for table with myCaption
                    table = (Table)tProp.Parent;
                }
            }

            if (table == null)
                throw new Exception("未找到填充数据的表格。(caption:statstics-data)");

            TableRow firstDataRow;
            if (result.FileType == FileType.Range1)
                firstDataRow = table.Elements<TableRow>().ToList()[2];
            else
                firstDataRow = table.Elements<TableRow>().ToList()[3];
            if (result.TemperatureRanges.Count == 0)
                return;

            var addedRow = firstDataRow;
            for (var i = 0; i < result.TemperatureRanges.Count; i++)
            {
                var standardTempture = result.TemperatureRanges[i];
                TableRow newRow;
                if (i == 0)
                    newRow = firstDataRow;
                else
                    newRow = (TableRow)firstDataRow.Clone();
                var cells = newRow.Elements<TableCell>().ToList();
                //输入第一个量程的
                this.SetCellValue(cells[0], standardTempture);
                var obj = result.Ranges.Where(t => t.DisplayStandardTemperature == standardTempture &&
                    t.Type == RangetType.范围1).FirstOrDefault();
                var difference = "—";
                var uValue = "—";
                if (obj != null)
                {
                    difference = obj.Difference;
                    uValue = obj.UValue;
                }
                this.SetCellValue(cells[1], difference);
                this.SetCellValue(cells[2], uValue);
                if (cells.Count > 3) //RangetType.范围1
                {
                    obj = result.Ranges.Where(t => t.DisplayStandardTemperature == standardTempture &&
                    t.Type == RangetType.范围2).FirstOrDefault();
                    difference = "—";
                    uValue = "—";
                    if (obj != null)
                    {
                        difference = obj.Difference;
                        uValue = obj.UValue;
                    }
                    this.SetCellValue(cells[3], difference);
                    this.SetCellValue(cells[4], uValue);
                }
                if (cells.Count > 5)
                {
                    obj = result.Ranges.Where(t => t.DisplayStandardTemperature == standardTempture &&
                   t.Type == RangetType.范围3).FirstOrDefault();
                    difference = "—";
                    uValue = "—";
                    if (obj != null)
                    {
                        difference = obj.Difference;
                        uValue = obj.UValue;
                    }
                    this.SetCellValue(cells[5], difference);
                    this.SetCellValue(cells[6], uValue);
                }
                if (cells.Count > 7)
                {
                    obj = result.Ranges.Where(t => t.DisplayStandardTemperature == standardTempture &&
                   t.Type == RangetType.范围4).FirstOrDefault();
                    difference = "—";
                    uValue = "—";
                    if (obj != null)
                    {
                        difference = obj.Difference;
                        uValue = obj.UValue;
                    }
                    this.SetCellValue(cells[7], difference);
                    this.SetCellValue(cells[8], uValue);
                }

                if (i > 0)
                    table.InsertAfter(newRow, addedRow);
                addedRow = newRow;

            }
        }
        private void SetFinallyNotes(TestingProcessResultDto result, WordprocessingDocument document)
        {
            var name = "{{FinallyNotes}}";


            var text = document.MainDocumentPart.Document.Descendants<Text>().Where(t => t.Text == name).FirstOrDefault();

            if (text != null)
            {
                text.Text = "";
                result.FinallyNotes = "注" + result.FinallyNotes;
                WordHelper.ComplexReplace((Paragraph)text.Parent.Parent, result.FinallyNotes, new char[] { 'U', 'k' });
            }

        }

        private void SetCellValue(TableCell cell, string value)
        {
            var paragraph = cell.Elements<Paragraph>().FirstOrDefault();
            if (paragraph == null)
            {
                paragraph = new Paragraph();
                cell.AppendChild(paragraph);
            }
            var run = paragraph.Elements<Run>().FirstOrDefault();
            if (run == null)
            {
                run = new Run();
                paragraph.AppendChild(run);
            }
            var text = run.Elements<Text>().FirstOrDefault();
            if (text == null)
            {
                text = new Text();
                run.AppendChild(text);
            }
            text.Text = value;


        }
    }
}
