using DocumentFormat.OpenXml.Packaging;
using NIM.CertificationGenerator.Core;
using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Wordprocessing;

namespace NIM.CertificationGenerator.BlackBodySurfaceRadiation
{
    public class WordParser : WordParseBase
    {
        public WordParser(string originalExcelFileFullName, string copyedExcelFileFullName, FilePathManager filePathManager) :
            base(originalExcelFileFullName, copyedExcelFileFullName, filePathManager)
        {

        }
        private TestingProcessResultDto Result { get; set; }

        private string ResultWordFileFullName { get; set; }


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

        private void SetWordFileFullName()
        {
            var productCode = this.Result.ProductName;
            //根据产品名称，取得配置文件中要求的产品名称
            productCode = ProductAliasName.GetProductName(productCode);



            var wordTemplateConfPath = WordTemplateConf.GetCertificateTemplateConfFileFullName(productCode, this.ExcelDataFileFullName);


            var wordResultPath = this.FilePathManager.GetWordResultPath(this.mOriginalExcelFileName);

            var excelFileShortName = System.IO.Path.GetFileNameWithoutExtension(this.mOriginalExcelFileName);


            var wordResultFileName = System.IO.Path.Combine(wordResultPath, excelFileShortName + ".docx");
            if (System.IO.File.Exists(wordResultFileName))
                System.IO.File.Delete(wordResultFileName);

            System.IO.File.Copy(wordTemplateConfPath, wordResultFileName);
            this.ResultWordFileFullName = wordResultFileName;

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

        private void InitResult()
        {
            using (SpreadsheetDocument document =
              SpreadsheetDocument.Open(this.ExcelDataFileFullName, false))
            {
                var resultProvider = new TestingResultProvider();
                this.Result = resultProvider.GetResult(document);
            }
        }


        private void PrepareResult()
        {

            this.Result.TestingStatstics = this.Result.TestingStatstics.OrderBy(t => decimal.Parse(t.温控设置)).ToList();

        }


        public void ParseCore()
        {
            var list = new List<(string replacedString, string replacingString)>();
            using (WordprocessingDocument document = WordprocessingDocument.Open(this.ResultWordFileFullName, true))
            {
                var olderProductName = this.Result.ProductName;
                if (this.Result.ProductName == "黑体辐射源*")
                    this.Result.ProductName = "黑体辐射源";
                this.SimpleLabelReplacement(this.Result, document);
                this.Result.ProductName = olderProductName;
                this.SetStatistcisTable(this.Result, document);

                this.SetTestingTools(this.Result, document);

                //再处理一下finnaly notes.B
                this.SetFinallyNotes(this.Result, document);
                //   this.SetAccordingByCondtions(this.Result, document);

                document.Save();
            }
        }

        private void SetTestingTools(TestingProcessResultDto result, WordprocessingDocument document)
        {
            for (var i = 0; i < result.TestingTools.Count; i++)
            {
                var flag = "";
                if (i > 0)
                    flag = (i + 1).ToString();
                var name = "{{TToolName" + flag + "}}";

                var text = document.MainDocumentPart.Document.Descendants<Text>().Where(t => t.Text == name).FirstOrDefault();
                if (text != null)
                    text.Text = result.TestingTools[i].Name;

                name = "{{TToolTempRange" + flag + "}}";

                text = document.MainDocumentPart.Document.Descendants<Text>().Where(t => t.Text == name).FirstOrDefault();
                if (text != null)
                    text.Text = result.TestingTools[i].TempRange;
                //{TToolNotSure}}
                name = "{{TToolNotSure" + flag + "}}";
                text = document.MainDocumentPart.Document.Descendants<Text>().Where(t => t.Text == name).FirstOrDefault();
                if (text != null)
                {
                    text.Text = "";
                    WordHelper.SetKeyWordItalicStyle((Paragraph)text.Parent.Parent, result.TestingTools[i].NotSure, new char[] { 'U', 'k' });
                }

                name = "{{TToolCode" + flag + "}}";

                text = document.MainDocumentPart.Document.Descendants<Text>().Where(t => t.Text == name).FirstOrDefault();
                if (text != null)
                    text.Text = result.TestingTools[i].Code;


                name = "{{TToolExpiredDesc" + flag + "}}";

                text = document.MainDocumentPart.Document.Descendants<Text>().Where(t => t.Text == name).FirstOrDefault();
                if (text != null)
                    text.Text = result.TestingTools[i].ExpiredDesc;

            }





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


            var firstDataRow = table.Elements<TableRow>().ToList()[2];
            if (result.TestingStatstics.Count == 0)
                return;

            var addedRow = firstDataRow;
            for (var i = 0; i < result.TestingStatstics.Count; i++)
            {
                var obj = result.TestingStatstics[i];
                TableRow newRow;
                if (i == 0)
                    newRow = firstDataRow;
                else
                    newRow = (TableRow)firstDataRow.Clone();
                var cells = newRow.Elements<TableCell>().ToList();
                this.SetCellValue(cells[0], obj.温控设置.ToString());

                this.SetCellValue(cells[1], obj.控温显示);
                this.SetCellValue(cells[2], obj.亮度温度);
                this.SetCellValue(cells[3], obj.U值);
                this.SetCellValue(cells[4], obj.温度计示值);
                this.SetCellValue(cells[5], obj.U值);


                if (i > 0)
                    table.InsertAfter(newRow, addedRow);
                addedRow = newRow;
            }


        }

        private void SimpleLabelReplacement(TestingProcessResultDto result, WordprocessingDocument document)
        {


            var list = WordKeyBulider<TestingProcessResultDto>.GetRepalcedValues(result);

            document.SearchAndReplaces(list);

        }

        //private void SetAccordingByCondtions(TestingProcessResultDto result, WordprocessingDocument document)
        //{
        //    var text = document.MainDocumentPart.Document.Body.Descendants<Text>().Where(t => t.Text == "{{AccordingByConditions}}").FirstOrDefault();
        //    if (text != null)
        //    {
        //        var run = (Run)text.Parent;
        //        var paragraph = (Paragraph)run.Parent;

        //        var paragraphParent = paragraph.Parent;

        //        var addedparagraph = paragraph;


        //        for (var i = 0; i < this.Result.AccordingByConditions.Count; i++)
        //        {
        //            if (i > 0)
        //            {
        //                paragraph = (Paragraph)paragraph.Clone();
        //            }
        //            paragraph.Elements<Run>().ToList()[0].Elements<Text>().ToList()[0].Text = this.Result.AccordingByConditions[i];
        //            if (i > 0)
        //            {
        //                paragraphParent.InsertAfter(paragraph, addedparagraph);
        //            }
        //            addedparagraph = paragraph;
        //        }
        //    }


        //    this.Result.AccordingByConditions.ForEach(fullString =>
        //    {
        //        document.SetNumberStyle(fullString);

        //    });

        //}
        private void SetFinallyNotes(TestingProcessResultDto result, WordprocessingDocument document)
        {
            var text = document.MainDocumentPart.Document.Body.Descendants<Text>().Where(t => t.Text == "{{finnallynotes}}").FirstOrDefault();
            if (text != null)
            {
                text.Text = "";
                var run = (Run)text.Parent;
                var paragraph = (Paragraph)run.Parent;

                var paragraphParent = paragraph.Parent;

                var clonedParagrapth = (Paragraph)paragraph.Clone();
                var addedparagraph = paragraph;

                for (var i = 0; i < this.Result.FinallyNotes.Count; i++)
                {
                    this.Result.FinallyNotes[i] = (i + 1).ToString() + "." + this.Result.FinallyNotes[i];
                   
                }
                for (var i = 0; i < this.Result.FinallyNotes.Count; i++)
                {
                    Paragraph thisParagraph;
                    if (i == 0)
                        thisParagraph = paragraph;
                    else
                        thisParagraph = (Paragraph)clonedParagrapth.Clone();
                    WordHelper.SetKeyWordItalicStyle(thisParagraph, this.Result.FinallyNotes[i], new char[] { 'U', 'k' });
                    if (i > 0)
                    {
                        paragraphParent.InsertAfter(thisParagraph, addedparagraph);
                    }

                    addedparagraph = thisParagraph;
                }
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
