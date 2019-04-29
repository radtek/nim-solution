using DocumentFormat.OpenXml.Packaging;
using NIM.CertificationGenerator.Core;
using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Wordprocessing;

namespace NIM.CertificationGenerator.RadiationThermomater
{
    public class WordParser: WordParseBase
    {
        public WordParser( string originalExcelFileFullName, string copyedExcelFileFullName, FilePathManager filePathManager):
            base(originalExcelFileFullName, copyedExcelFileFullName, filePathManager)
        {

        }
        private TestingProcessResultDto Result { get;  set; }
       
        private string ResultWordFileFullName { get;  set; }
 


        private void SetWordFileFullName()
        {
            var productCode = this.Result.ProductName;
            //根据产品名称，取得配置文件中要求的产品名称
            productCode = ProductAliasName.GetProductName(productCode);

            if (this.Result.光阑直径 == "/")
                productCode += "_校准";
            else
                productCode += "_工作用";

            var wordTemplateConfPath = WordTemplateConf.GetCertificateTemplateConfFileFullName(productCode, this.ExcelDataFileFullName);


            var wordResultPath = this.FilePathManager.GetWordResultPath(this.mOriginalExcelFileName);

            var excelFileShortName = System.IO.Path.GetFileNameWithoutExtension(this.mOriginalExcelFileName);


            var wordResultFileName = System.IO.Path.Combine(wordResultPath, excelFileShortName + ".docx");
            if (System.IO.File.Exists(wordResultFileName))
                System.IO.File.Delete(wordResultFileName);
            
            System.IO.File.Copy(wordTemplateConfPath, wordResultFileName);
            this.ResultWordFileFullName = wordResultFileName;

        }
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
            this.Result.TestingStatstics = this.Result.TestingStatstics.OrderBy(t => decimal.Parse(t.StandardTemperature)).ThenBy(t => t.Note).ToList();

        }


        public void ParseCore()
        {
            var list = new List<(string replacedString, string replacingString)>();
            using (WordprocessingDocument document = WordprocessingDocument.Open(this.ResultWordFileFullName, true))
            {
                this.SimpleLabelReplacement(this.Result, document);

                this.SetStatistcisTable(this.Result, document);

                this.SetTestingTools(this.Result, document);

                //再处理一下finnaly notes.

                var text = document.MainDocumentPart.Document.Body.Descendants<Text>().Where(t => t.Text == "{{finnallynotes}}").FirstOrDefault();
                if (text != null)
                {
                    var run = (Run)text.Parent;
                    var paragraph = (Paragraph)run.Parent;

                    var paragraphParent = paragraph.Parent;

                    var addedparagraph = paragraph;

                    for (var i = 0; i < this.Result.FinallyNotes.Count; i++)
                    {
                        this.Result.FinallyNotes[i] = (i + 1).ToString() + "." + this.Result.FinallyNotes[i];
                    }
                    for (var i = 0; i < this.Result.FinallyNotes.Count; i++)
                    {
                        if (i > 0)
                        {
                            paragraph = (Paragraph)paragraph.Clone();
                        }
                        paragraph.Elements<Run>().ToList()[0].Elements<Text>().ToList()[0].Text = this.Result.FinallyNotes[i];
                        if (i > 0)
                        {
                            paragraphParent.InsertAfter(paragraph, addedparagraph);
                        }
                        addedparagraph = paragraph;
                    }
                }


                this.Result.FinallyNotes.ForEach(fullString =>
                {
                    document.SetNumberStyle(fullString);

                });
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


            var firstDataRow = table.Elements<TableRow>().ToList()[1];
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
                this.SetCellValue(cells[0], obj.StandardTemperature.ToString());

                this.SetCellValue(cells[1], obj.ResultTemperature);
                this.SetCellValue(cells[2], obj.Difference);
                this.SetCellValue(cells[3], obj.ExtendDifference);
                if (this.Result.光阑直径 == "/")
                    this.SetCellValue(cells[4], obj.Note);
                if (i > 0)
                    table.InsertAfter(newRow, addedRow);
                addedRow = newRow;
            }
            if (this.Result.光阑直径 == "/")
            // if (result.CertificationType == CertificationType.Calibrating) //如果是标准用，那么需要对最后一列进行设置
            {

                //如果是校准的话，那最的一列保留，并进行垂直合并
                var rows = table.Elements<TableRow>().ToList().ToList();
                //var lastName = "";
                var lastNote = "";

                for (var i = 0; i < this.Result.TestingStatstics.Count; i++)
                {
                    var row = rows[i + 1];

                    MergedCellValues val = MergedCellValues.Restart;
                    var obj = this.Result.TestingStatstics[i];
                    //  if (obj.Name == lastName && obj.Note == lastNote)
                    if (obj.Note == lastNote)
                    {
                        val = MergedCellValues.Continue;
                    }
                    else
                    {
                        val = MergedCellValues.Restart;
                    }
                    var tabelCell = row.Elements<TableCell>().ToList()[4];
                    var cellProperties = tabelCell.Elements<TableCellProperties>().FirstOrDefault();
                    if (cellProperties == null)
                    {
                        cellProperties = new TableCellProperties();
                        tabelCell.Append(cellProperties);
                    }
                    var mergin = cellProperties.Elements<VerticalMerge>().FirstOrDefault();
                    if (mergin == null)
                    {
                        mergin = new VerticalMerge();
                        cellProperties.Append(mergin);
                    }
                    mergin.Val = val;
                    if (val == MergedCellValues.Restart)
                    {
                        var paragraph = tabelCell.Elements<Paragraph>().FirstOrDefault();
                        if (paragraph == null)
                        {
                            paragraph = new Paragraph();
                            tabelCell.AppendChild(paragraph);
                        }
                        var run = paragraph.Elements<Run>().FirstOrDefault();
                        if (run == null)
                        {
                            run = new Run();
                            paragraph.AppendChild(run);
                        }
                        var _text = run.Elements<Text>().FirstOrDefault();
                        if (_text == null)
                        {
                            _text = new Text();
                            run.AppendChild(_text);
                        }
                        _text.Text = "黑体辐射源，";

                        paragraph = (Paragraph)paragraph.Clone();

                        paragraph.Elements<Run>().First().Elements<Text>().First().Text = "空腔开口直径";
                        tabelCell.AppendChild(paragraph);

                        paragraph = (Paragraph)paragraph.Clone();

                        paragraph.Elements<Run>().First().Elements<Text>().First().Text = obj.Note;
                        tabelCell.AppendChild(paragraph);
                    }

                    var verticalAlignment = cellProperties.Elements<TableCellVerticalAlignment>().FirstOrDefault();
                    if (verticalAlignment == null)
                    {
                        verticalAlignment = new TableCellVerticalAlignment();
                        cellProperties.Append(verticalAlignment);
                    }
                    verticalAlignment.Val = TableVerticalAlignmentValues.Center;
                    //lastName = obj.Name;
                    lastNote = obj.Note;
                }
            } //由于我们分开了标准用和校准的WORD模板，因此不需要再对statistics进行操作了。


        }

        private void SimpleLabelReplacement(TestingProcessResultDto result, WordprocessingDocument document)
        {
            

            var list = WordKeyBulider<TestingProcessResultDto>.GetRepalcedValues(result);
          
            document.SearchAndReplaces(list);

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
