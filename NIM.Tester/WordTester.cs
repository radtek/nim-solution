using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace ExcelTestor
{
    [TestClass]
    public class WordTester
    {
        private string templatePath = @"d:\word.docx";


        [TestMethod]
        public void ReplaceNoteTest()
        {
            var fileName = @"C:\Users\Exten\Dropbox\nim\documents\证书生成结果文件夹\rr-RGrr2017-0-0.95-模板-yangx - Copy.docx";
            using (WordprocessingDocument document
              = WordprocessingDocument.Open(fileName, true))
            {
                var fullString = "U=（0.14～6.5）°C（k=2）";

                var originalText = document.MainDocumentPart.Document.Body.Descendants<Text>().Where(t => t.Text == fullString).FirstOrDefault();

                originalText.Text = "";

                var originalRun = (Run)originalText.Parent;

                var newText = (Text)originalText.Clone();
                newText.Text = "";
                var newRun = (Run)originalRun.Clone();

                var paragraph = (Paragraph)originalRun.Parent;
                originalRun.RemoveChild(originalText);
                paragraph.RemoveChild(originalRun);


                
                var lastWord = "";
                Text _text;
                Run _run;
                for (var i = 0; i < fullString.Length; i++)
                {
                    var s = fullString[i];
                    if (s == 'U' || s == 'k')
                    {
                        if (lastWord != "")
                        {
                            _run = (Run)newRun.Clone();
                            _text = (Text)newText.Clone();
                            _text.Text = lastWord;
                            _run.Append(_text);
                            paragraph.Append(_run);
                        }


                        _text = (Text)newText.Clone();
                        _text.Text = s.ToString();
                        _run = (Run)originalRun.Clone();
                        _run.Append(_text);
                        paragraph.Append(_run);

                        var runProperties = _run.Elements<RunProperties>().FirstOrDefault();
                        if (runProperties == null)
                        {
                            runProperties = new RunProperties();
                            _run.Append(runProperties);
                        }
                        var italic = runProperties.Elements<Italic>().FirstOrDefault();
                        if (italic == null)
                        {
                            italic = new Italic();
                            runProperties.Append(italic);
                        }
                        lastWord = "";
                    }

                    else
                        lastWord += s;
                }
                if (lastWord != "")
                {
                    _run = (Run)originalRun.Clone();
                    _text = (Text)newText.Clone();
                    _text.Text = lastWord;
                    _run.Append(_text);
                    paragraph.Append(_run);

                }

                document.Save();


            }


        }





        private (TableCell Cell, Paragraph Paragraph) findParagraph(List<Table> tables)
        {


            for (var i = 0; i < tables.Count; i++)
            {
                var table = tables[i];
                var tableRows = table.Elements<TableRow>().ToList();
                for (var j = 0; j < tableRows.Count; j++)
                {
                    var row = tableRows[j];
                    var cells = row.Elements<TableCell>().ToList();
                    for (var k = 0; k < cells.Count; k++)
                    {
                        var cell = cells[k];
                        var paragraphs = cell.Elements<Paragraph>().ToList();
                        for (var kk = 0; kk < paragraphs.Count; kk++)
                        {
                            var paragraph = paragraphs[kk];
                            var runs = paragraph.Elements<Run>().ToList();
                            for (var kkk = 0; kkk < runs.Count(); kkk++)
                            {
                                var texts = runs[kkk].Elements<Text>().ToList();
                                if (texts.Count() != 1)
                                    continue;
                                if (texts[0].Text != "{{testingresultnoteplaceholder}}")
                                    continue;
                                return (cell, paragraph);
                            }
                        }

                    }
                }
            }
            return (null, null);


        }




        [TestMethod]
        public void ReplaceWordTest()
        {
            var a1 = "12";
            var b1 = "14";
            var c1 = "26";
            var file2 = @"d:\" + Guid.NewGuid().ToString() + ".docx";
            System.IO.File.Copy(templatePath, file2);
            SearchAndReplace(file2, "{{A1}}", a1);
            SearchAndReplace(file2, "{{B1}}", b1);
            SearchAndReplace(file2, "{{A1+B1}}", c1);
        }

        [TestMethod]
        public void InsertTableTest()
        {
            CreateTable(templatePath);
        }



        [TestMethod]
        public void InsertRowsTest()
        {
            InsrtRows(templatePath);
        }

        public static void InsrtRows(string fileName)
        {
            // Use the file name and path passed in as an argument 
            // to open an existing Word 2007 document.

            using (WordprocessingDocument doc
                = WordprocessingDocument.Open(fileName, true))
            {

                // Find the first table in the document.
                Table table =
                    doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                var row = table.Elements<TableRow>().ElementAt(0);

                var rowClone = (TableRow)row.Clone();

                var cell = rowClone.Elements<TableCell>().ElementAt(0);
                SetCellValue(cell, "10000");
                // rowClone.Append(cell);

                cell = rowClone.Elements<TableCell>().ElementAt(1);
                SetCellValue(cell, "123.45");
                //rowClone.Append(cell);

                table.Append(rowClone);

                doc.Save();
            }
        }

        private static void SetCellValue(TableCell cell, string txt)
        {
            Paragraph p = cell.Elements<Paragraph>().First();
            Run r = p.Elements<Run>().First();
            Text t = r.Elements<Text>().First();
            t.Text = txt;
        }


        // Insert a table into a word processing document.
        public static void CreateTable(string fileName)
        {
            // Use the file name and path passed in as an argument 
            // to open an existing Word 2007 document.

            using (WordprocessingDocument doc
                = WordprocessingDocument.Open(fileName, true))
            {
                // Create an empty table.
                Table table = new Table();

                // Create a TableProperties object and specify its border information.
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new BottomBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new LeftBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new RightBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new InsideHorizontalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new InsideVerticalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        }
                    )
                );

                // Append the TableProperties object to the empty table.
                table.AppendChild<TableProperties>(tblProp);

                // Create a row.
                TableRow tr = new TableRow();

                // Create a cell.
                TableCell tc1 = new TableCell();

                // Specify the width property of the table cell.
                tc1.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

                // Specify the table cell content.
                tc1.Append(new Paragraph(new Run(new Text("some text"))));

                // Append the table cell to the table row.
                tr.Append(tc1);

                // Create a second table cell by copying the OuterXml value of the first table cell.
                TableCell tc2 = new TableCell(tc1.OuterXml);

                // Append the table cell to the table row.
                tr.Append(tc2);

                // Append the table row to the table.
                table.Append(tr);

                // Append the table to the document.
                doc.MainDocumentPart.Document.Body.Append(table);
            }
        }
        public static void SearchAndReplace(string document, string replacedString, string replacingString)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                // Regex regexText = new Regex(replacedString);
                docText = docText.Replace(replacedString, replacingString);
                // docText = regexText.Replace(docText, replacedString);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
                wordDoc.Save();
            }
        }
    }
}
