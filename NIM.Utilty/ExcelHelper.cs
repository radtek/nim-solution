using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.Utilty
{
    public static class ExcelHelper
    {


        public static string GetCellValue(this WorksheetPart wsPart, WorkbookPart wbPart, string addressName)
        {


            string value = null;
            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();


            // If the cell does not exist, return an empty string.
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                    if (theCell.CellFormula != null)
                    {
                        value = theCell.CellValue.InnerText;
                    }
                }
                else
                {
                    if (theCell.CellFormula != null)
                    {
                        value = theCell.CellValue.InnerText;
                    }
                }
            }
            return value;
        }

        public static string GetCellValue(string fileName,
             string sheetName,
             string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, true))
            {
                value = document.GetCellValue(sheetName, addressName);
            }
            return value;
        }


        public static string GetCellValue(this SpreadsheetDocument document,
           string sheetName,
           string addressName)
        {
            string value = "";

            // Open the spreadsheet document for read-only access.

            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            // Throw an exception if there is no sheet.
            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart =
                (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();

            // If the cell does not exist, return an empty string.
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                    if (theCell.CellFormula != null)
                    {
                        value = theCell.CellValue.InnerText;
                    }
                }
                else
                {
                    if (theCell.CellFormula != null)
                    {
                        value = theCell.CellValue.InnerText;
                    }
                }
            }

            return value;
        }

        public static Cell GetCellObject(this SpreadsheetDocument document,
           string sheetName,
           string addressName)
        {

            // Open the spreadsheet document for read-only access.

            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            // Throw an exception if there is no sheet.
            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart =
                (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();

            if (theCell == null)
                throw new Exception("未找到EXCEL对应的列" + addressName);
            return theCell;
        }

        public static bool GetHasWhiteTheme(this SpreadsheetDocument document,
           string sheetName,
           string addressName)
        {
            var theCell = document.GetCellObject(sheetName, addressName);

            if (theCell.StyleIndex == null)
                return false;
            int cellStyleIndex = (int)theCell.StyleIndex.Value;

            WorkbookPart wbPart = document.WorkbookPart;
            WorkbookStylesPart styles = wbPart.WorkbookStylesPart.Stylesheet.WorkbookStylesPart;
            CellFormat cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[cellStyleIndex];

            var font = (Font)styles.Stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value];

            if (font == null)
                return false;
            var color = font.Elements<Color>().FirstOrDefault();
            if (color == null)
                return false;
            if (color.Theme == null)
                return false;
            return color.Theme.Value == 0; //0 white

        }
       public static bool GetHasGoodFgColor(this SpreadsheetDocument document,
           string sheetName,
           string addressName)
        {
            var theCell = document.GetCellObject(sheetName, addressName);

            if (theCell.StyleIndex == null)
                return false;
            int cellStyleIndex = (int)theCell.StyleIndex.Value;

            WorkbookPart wbPart = document.WorkbookPart;
            WorkbookStylesPart styles = wbPart.WorkbookStylesPart.Stylesheet.WorkbookStylesPart;
            CellFormat cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[cellStyleIndex];

            if (cellFormat.FillId == null)
                return false;

            var fill = (Fill)styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value];
            if (fill == null)
                return false;

            var patternFill = fill.Elements<PatternFill>().FirstOrDefault();
            if (patternFill == null)
                return false;
            var color = patternFill.Elements<ForegroundColor>().FirstOrDefault();
            if (color == null)
                return false;
            if (color.Rgb == null)
                return false;
            return color.Rgb.Value == "FFC6EFCE";
             

        }
    }
}
