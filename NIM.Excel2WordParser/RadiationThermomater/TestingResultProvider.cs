using DocumentFormat.OpenXml.Packaging;
using NIM.CertificationGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NIM.Utilty;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NIM.CertificationGenerator.RadiationThermomater
{
    public class TestingResultProvider
    {
        private const int StatsticsStartIndex = 16;//检测数据从第16行开始
        private const string StatsticsSheetName = "记录";
        public TestingProcessResultDto GetResult(SpreadsheetDocument document)
        {

            var result = new TestingProcessResultDto();
            var propertyDescriptions = PropertyInfoBulider<TestingProcessResultDto>.GetPropertyAttributes();
            propertyDescriptions.ForEach(propertyDescription =>
            {
                this.SetPropertyValue(result, document, propertyDescription);

            });

            this.SetStandardToolValues(result, document);
            if (result.ProductName == "精密辐射温度计")
            {
                this.SetPrecisionProduct(result, document);
            }
            else
            {
                this.SetStatsticsValues(result, document);
            }
            this.SetStatisticsNotes(result, document);
            this.SetFinallyNotes(result, document);


            return result;

        }


        private Dictionary<string, List<(string AValue, int Index)>> mSheetBasicInfos = new Dictionary<string, List<(string AValue, int Index)>>();

        private List<(string AValue, int Index)> GetSheeValues(SpreadsheetDocument document, string sheetName)
        {
            if (!mSheetBasicInfos.Keys.Contains(sheetName))
            {
                lock (mSheetBasicInfos)
                {
                    if (!mSheetBasicInfos.Keys.Contains(sheetName))
                    {
                        var values = new List<(string AValue, int Index)>();
                        var isEnd = false;
                        var startIndex = 1;
                        while (!isEnd)
                        {
                            var address = "A" + startIndex.ToString();
                            var value = document.GetCellValue(sheetName, address);
                            if (!string.IsNullOrEmpty(value))
                                values.Add((value, startIndex));
                            else
                                isEnd = true;
                            startIndex++;
                        }
                        mSheetBasicInfos.Add(sheetName, values);
                    }

                }
            }

            return mSheetBasicInfos[sheetName];

        }


        private void SetPropertyValue(TestingProcessResultDto result, SpreadsheetDocument document, PropertyDescription propertyDescription)
        {
            try
            {
                var excelAddress = propertyDescription.Attribute.ExcelAddressName;
                var arr = excelAddress.Split(':');

                var sheetName = arr[0];
                var address = arr[1];



                var cellValue = document.GetCellValue(sheetName, address);
                object value = null;


                //if (propertyDescription.Property.Name == nameof(result.CertificationType))
                //{
                //    if (cellValue != "校准" && cellValue != "工作用")
                //        throw new Exception(propertyDescription.Attribute.Description + "必须为:校准,或者工作用，EXCEL提供的值为：" + cellValue);
                //    if (cellValue == "校准")
                //        value = CertificationType.Calibrating;
                //    else
                //        value = CertificationType.Woring;
                //}
                //else 

                if (propertyDescription.Property.Name == nameof(result.TTemperature) || propertyDescription.Property.Name == nameof(result.THumidity))
                {
                    double db = Double.Parse(cellValue);

                    value = db.Round(1).ToString().EnsureRound1();
                }
                else
                {
                    var propertyType = propertyDescription.Property.PropertyType;


                    if (propertyType == typeof(DateTime))
                    {
                        value = DateTime.FromOADate(double.Parse(cellValue));
                    }
                    else if (propertyType == typeof(int))
                    {
                        value = int.Parse(cellValue);
                    }
                    else if (propertyType == typeof(decimal))
                        value = decimal.Parse(cellValue);
                    else if (propertyType == typeof(double))
                        value = double.Parse(cellValue);
                    else if (propertyType == typeof(string))
                        value = cellValue;
                    else
                        throw new Exception("未知的属性类型." + propertyType.ToString());
                }
                propertyDescription.Property.SetValue(result, value);


            }

            catch (Exception ex)
            {
                while (ex.InnerException != null)
                    ex = ex.InnerException;
                var description = propertyDescription.Attribute.ExcelAddressName;

                throw new Exception($"未能取得 {propertyDescription.Attribute.Description} 值，对应的EXCEL地址是:{propertyDescription.Attribute.ExcelAddressName} ({ex.Message}),");
            }
        }


        private void SetStandardToolValues(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            result.TestingTools = new List<TestingTool>();

            var stringFalg = "计量标准信息"; //查找关键字

            var toRowIndex = 1000; //扫描最大的行数

            var startRowIndex = StatsticsStartIndex;
            var foundRowIndex = 0;
            while (startRowIndex <= toRowIndex)
            {
                var address = "A" + startRowIndex.ToString();
                var str = document.GetCellValue(StatsticsSheetName, address);
                if (str == stringFalg)
                {
                    foundRowIndex = startRowIndex;
                    break;
                }
                startRowIndex++;
            }
            if (foundRowIndex == 0)
                throw new Exception($@"未找到计量标准信息，它应该出现在EXCEL的[{StatsticsSheetName}]sheet里面，并且第A列");
            foundRowIndex = foundRowIndex + 1;//移动到下一行
            toRowIndex = 3;//最三个仪器，现在默认是3个
            for (var i = 0; i < toRowIndex; i++)
            {
                var obj = new TestingTool
                {
                    Name = "",
                    TempRange = "",
                    NotSure = "",
                    Code = "",
                    ExpiredDesc = ""

                };
                var thisRowIndex = foundRowIndex + i;
                var toolName = document.GetCellValue(StatsticsSheetName, "B" + thisRowIndex.ToString());
                var tempRange = document.GetCellValue(StatsticsSheetName, "C" + thisRowIndex.ToString());
                var notSure = document.GetCellValue(StatsticsSheetName, "D" + thisRowIndex.ToString());
                var code = document.GetCellValue(StatsticsSheetName, "E" + thisRowIndex.ToString());
                var cellValue = document.GetCellValue(StatsticsSheetName, "F" + thisRowIndex.ToString());

                if (string.IsNullOrEmpty(tempRange) || string.IsNullOrEmpty(notSure) || string.IsNullOrEmpty(notSure) || string.IsNullOrEmpty(cellValue))
                {
                    toolName = "";
                }
                if (!string.IsNullOrEmpty(toolName))
                {
                    obj.Name = toolName;
                    obj.TempRange = tempRange;
                    obj.NotSure = notSure;
                    obj.Code = code;
                    double oaDate = 0;
                    if (toolName != "")
                    {
                        if (!double.TryParse(cellValue, out oaDate))
                            throw new Exception("计量标准信息->证书有效期 数据不合法");

                    }
                    obj.ExpiredDesc = DateTime.FromOADate(oaDate).ToString("yyyy-MM-dd");
                }
                result.TestingTools.Add(obj);


            }


            //foundRowIndex += 1; //move to next row
            //result.TToolName = document.GetCellValue(StatsticsSheetName, "B" + foundRowIndex.ToString());
            //result.TToolTempRange = document.GetCellValue(StatsticsSheetName, "C" + foundRowIndex.ToString());
            //result.TToolNotSure = document.GetCellValue(StatsticsSheetName, "D" + foundRowIndex.ToString());
            //result.TToolCode = document.GetCellValue(StatsticsSheetName, "E" + foundRowIndex.ToString());
            //var cellValue = document.GetCellValue(StatsticsSheetName, "F" + foundRowIndex.ToString());
            //double oaDate;
            //if (!double.TryParse(cellValue, out oaDate))
            //    throw new Exception("计量标准信息->证书有效期 数据不合法");
            //result.TToolExpired = DateTime.FromOADate(oaDate);

        }

        /// <summary>
        /// 设置具体测试记录
        /// </summary>
        /// <param name="result"></param>
        /// <param name="document"></param>
        private void SetStatsticsValues(TestingProcessResultDto result, SpreadsheetDocument document)
        {

            result.TestingStatstics = new List<TestinStatisticsDto>();
            //查找依据：从startRowIndex开始，直到A到的值为：结论
            //如何断定当前行是否为有效的数据行： A列必须为整数，并且E列和F列都必须同时有值
            //取B列的值当作名称，N，O,P

            var endFalg = "结论"; //查找关键字

            var toRowIndex = 1000; //扫描最大的行数

            var startRowIndex = StatsticsStartIndex;
            var lastStandardName = "";//标准器

            var roundParameter = 1;
            if (result.IsDisplayResolutionInt)
                roundParameter = 0;
            while (startRowIndex <= toRowIndex)
            {
                var address = "A" + startRowIndex.ToString();
                var str = document.GetCellValue(StatsticsSheetName, address);
                if (str == endFalg)
                {
                    break;
                }
                int standardTemplate;
                if (int.TryParse(str, out standardTemplate)) //不错，这是一个温度
                {
                    //由于是第一行有可能是非常的，所以standard name必须始终读取
                    var value = document.GetCellValue(StatsticsSheetName, "B" + startRowIndex.ToString());
                    if (!string.IsNullOrEmpty(value))
                        lastStandardName = value;

                    //并且E列和F列都必须同时有值
                    var eValue = document.GetCellValue(StatsticsSheetName, "E" + startRowIndex.ToString());
                    var fValue = document.GetCellValue(StatsticsSheetName, "F" + startRowIndex.ToString());
                    //David 2017-7-25 如果O列的字体的色彩为白色，就表示不需要在证书中显示。
                    if (!string.IsNullOrEmpty(eValue) && !string.IsNullOrEmpty(fValue))
                    {
                        address = "O" + startRowIndex.ToString();
                        if (!document.GetHasWhiteTheme(StatsticsSheetName, address))
                        {
                            var obj = new TestinStatisticsDto();

                            obj.Name = lastStandardName;
                            try
                            {
                                double standard;
                                if (!double.TryParse(document.GetCellValue(StatsticsSheetName, "N" + startRowIndex.ToString()), out standard))
                                    throw new Exception("无法取得" + "N" + startRowIndex.ToString() + "值，它必须是一个实数");

                                obj.StandardTemperature = standard.ToString().EnsureRound1();
                                var oValue = document.GetCellValue(StatsticsSheetName, "O" + startRowIndex.ToString());
                                if(!double.TryParse(oValue,out standard))
                                    throw new Exception("无法取得" + "O" + startRowIndex.ToString() + "值，它必须是一个实数");

                                  oValue = document.GetCellValue(StatsticsSheetName, "P" + startRowIndex.ToString());
                                if (!double.TryParse(oValue, out standard))
                                    throw new Exception("无法取得" + "P" + startRowIndex.ToString() + "值，它必须是一个实数");

                                obj.ResultTemperature = double.Parse(document.GetCellValue(StatsticsSheetName, "O" + startRowIndex.ToString())).Round(roundParameter).ToString();






                                obj.Difference = double.Parse(document.GetCellValue(StatsticsSheetName, "P" + startRowIndex.ToString())).Round(roundParameter).ToString();
                                if (!result.IsDisplayResolutionInt)
                                {
                                    obj.ResultTemperature = obj.ResultTemperature.EnsureRound1();
                                    obj.Difference = double.Parse(document.GetCellValue(StatsticsSheetName, "P" + startRowIndex.ToString())).Round(roundParameter).ToString().EnsureRound1();

                                }

                                obj.ExtendDifference = document.GetCellValue(StatsticsSheetName, "Q" + startRowIndex.ToString());
                                if (!string.IsNullOrEmpty(obj.ExtendDifference))
                                {
                                    try
                                    {
                                        var db = double.Parse(obj.ExtendDifference);
                                        db = db.Round(1);

                                        obj.ExtendDifference = db.ToString().EnsureRound1();

                                    }
                                    catch { }
                                }

                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"在获取{obj.Name}的测试数据时,出错 (" + ex.Message + ")");
                            }
                            result.TestingStatstics.Add(obj);
                        }
                    }
                }
                startRowIndex++;
            }

        }

        private void SetFinallyNotes(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            result.FinallyNotes = new List<string>();
            var sheetName = "数据说明";
            var toRowIndex = 100; //扫描最大的行数
            var startRowIndex = 1;
            while (startRowIndex <= toRowIndex)
            {
                var value = document.GetCellValue(sheetName, "A" + startRowIndex.ToString());
                if (string.IsNullOrEmpty(value))
                    break;
                ////sometime ,the _.name is 说明3 ,sometime _name is 说明3:

                if (value == result.FinallyNotesFlag || value == result.FinallyNotesFlag + ":" || value == result.FinallyNotesFlag + "：")
                {
                    value = document.GetCellValue(sheetName, "B" + startRowIndex.ToString());
                    if (!String.IsNullOrEmpty(value))
                        result.FinallyNotes.Add(value);
                }
                startRowIndex++;
            }

        }


        private class _CheckToolNote
        {
            public string Name
            {
                get; set;
            }
            public string Note
            {
                get; set;
            }
        }
        private List<string> _colums = new List<string>() { "A", "B", "C", "D", "E", "F", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        /// <summary>
        /// 查找 辐射源型号 ，找到数据的行号，名字的列名和直径的列名
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private (int RowIndex, int nameColumnIndex, int noteColumnIndex) _GetTestingToolNameRowIndex(SpreadsheetDocument document)
        {
            var list = new List<_CheckToolNote>();

            var startFlag = "辐射源型号";//查找关键字
            var toRowIndex = 1000; //扫描最大的行数

            var valueFlag = "辐射源直径";


            var startRowIndex = StatsticsStartIndex;

            while (startRowIndex <= toRowIndex)
            {
                for (var i = 0; i < _colums.Count() - 2; i++) //-2的意思，就是要向后面的第二列，取得备注
                {
                    var address = _colums[i] + startRowIndex.ToString();
                    var str = document.GetCellValue(StatsticsSheetName, address);
                    if (str == startFlag)
                    {
                        for (var j = i; j < _colums.Count; j++)
                        {
                            address = _colums[j] + startRowIndex.ToString();
                            str = document.GetCellValue(StatsticsSheetName, address);
                            if (str == valueFlag)
                            {
                                return (startRowIndex + 1, i, j);
                            }

                        }


                    }
                }
                startRowIndex++;
            }
            throw new Exception($"在{StatsticsSheetName}的表格里面，未找到 辐射源型号 关键字");

        }

        /// <summary>
        /// 设置检定仪器的备注，在本例中，就是辐射源直径
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private List<_CheckToolNote> _FoundTestingToolStatisNotes(SpreadsheetDocument document)
        {
            var list = new List<_CheckToolNote>();

            var flag = _GetTestingToolNameRowIndex(document);
            var nameColumn = _colums[flag.nameColumnIndex];
            var noteColumn = _colums[flag.noteColumnIndex];
            var startInex = flag.RowIndex;

            while (true)
            {
                var name = document.GetCellValue(StatsticsSheetName, nameColumn + startInex.ToString());
                if (string.IsNullOrEmpty(name))
                    break;
                var value = document.GetCellValue(StatsticsSheetName, noteColumn + startInex);
                list.Add(new _CheckToolNote
                {
                    Name = name,
                    Note = value
                });

                startInex++;
            }

            return list;
        }
        private void SetStatisticsNotes(TestingProcessResultDto result, SpreadsheetDocument document)
        {

            var lists = this._FoundTestingToolStatisNotes(document);
            for (var i = 0; i < result.TestingStatstics.Count; i++)
            {
                var obj = result.TestingStatstics[i];
                var toolNote = lists.FirstOrDefault(t => obj.Name.Contains(t.Name));
                if (toolNote != null)
                    obj.Note = toolNote.Note;

            }
        }




        /// <summary>
        /// 设置精密辐身温度计
        /// </summary>
        private void SetPrecisionProduct(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            var sheetNames = new List<string>();
            sheetNames.Add("记录");
            //记录-x;
            WorkbookPart wbPart = document.WorkbookPart;

            var sheets = wbPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList();
            sheets.ForEach(t =>
            {
                if (t.Name.InnerText.StartsWith("记录-"))
                    sheetNames.Add(t.Name.InnerText);
            });

            var addressName = "I6";//温度所在的EXCEL addressfo
            var values = new List<decimal>();
            sheetNames.ForEach(t =>
            {
                decimal v;
                var cellValue = document.GetCellValue(t, addressName);
                if (!decimal.TryParse(cellValue, out v))
                {
                    throw new Exception($@"Sheet {t} 中的温度输入不正确.");

                }
                double db = Double.Parse(v.ToString()).Round(1);
                values.Add((decimal)db);
            });
            var minValue = values.Min();
            var maxValue = values.Max();
            if (minValue == maxValue)
                result.TTemperature = minValue.ToString().EnsureRound1();
            else
            {
                result.TTemperature = "(" + minValue.ToString().EnsureRound1() + "-" + maxValue.ToString().EnsureRound1() + ")";
            }


            addressName = "I7";//湿度所在的EXCEL addressfo
            values.Clear();
            sheetNames.ForEach(t =>
            {
                decimal v;
                var cellValue = document.GetCellValue(t, addressName);
                if (!decimal.TryParse(cellValue, out v))
                {
                    throw new Exception($@"Sheet {t} 中的湿度输入不正确.");

                }

                double db = Double.Parse(v.ToString()).Round(1);
                values.Add((decimal)db);
            });
            minValue = values.Min();
            maxValue = values.Max();
            if (minValue == maxValue)
                result.THumidity = minValue.ToString().EnsureRound1();
            else
            {
                result.THumidity = "(" + minValue.ToString().EnsureRound1() + "-" + maxValue.ToString().EnsureRound1() + ")";
            }

            var checkToolNotes = this._FoundTestingToolStatisNotes(document);



            var sheetName = "汇总";
            result.TestingStatstics = new List<TestinStatisticsDto>();
            var startRowIndex = 12; //从第11行开始读数

            var roundParameter = 1;
            if (result.IsDisplayResolutionInt)
                roundParameter = 0;

            while (true)
            {
                addressName = "A" + startRowIndex.ToString();
                var str = document.GetCellValue(sheetName, addressName);
                decimal standardTemperature;
                if (!decimal.TryParse(str, out standardTemperature))
                    break;
                var obj = new TestinStatisticsDto();
                obj.Name = "";
                obj.StandardTemperature = standardTemperature.ToString().EnsureRound1();

                addressName = "B" + startRowIndex.ToString();
                var val = document.GetCellValue(sheetName, addressName);
                double v;
                if (double.TryParse(val, out v))
                {
                    obj.ResultTemperature = double.Parse(document.GetCellValue(sheetName, addressName)).Round(roundParameter).ToString().EnsureRound1();

                    addressName = "C" + startRowIndex.ToString();
                    obj.Difference = double.Parse(document.GetCellValue(sheetName, addressName)).Round(roundParameter).ToString().EnsureRound1();

                    addressName = "D" + startRowIndex.ToString();
                    obj.ExtendDifference = double.Parse(document.GetCellValue(sheetName, addressName)).Round(1).ToString().EnsureRound1();

                    obj.Note = _GetProductNote(document, checkToolNotes, standardTemperature);
                    result.TestingStatstics.Add(obj);
                }
                startRowIndex = startRowIndex + 1;
            }
        }

        private string _GetProductNote(SpreadsheetDocument document, List<_CheckToolNote> toolNotes, decimal standardTemperature)
        {



            var sheetName = "记录";

            var startRowIndex = 16; //从第16行开始读数

            while (true)
            {
                var addressName = "A" + startRowIndex.ToString();
                var str = document.GetCellValue(sheetName, addressName);
                decimal cellStandardTemperature;
                if (!decimal.TryParse(str, out cellStandardTemperature))
                    break;
                if (cellStandardTemperature == standardTemperature)
                {
                    str = "";
                    for (var i = startRowIndex; i >= 0; i--)
                    {
                        addressName = "B" + i.ToString();
                        str = document.GetCellValue(sheetName, addressName);
                        if (!string.IsNullOrEmpty(str))
                            break;
                    }
                    if (!string.IsNullOrEmpty(str))
                    {
                        var notes = toolNotes.Where(t => str.Contains(t.Name)).FirstOrDefault();
                        if (notes != null)
                            return notes.Note;
                    }




                }
                startRowIndex = startRowIndex + 1;
            }
            return "";
        }
    }
}

