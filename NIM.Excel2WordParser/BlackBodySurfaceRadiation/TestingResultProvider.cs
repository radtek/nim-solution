using DocumentFormat.OpenXml.Packaging;
using NIM.CertificationGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NIM.Utilty;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NIM.CertificationGenerator.BlackBodySurfaceRadiation
{
    public class TestingResultProvider
    {

        public TestingProcessResultDto GetResult(SpreadsheetDocument document)
        {

            var result = new TestingProcessResultDto();
            var propertyDescriptions = PropertyInfoBulider<TestingProcessResultDto>.GetPropertyAttributes();
            propertyDescriptions.ForEach(propertyDescription =>
            {
                this.SetPropertyValue(result, document, propertyDescription);

            });
         

            this.SetStandardToolValues(result, document);
            this.SetStatsticsValues(result, document);
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




                if (propertyDescription.Property.Name == nameof(result.TTemperature) || propertyDescription.Property.Name == nameof(result.THumidity))
                {
                    value = InternalHepers.ChangeTemperatureHumidityValue(cellValue);
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
            var sheetName = "汇总";

            result.TestingTools = new List<TestingTool>();

            var stringFalg = "黑体辐射源校准装置"; //查找关键字

            var toRowIndex = 1000; //扫描最大的行数

            var startRowIndex = 0;
            var foundRowIndex = 0;
            while (startRowIndex <= toRowIndex)
            {
                var address = "B" + startRowIndex.ToString();
                var str = document.GetCellValue(sheetName, address);
                if (str == stringFalg)
                {
                    foundRowIndex = startRowIndex;
                    break;
                }
                startRowIndex++;
            }
            if (foundRowIndex == 0)
                throw new Exception($@"未找到`黑体辐射源校准装置`，它应该出现在EXCEL的{sheetName}的工作表之中的B列");

            toRowIndex = 2;//最多3个仪器，现在默认是2个
            for (var i = 0; i < toRowIndex; i++)
            {
                var thisRowIndex = foundRowIndex + i;
                //约定，如果字体色是白色的风格的话，就忽略这一行
                if (document.GetHasWhiteTheme(sheetName, "B" + thisRowIndex))
                    continue;

                var toolName = document.GetCellValue(sheetName, "B" + thisRowIndex.ToString());
                var tempRange = document.GetCellValue(sheetName, "C" + thisRowIndex.ToString());
                var notSure = document.GetCellValue(sheetName, "D" + thisRowIndex.ToString());
                var code = document.GetCellValue(sheetName, "E" + thisRowIndex.ToString());
                var expiredDataString = document.GetCellValue(sheetName, "F" + thisRowIndex.ToString());

                if (string.IsNullOrEmpty(tempRange) || string.IsNullOrEmpty(notSure) || string.IsNullOrEmpty(notSure) || string.IsNullOrEmpty(expiredDataString))
                {
                    toolName = "";
                }
                var obj = new TestingTool
                {
                    Name = "",
                    TempRange = "",
                    NotSure = "",
                    Code = "",
                    ExpiredDesc = ""
                };

                obj.Name = toolName;
                obj.TempRange = tempRange;
                obj.NotSure = notSure;
                obj.Code = code;
                double oaDate = 0;

                if (!double.TryParse(expiredDataString, out oaDate))
                    throw new Exception("计量标准信息->证书有效期 数据不合法" + expiredDataString);

                obj.ExpiredDesc = DateTime.FromOADate(oaDate).ToString("yyyy-MM-dd");

                result.TestingTools.Add(obj);
            }
            var leftCount = 3 - result.TestingTools.Count;
            for (var i = 0; i < leftCount; i++)
            {
                var obj = new TestingTool
                {
                    Name = "",
                    TempRange = "",
                    NotSure = "",
                    Code = "",
                    ExpiredDesc = ""
                };
                result.TestingTools.Add(obj);
            }
        }

        /// <summary>
        /// 设置具体测试记录
        /// </summary>
        /// <param name="result"></param>
        /// <param name="document"></param>
        private void SetStatsticsValues(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            result.TestingStatstics = new List<TestinStatisticsDto>();
            var columnName = "A";
            var sheetName = "汇总";
            var falg = "校准结果";
            var startRowIndex = 1;
            var endRowIndex = 1000;
            var foundRowIndex = 0;
            for (var i = startRowIndex; i < endRowIndex; i++)
            {
                var excelAddress = columnName.ToString() + i.ToString();
                var cellValue = document.GetCellValue(sheetName, excelAddress);
                if (falg.Equals(cellValue))
                {
                    foundRowIndex = i + 1;
                    break;
                }
            }
            if (foundRowIndex == 0)
                throw new Exception($"在{sheetName}工作表里面的A列，未找到{falg}");

            result.TestingTemperatureDisplayValue = document.GetCellValue(sheetName, "B" + foundRowIndex.ToString());
            foundRowIndex += 2;//再指向到下三行
            var startIndex = foundRowIndex;

     
            while (true)
            {
                var cellValue = document.GetCellValue(sheetName, "A" + startIndex.ToString());
                if (string.IsNullOrEmpty(cellValue)) //如果名义温度为空的话，那就忽略此行
                    break;
                decimal _;
                if (!decimal.TryParse(cellValue, out _))//如果不是数值 ，那也为空
                    break;
                var obj = new TestinStatisticsDto();
                obj.温控设置 = BlackBodyHelpers.GetStaticeNumber(cellValue);

                cellValue = document.GetCellValue(sheetName, "B" + startIndex.ToString());
                obj.控温显示 = BlackBodyHelpers.GetStaticeNumber(cellValue);


                cellValue = document.GetCellValue(sheetName, "C" + startIndex.ToString());
                obj.亮度温度 = BlackBodyHelpers.GetStaticeNumber(cellValue);

                cellValue = document.GetCellValue(sheetName, "D" + startIndex.ToString());
                obj.U值 = BlackBodyHelpers.GetStaticeNumber(cellValue);

                cellValue = document.GetCellValue(sheetName, "E" + startIndex.ToString());
                obj.温度计示值 = BlackBodyHelpers.GetStaticeNumber(cellValue);

                cellValue = document.GetCellValue(sheetName, "F" + startIndex.ToString());
                obj.扩展不确定度 = BlackBodyHelpers.GetStaticeNumber(cellValue);



                if (obj.亮度温度 != "#DIV/0!")
                    result.TestingStatstics.Add(obj);
                startIndex++;
            }
        }

        private void SetFinallyNotes(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            result.FinallyNotes = new List<string>();
            var sheetName = "记录";
            var flag = document.GetCellValue(sheetName, "I2");
            flag = flag.Replace("：", ":");
            if (!flag.EndsWith(":"))
                flag += ":";

            sheetName = "数据说明";
            var startRowIndex = 1;
            while (true)
            {
                var value = document.GetCellValue(sheetName, "A" + startRowIndex.ToString());
                if (string.IsNullOrEmpty(value))
                    break;
                if (value.Replace("：", ":").StartsWith(flag))
                {
                    value = document.GetCellValue(sheetName, "B" + startRowIndex.ToString());
                    result.FinallyNotes.Add(value);
                }

                startRowIndex++;
            }
        }

    }
}

