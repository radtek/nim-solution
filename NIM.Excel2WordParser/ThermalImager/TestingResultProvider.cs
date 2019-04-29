using DocumentFormat.OpenXml.Packaging;
using NIM.CertificationGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NIM.Utilty;

namespace NIM.CertificationGenerator.ThermalImager
{
    public class TestingResultProvider
    {
        private const int StatsticsStartIndex = 13;//检测数据从第12行开始
        private const string StatsticsSheetName = "记录";
        public TestingProcessResultDto GetResult(SpreadsheetDocument document)
        {

            var result = new TestingProcessResultDto();
            var propertyDescriptions = PropertyInfoBulider<TestingProcessResultDto>.GetPropertyAttributes();
            propertyDescriptions.ForEach(propertyDescription =>
            {
                this.SetPropertyValue(result, document, propertyDescription);

            });

            this.SetStatsticsValues(result, document);
            this.SetFinallyNotes(result, document);
            this.SetPointsValues(result, document);
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
                    decimal s;
                    if (decimal.TryParse(cellValue, out s))
                    {
                        value = NumberHelper.Round(s, 1).ToString().EnsureRound1();
                        //s.Round(1).ToString().EnsureRound1();

                        //value = s.Round(1).ToString().EnsureRound1();

                        //var _value = System.Math.Round(s, 4).ToString();
                        //if (_value.IndexOf('.') >= 0) //如果不是整数的话
                        //    _value = _value.TrimEnd('0');
                        //value = _value.EnsureRound1();

                        //  value = valud.ToString().Trim('0').EnsureRound1();

                        // value = System.Math.Round(s, 4).ToString().TrimEnd('0').EnsureRound1();
                    }
                    else
                        value = cellValue.EnsureRound1();
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


        private void SetFinallyNotes(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            result.FinallyNotes = "";
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
                    {
                        result.FinallyNotes = value;
                        break;
                    }
                }
                startRowIndex++;
            }
        }

        /// <summary>
        /// 设置具体测试记录
        /// </summary>
        /// <param name="result"></param>
        /// <param name="document"></param>
        private void SetStatsticsValues(TestingProcessResultDto result, SpreadsheetDocument document)
        {
            result.Ranges = new List<TemperatureRange>();
            //注意，从STUV这四列中取得数据 
            var startRowIndex = StatsticsStartIndex;
            startRowIndex = startRowIndex - 1;
            var lastDescription = "";
            RangetType? lastRangeType = null;

            var endFalg = "不确定度"; //查找关键字
            var toRowIndex = 1000; //扫描最大的行数

            while (startRowIndex <= toRowIndex) //退出循环的标志就是S列的值为空
            {
                startRowIndex++;
                var address = "A" + startRowIndex.ToString();
                var str = document.GetCellValue(StatsticsSheetName, address);
                if (str == endFalg)
                {
                    break;
                }
                var addressName = "W" + startRowIndex.ToString();
                var description = document.GetCellValue(StatsticsSheetName, addressName);
                if (result.FileType != FileType.Range1) //因为一量程，是没有温度这一列的，所以一量程是不需要检测是的
                {
                    if (string.IsNullOrEmpty(description)) //如果当前的温度范围为空，那就会取自上次的温度范围，如果上次的温度范围也为空，那么就说明数据有问题
                    {
                        if (string.IsNullOrEmpty(lastDescription))
                            throw new Exception("温度范围的第一行，不能为空");
                        description = lastDescription; //应用上次的温度范围
                    }
                    else //如果本次有值，那么，如果温度范围的描述不相同，那就是进入了下一个温度范围了。
                    {
                        if (description != lastDescription)
                        {
                            if (lastRangeType == null)
                                lastRangeType = RangetType.范围1;
                            else if (lastRangeType.Value == RangetType.范围1)
                                lastRangeType = RangetType.范围2;
                            else if (lastRangeType.Value == RangetType.范围2)
                                lastRangeType = RangetType.范围3;
                            else if (lastRangeType.Value == RangetType.范围3)
                                lastRangeType = RangetType.范围4;
                        }
                    }
                    //用本次有值的描述来替换
                    lastDescription = description;
                }

                addressName = "T" + startRowIndex.ToString();
                var cellValue = document.GetCellValue(StatsticsSheetName, addressName);
                //如果T列的值的前景色是白色的话，那就忽略不在证书中显示
                if (!document.GetHasWhiteTheme(StatsticsSheetName, addressName))
                {
                    double d;
                    if (!double.TryParse(cellValue, out d))
                    {
                        continue;
                    }

                    var obj = new TemperatureRange();
                    obj.DisplayStandardTemperature = cellValue.EnsureRound1();


                    obj.StandardTemperature = d;

                    addressName = "U" + startRowIndex;

                    cellValue = document.GetCellValue(StatsticsSheetName, addressName);
                    double difference;
                    if (cellValue == "#DIV/0!")
                        continue;
                    if (!double.TryParse(cellValue, out difference))
                        throw new Exception($"在{StatsticsSheetName}中的CELL {addressName} 值必须为一个数字");
                    var roundParameter = 1;
                    if (result.IsDisplayResolutionInt)
                        roundParameter = 0;
                    obj.Difference = difference.Round(roundParameter).ToString();
                    if (!result.IsDisplayResolutionInt)
                        obj.Difference = obj.Difference.EnsureRound1();

                    addressName = "V" + startRowIndex.ToString();
                    double _uValue;
                    var uValue = document.GetCellValue(StatsticsSheetName, addressName);
                    if (double.TryParse(uValue, out _uValue))
                        obj.UValue = _uValue.Round(1).ToString();
                    else
                        obj.UValue = uValue;
                    obj.UValue = obj.UValue.EnsureRound1();

                    obj.Description = description;
                    if (result.FileType == FileType.Range1)
                        obj.Type = RangetType.范围1;
                    else
                        obj.Type = lastRangeType.Value;



                    result.Ranges.Add(obj);
                }

            }



        }


        private void SetPointsValues(TestingProcessResultDto result, SpreadsheetDocument document)
        {

            var roundParameter = 1;
            if (result.IsDisplayResolutionInt)
                roundParameter = 0;

            var toRowIndex = 1000; //扫描最大的行数

            var valueFlag = "测温一致性（°C）";


            var startRowIndex = StatsticsStartIndex;


            while (startRowIndex <= toRowIndex)
            {
                var addressName = "M" + startRowIndex;
                if (document.GetCellValue(StatsticsSheetName, addressName) == valueFlag)
                    break;
                startRowIndex++;
            }
            if (startRowIndex == 1)
                throw new Exception($"在{StatsticsSheetName}的表格里面的M列，未找到 测温一致性（°C） 关键字");

            startRowIndex++;
            startRowIndex++;
            result.TestingPointValues = new List<PointTemperature>();
            for (var i = 0; i < 9; i++)
            {
                var addressName = "M" + (startRowIndex + i).ToString();
                var location = document.GetCellValue(StatsticsSheetName, addressName);
                if (string.IsNullOrEmpty(location))
                    break;
                addressName = "P" + (startRowIndex + i).ToString();
                var value = document.GetCellValue(StatsticsSheetName, addressName);
                if (string.IsNullOrEmpty(value))
                    break;

                double v;
                if (double.TryParse(value, out v))
                {
                    value = v.Round(roundParameter).ToString();
                    if (roundParameter == 1)
                        value = value.EnsureRound1();
                }
                else
                {
                    throw new Exception("测温一致性（°C）的值不正确，CELL位置：" + addressName + ",Value:" + value);
                }

                result.TestingPointValues.Add(new PointTemperature
                {
                    Location = location,
                    Value = value

                });

            }


        }



    }
}

