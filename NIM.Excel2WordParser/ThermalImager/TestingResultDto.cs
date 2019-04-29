using NIM.CertificationGenerator.Core;
using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.ThermalImager
{

    public enum FileType
    {
        Range1,
        Range2,
        Range3,
        Range4
    }
    public class TestingProcessResultDto : ITestingResultDto
    {

        [WordKey]
        [SimpleExcelMaping("器具名称", ExcelAddressName = "记录:B2")]

        public string ProductName
        {
            get; set;
        }


        [SimpleExcelMaping("证书类型", ExcelAddressName = "记录:F2")]
        public string CertificateType
        {
            get; set;
        }
        public FileType FileType
        {
            get
            {
                if (this.CertificateType == "校准-单一量程")
                    return FileType.Range1;
                else if (this.CertificateType == "校准-两量程")
                    return FileType.Range2;
                else if (this.CertificateType == "校准-三量程")
                    return FileType.Range3;
                else if (this.CertificateType == "校准-四量程")
                    return FileType.Range4;
                throw new Exception("非法的证书类型，必须为 校准-单一量程,校准-两量程,校准-三量程,校准-四量程 (" + this.CertificateType + ")");
            }
        }

        [SimpleExcelMaping("数据说明", ExcelAddressName = "记录:K2")]
        public string FinallyNotesFlag
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("型号/规格", ExcelAddressName = "记录:B3")]
        public string ModelSpecification
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("光谱范围", ExcelAddressName = "记录:F3")]
        public string SpectralRange
        {
            get; set;
        }

        [SimpleExcelMaping("实验日期", ExcelAddressName = "记录:K3")]
        public DateTime TestingDate
        {
            get; set;
        }

        [WordKey]
        public string TestingDateStr
        {
            get
            {
                return string.Format("{0:yyyy}年{0:MM}月{0:dd}日", this.TestingDate);
            }
        }

        [WordKey]
        [SimpleExcelMaping("证书编号", ExcelAddressName = "记录:Q3")]
        public string CertificationNo
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("出厂编号", ExcelAddressName = "记录:B4")]
        public string FactoryNo
        {
            get; set;
        }


        [WordKey]
        [SimpleExcelMaping("测量距离", ExcelAddressName = "记录:F4")]
        public string MeasureDistances
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("实验地点", ExcelAddressName = "记录:K4")]
        public string TestingPlace
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("生产厂商", ExcelAddressName = "记录:B5")]
        public string FactoryName
        {
            get; set;
        }


        [WordKey]
        [SimpleExcelMaping("测试环境温度", ExcelAddressName = "记录:K6")]
        public string TTemperature
        {
            get; set;
        }


        [WordKey]
        [SimpleExcelMaping("客户名称", ExcelAddressName = "记录:Q6")]

        public string CustomerName
        {
            get; set;
        }

        
        [WordKey]
        [SimpleExcelMaping("测试环境湿度", ExcelAddressName = "记录:K7")]
        public string THumidity
        {
            get; set;
        }


        [WordKey]
        [SimpleExcelMaping("客户地址", ExcelAddressName = "记录:Q7")]
        public string CustomerAddress
        {
            get; set;
        }



        [SimpleExcelMaping("显示分辨力", ExcelAddressName = "记录:B8")]
        public string DisplayResolutionValue
        {
            get; set;
        }
        private bool? isDisplayResolutionInt;
        public bool IsDisplayResolutionInt
        {
            get
            {
                if (this.isDisplayResolutionInt == null)
                {
                    this.isDisplayResolutionInt = NumberHelper.IsDisplayResolutionInt(this.DisplayResolutionValue);

                    var v = this.DisplayResolutionValue.Replace("°C", "").Trim();
                    int _;
                    if (int.TryParse(v, out _)) //如果值为整数，比如，1，2，3，4..
                        this.isDisplayResolutionInt = true;
                    else
                        this.isDisplayResolutionInt = false;
                }
                return this.isDisplayResolutionInt.Value;
            }
        }

        [WordKey]
        [SimpleExcelMaping("其他", ExcelAddressName = "记录:K8")]
        public string TOther
        {
            get; set;
        }
  
        public string FinallyNotes
        {
            get; set;
        }



        public List<string> TemperatureRanges
        {
            get; set;
        }

        public List<TemperatureRange> Ranges
        {
            get; set;
        }

        public List<PointTemperature> TestingPointValues
        {
            get; set;
        }



    }

    public class PointTemperature
    {
        public string Location
        {
            get;set;
        }
        public string Value
        {
            get;set;
        }
    }

    public enum RangetType
    {
        范围1 = 1,
        范围2 = 2,
        范围3 = 3,
        范围4 = 4
    }

    /// <summary>
    /// 测量结果
    /// </summary>
    public class TemperatureRange
    {
        public double StandardTemperature
        {
            get; set;
        }
        public string Description
        {
            get; set;
        }
        //温度
        public string DisplayStandardTemperature
        {
            get; set;
        }

        //示值误差
        public string Difference
        {
            get; set;
        }
        //U
        public string UValue
        {
            get; set;
        }
        public RangetType Type
        {
            get; set;
        }
        //—
    }
}

