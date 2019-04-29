using NIM.CertificationGenerator.Core;
using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.RadiationThermomater
{
    //public enum CertificationType
    //{
    //    /// <summary>
    //    /// 校准
    //    /// </summary>
    //    Calibrating,
    //    /// <summary>
    //    /// 工作用
    //    /// </summary>
    //    Woring
    //}
    /// <summary>
    /// 测量结果
    /// </summary>
    public class TestingProcessResultDto : ITestingResultDto
    {
        //[WordKey]
        //[SimpleExcelMaping("数据导入标识符", ExcelAddressName = "记录:F2")]
        //public CertificationType CertificationType
        //{
        //    get; set;
        //}

        [SimpleExcelMaping("光阑直径", ExcelAddressName = "记录:F7")]
        public string 光阑直径
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
                }
                return this.isDisplayResolutionInt.Value;
            }
        }

        [SimpleExcelMaping("数据说明", ExcelAddressName = "记录:I2")]
        public string FinallyNotesFlag
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("证书编号", ExcelAddressName = "记录:M3")]
        public string CertificationNo
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("客户名称", ExcelAddressName = "记录:M6")]

        public string CustomerName
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("器具名称", ExcelAddressName = "记录:B2")]

        public string ProductName
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("型号/编号", ExcelAddressName = "记录:B3")]
        public string ModelSpecification
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
        [SimpleExcelMaping("生产厂商", ExcelAddressName = "记录:B5")]
        public string FactoryName
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("客户地址", ExcelAddressName = "记录:M7")]
        public string CustomerAddress
        {
            get; set;
        }

        [SimpleExcelMaping("校准日期", ExcelAddressName = "记录:I3")]
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
        [SimpleExcelMaping("测试环境温度", ExcelAddressName = "记录:I6")]
        public string TTemperature
        {
            get; set;
        }
        [WordKey]
        /// <summary>
        /// 测试环境温度
        /// </summary>
        [SimpleExcelMaping("测试环境湿度", ExcelAddressName = "记录:I7")]
        public string THumidity
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("测试环境地址", ExcelAddressName = "记录:I4")]

        public string TAddress
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("测试环境其他说明", ExcelAddressName = "记录:I8")]

        public string TOthers
        {
            get;
            set;
        }


        public List<TestingTool> TestingTools
        {
            get; set;
        }

        public List<TestinStatisticsDto> TestingStatstics
        {
            get; set;
        }
        public List<string> FinallyNotes
        {
            get; set;
        }

    }


    public class TestingTool
    {
        public string Name
        {
            get; set;
        }

        public string TempRange
        {
            get; set;
        }

        public string NotSure
        {
            get; set;
        }

        public string Code
        {
            get; set;
        }


        public string ExpiredDesc
        {
            get; set;
        }

    }




    /// <summary>
    /// 测量结果
    /// </summary>
    public class TestinStatisticsDto
    {
        public string Name
        {
            get; set;
        }
        public string StandardTemperature
        {
            get; set;
        }
        public string ResultTemperature
        {
            get; set;
        }
        public string Difference
        {
            get; set;
        }

        public string ExtendDifference
        {
            get; set;
        }
        public string Note
        {
            get; set;
        }


    }
}
