using NIM.CertificationGenerator.Core;
using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.BlackBodySingleRadiationiation
{
    /// <summary>
    /// 黑体测试数据
    /// </summary>
    public class TestingProcessResultDto : ITestingResultDto          , IAdvanceValue
    {
        public bool IsAdvance
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("证书编号", ExcelAddressName = "汇总:J3")]
        public string CertificationNo
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("客户名称", ExcelAddressName = "汇总:J6")]

        public string CustomerName
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("器具名称", ExcelAddressName = "汇总:B2")]
        public string ProductName
        {
            get; set;
        }

        [WordKey]
        [SimpleExcelMaping("型号/规格", ExcelAddressName = "汇总:B3")]
        public string ModelSpecification
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("出厂编号", ExcelAddressName = "汇总:B4")]
        public string FactoryNo
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("生产厂商", ExcelAddressName = "汇总:B5")]
        public string FactoryName
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("客户地址", ExcelAddressName = "汇总:J7")]
        public string CustomerAddress
        {
            get; set;
        }

        [SimpleExcelMaping("校准日期", ExcelAddressName = "汇总:F3")]
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
        [WordKey(Key = "{{TT}}")]
        [SimpleExcelMaping("环境温度", ExcelAddressName = "汇总:F6")]
        public string TTemperature
        {
            get; set;
        }
        [WordKey(Key = "{{TH}}")]
        /// 测试环境温度
        /// </summary>
        [SimpleExcelMaping("环境湿度", ExcelAddressName = "汇总:F7")]
        public string THumidity
        {
            get; set;
        }
        [WordKey]
        [SimpleExcelMaping("校准地点", ExcelAddressName = "汇总:F4")]

        public string TAddress
        {
            get; set;
        }

        [WordKey]
        public string TOther => "/";
        [WordKey(Key = "{{DisV}}")]
        public string TestingTemperatureDisplayValue
        {
            get; set;
        }

        //标准器具
        public List<TestingTool> TestingTools
        {
            get; set;
        }
        //测试数据

        public List<TestinStatisticsDto> TestingStatstics
        {
            get; set;
        }
        public List<string> FinallyNotes
        {
            get; set;
        }
        public AdvanceValue AdvanceValue { get; set; }
    }


    public class TestingTool
    {
        //标准器名称
        public string Name
        {
            get; set;
        }
        //测量范围
        public string TempRange
        {
            get; set;
        }
        //不确定等级
        public string NotSure
        {
            get; set;
        }
        //证书编号
        public string Code
        {
            get; set;
        }

        //有效期
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
        //名义温度
        public string 温控设置
        {
            get; set;
        }

        public string 控温显示
        {
            get; set;
        }
        public string 亮度温度
        {
            get; set;
        }
        public string U值
        {
            get; set;
        }
    }
}