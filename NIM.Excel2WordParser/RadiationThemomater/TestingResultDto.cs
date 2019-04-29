using NIM.CertificationGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.RadiationThemomater
{
    /// <summary>
    /// 测量结果
    /// </summary>
    public class TestingProcessResultDto
    {
        [SimpleExcelMappting("证书编号", "AA")]
        public string CertificationNo
        {
            get; set;
        }
        [SimpleExcelMappting("客户名称", "AA")]

        public string CustomerName
        {
            get; set;
        }
        [SimpleExcelMappting("器具名称", "AA")]

        public string ProductName
        {
            get; set;
        }
        [SimpleExcelMappting("型号/规格", "AA")]
        public string ModelSpecification
        {
            get; set;
        }
        [SimpleExcelMappting("出厂编号", "AA")]
        public string FactoryNo
        {
            get; set;
        }
        [SimpleExcelMappting("生产厂商", "AA")]
        public string FactoryName
        {
            get; set;
        }
        [SimpleExcelMappting("客户地址", "AA")]
        public string CustomerAddress
        {
            get; set;
        }
        [SimpleExcelMappting("校准日期", "AA")]
        public DateTime TestingDate
        {
            get; set;
        }

        [SimpleExcelMappting("测试环境温度", "AA")]
        public string TestingEnviomentTemperature
        {
            get; set;
        }
        /// <summary>
        /// 测试环境温度
        /// </summary>
        [SimpleExcelMappting("测试环境湿度", "AA")]
        public string TestingEnviomentHumidity
        {
            get; set;
        }
        [SimpleExcelMappting("测试环境地址", "AA")]

        public string TestingEnviomentAddress
        {
            get; set;
        }
        [SimpleExcelMappting("测试环境其他说明", "AA")]

        public string Others
        {
            get;
            set;
        }


        public List<TestingMachine> TestsingMachines
        {
            get; set;
        }


    }


    /// <summary>
    /// 测量装置/主要仪器
    /// </summary>
    public class TestingMachine
    {
        [SimpleExcelMappting("名称", "AA")]
        public string Name
        {
            get; set;
        }
        [SimpleExcelMappting("测量范围", "AA")]
        public string TestingRange
        {
            get; set;
        }
        [SimpleExcelMappting("不确定度", "AA")]
        public string UndeterminedStuff
        {
            get; set;
        }
        [SimpleExcelMappting("证书编号说明", "AA")]
        public string CertificattionNoNote
        {
            get; set;
        }
        [SimpleExcelMappting("证书有效期", "AA")]
        public DateTime CertificateExpired
        {
            get; set;
        }

        public List<TestinResultDto> Results
        {
            get; set;
        }

        public string Notes
        {
            get; set;
        }
    }



    /// <summary>
    /// 测量结果
    /// </summary>
    public class TestinResultDto
    {
        public int StandardTemperature
        {
            get; set;
        }
        public int ResultTemperature
        {
            get; set;
        }
        public int Difference
        {
            get; set;
        }
        public decimal ExtendUndeterminedStuff
        {
            get; set;
        }

        public string RadiationSourceDescription
        {
            get; set;
        }

    }
}
