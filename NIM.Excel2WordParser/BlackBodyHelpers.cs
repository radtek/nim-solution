using NIM.CertificationGenerator.Core;
using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator
{
    public interface IAdvanceValue
    {
        AdvanceValue AdvanceValue { get; set; }
    }
    public class AdvanceValue
    {
        [SimpleExcelMaping("环境湿度", ExcelAddressName = "汇总:F7")]
        public int Temperature1 { get; set; }
        public decimal TopValue1 { get; set; }
        public decimal ButtomValue1 { get; set; }
        public decimal LeftValue1 { get; set; }
        public decimal RightValue1 { get; set; }
        public decimal AvgValue1 { get; set; }


        public int Temperature2 { get; set; }
        public decimal TopValue2 { get; set; }
        public decimal ButtomValue2 { get; set; }
        public decimal LeftValue2 { get; set; }
        public decimal RightValue2 { get; set; }
        public decimal AvgValue2 { get; set; }

        public int Temperature3 { get; set; }
        public decimal TopValue3 { get; set; }
        public decimal ButtomValue3 { get; set; }
        public decimal LeftValue3 { get; set; }
        public decimal RightValue3 { get; set; }
        public decimal AvgValue3 { get; set; }

    }
    public class BlackBodyHelpers
    {
        public static string GetStaticeNumber(string str)
        {
            decimal v;
            double d;
            if (!double.TryParse(str, out d))
                return str;
            v = (decimal)d;
            v = NumberHelper.Round(v, 1);
            v = ((decimal)((int)(v * 10)) + 0.00m) / 10m;
            return Math.Round(v, 1).ToString();
        }



    }
}
