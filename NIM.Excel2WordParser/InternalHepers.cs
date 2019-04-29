using NIM.Utilty;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator
{
  internal static  class InternalHepers
    {


          public static string ChangeTemperatureHumidityValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            var arr = value.Split('-');
            if (arr.Length != 2)
                return value;
            try
            {
                var a1 = NumberHelper.EnsureRound1(arr[0]);
                var a2 = NumberHelper.EnsureRound1(arr[1]);
                return a1 + "-" + a2;
            }
            catch
            {
                return value;
            }
        }
    }
}
