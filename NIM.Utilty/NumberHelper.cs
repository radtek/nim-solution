using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.Utilty
{
    public static class NumberHelper
    {
        public static double Round(this double v,int roundParameter)
        {
            return Math.Round(v, roundParameter, MidpointRounding.AwayFromZero);
        }
        public static decimal Round(this decimal v,int roundParameter)
        {
            return Math.Round(v, roundParameter, MidpointRounding.AwayFromZero);
        }

        public static string EnsureRound1(this string str)
        {
            int i;
            if (int.TryParse(str, out i))
            {
                return str + ".0";
            }
            return str;
        }


        public static bool IsChinese(this string text)
        {

            char[] c = text.ToCharArray();

            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] >= 9312 && c[i] <= 9331) //9312: '①' 9331 '⑳'
                    return true;
                if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsDisplayResolutionInt(string displayResolutionValue)
        {

            //David 2017-07-14
            //规则：只有当显示分辨力的值为整数时，才不留小数位，否则就保留一位小数
            displayResolutionValue = displayResolutionValue.Replace("°C", "").Trim();
            int _;
            if (int.TryParse(displayResolutionValue, out _)) //如果值为整数，比如，1，2，3，4..
                return true;
            else
                return false;
        }
    }
}
