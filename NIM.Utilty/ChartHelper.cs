using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using asp = Aspose.Cells;

namespace NIM.Utilty
{
    public class ChartHelper
    {

        public static void ToImage(asp.Charts.Chart chart, string imagePath, System.Drawing.Imaging.ImageFormat imageFormat)
        {
            LicenseHelper.ModifyInMemory.EnsureActivateMemoryPatching();

            chart.ToImage(imagePath, imageFormat);

        }
    }
}
