using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter
{
    public static class ExtensionMethods
    {
        public static Color ColorRGB(this Excel.Range rng)
        {
            return ColorTranslator.FromOle(Convert.ToInt32(rng.Interior.Color));
        }

        public static bool ContainsRange(this Excel.Range rng, Excel.Range range)
        {
            Excel.Range r1 = ((Excel.Range)rng.Cells[1, 1]);
            Excel.Range r2 = ((Excel.Range)rng.Cells[1, rng.Columns.Count]);
            Excel.Range r3 = ((Excel.Range)rng.Cells[rng.Rows.Count, 1]);

            int x1, x2, y1, y2;
            x1 = r1.Column;
            x2 = r2.Column;

            y1 = r1.Row;
            y2 = r3.Row;

            Marshal.ReleaseComObject(r1);
            Marshal.ReleaseComObject(r2);
            Marshal.ReleaseComObject(r3);

            if (range.Column >= x1 && range.Column <= x2 && range.Row >= y1 && range.Row <= y2)
                return true;
            else
                return false;
        }
    }
    //public class RangesClass
    //{
    //    Excel.Range range;
    //    public RangesClass(Excel.Range range)
    //    {
    //        this.range = range;
    //    }

    //    public IEnumerator<Excel.Range> GetEnumerator() => new RangesEnumerator(range);
    //}
    internal static class ControlExtensions
    {
        public static ForTags? SpecialTag(this Control control)
        {
            if ((ForTags)control.Tag != null)
            {
                return (ForTags)control.Tag;
            }
            else
            {
                return null;
            }
        }
    }
}
