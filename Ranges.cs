using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter
{
    //class RangesEnumerator : IEnumerator<Excel.Range>
    //{
    //    List<Excel.Range> _ranges;
    //    int position = -1;
    //    public RangesEnumerator(Excel.Range range)
    //    {
    //        int rows = range.Rows.Count;
    //        int cols = range.Columns.Count;
    //        _ranges = new List<Excel.Range>();
    //        for (int y = 0; y < rows; y++)
    //        {
    //            for (int x = 0; x < cols; x++)
    //            {
    //                _ranges.Add((Excel.Range)range.Cells[y, x]);
    //            }
    //        }
    //    }
    //    public Excel.Range Current
    //    {
    //        get
    //        {
    //            if (position == -1 || position >= _ranges.Count)
    //                throw new ArgumentException();
    //            return _ranges[position];
    //        }
    //    }
    //    object IEnumerator.Current => throw new NotImplementedException();
    //    public bool MoveNext()
    //    {
    //        if (position < _ranges.Count - 1)
    //        {
    //            position++;
    //            return true;
    //        }
    //        else
    //            return false;
    //    }
    //    public void Reset() => position = -1;
    //    public void Dispose() { }
    //}
}
