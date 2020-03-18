using ExcelReport.Driver.NPOI.Extends;
using NPOI.Extend;
using NPOI.SS.Util;
using System.Collections;
using System.Collections.Generic;
using NpoiRow = NPOI.SS.UserModel.IRow;
using NpoiSheet = NPOI.SS.UserModel.ISheet;

namespace ExcelReport.Driver.NPOI
{
    public class Sheet : ISheet
    {
        public NpoiSheet NpoiSheet { get; private set; }

        public Sheet(NpoiSheet npoiSheet)
        {
            NpoiSheet = npoiSheet;
        }

        public IRow this[int rowIndex] => NpoiSheet.GetRow(rowIndex).GetAdapter();

        public string SheetName => NpoiSheet.SheetName;

        public int CopyRows(int start, int end)
        {
            return NpoiSheet.CopyRows(start, end);
        }

        public int RemoveRows(int start, int end)
        {
            return NpoiSheet.RemoveRows(start, end);
        }

        public IEnumerator<IRow> GetEnumerator()
        {
            foreach (NpoiRow npoiRow in NpoiSheet)
            {
                yield return npoiRow.GetAdapter();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return NpoiSheet;
        }

        public void MergeCells(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            NpoiSheet.AddMergedRegion(region);
        }
    }
}