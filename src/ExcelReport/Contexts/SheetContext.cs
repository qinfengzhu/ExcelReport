using ExcelReport.Accumulations;
using ExcelReport.Driver;
using ExcelReport.Extends;
using ExcelReport.Meta;
using System;

namespace ExcelReport.Contexts
{
    public sealed class SheetContext
    {
        private readonly RowIndexAccumulation _rowIndexAccumulation = new RowIndexAccumulation();
        private readonly ISheet _sheet;

        public SheetContainer WorksheetContainer { get; }

        public bool IsEmpty()
        {
            return _sheet.IsNull();
        }

        public SheetContext(ISheet sheet, SheetContainer worksheetContainer)
        {
            _sheet = sheet;
            WorksheetContainer = worksheetContainer;
        }

        public ICell GetCell(Location location)
        {
            var rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(location.RowIndex);

            IRow row = _sheet[rowIndex];
            if (row.IsNull())
            {
                return null;
            }
            return row[location.ColumnIndex];
        }

        public ICell GetCell(int row, int column, bool copyBeforeColumnCell = false)
        {
            IRow r = _sheet[row];

            if (!copyBeforeColumnCell)
            {
                if (r.IsNull())
                {
                    return null;
                }
            }
            else
            {
                if (column - 1 >= 0)
                {
                    r.CopyCell(column - 1, column);
                }
            }

            return r[column];
        }

        public void CopyRepeaterTemplate(Repeater repeater, Action processTemplate)
        {
            var startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.Start.RowIndex);
            var endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.End.RowIndex);

            int span = _sheet.CopyRows(startRowIndex, endRowIndex);
            ICell startCell = GetCell(repeater.Start);
            startCell.Value = startCell.GetStringValue().CutEndOf($"<[{repeater.Name}]");
            processTemplate();
            ICell endCell = GetCell(repeater.End);
            endCell.Value = endCell.GetStringValue().CutStartOf($">[{repeater.Name}]");
            _rowIndexAccumulation.Add(span);
        }

        public void RemoveRepeaterTemplate(Repeater repeater)
        {
            var startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.Start.RowIndex);
            var endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.End.RowIndex);

            int span = _sheet.RemoveRows(startRowIndex, endRowIndex);
            _rowIndexAccumulation.Add(-span);
        }

        public void FillDataTablerLayout(DataTabler dataTabler, Action processDataTableFilled)
        {
            var startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(dataTabler.DataStart.RowIndex);
            var endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(dataTabler.DataEnd.RowIndex);

            ICell headerCell = GetCell(dataTabler.HeaderStart);
            headerCell.Value = headerCell.GetStringValue().CutEndOf($"$<{dataTabler.Name}>[header]");

            ICell dataStartCell = GetCell(dataTabler.DataStart);
            dataStartCell.Value = dataStartCell.GetStringValue().CutEndOf($"$<{dataTabler.Name}>[data]");

            for (int s = 0, l = endRowIndex - startRowIndex; s < l; s++)
            {
                _sheet.CopyRows(startRowIndex + s, startRowIndex + s);
            }

            processDataTableFilled();

            _rowIndexAccumulation.Add(endRowIndex - startRowIndex + 1);
        }

        public void MergeCells(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            if (firstRow < 0 || lastRow < firstRow || firstCol < 0 || lastCol < firstCol)
            {
                return;
            }
            _sheet.MergeCells(firstRow, lastRow, firstCol, lastCol);
        }
    }
}