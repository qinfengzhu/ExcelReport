﻿using ExcelReport.Driver.NPOI.Extends;
using System.Collections;
using System.Collections.Generic;
using NpoiCell = NPOI.SS.UserModel.ICell;
using NpoiRow = NPOI.SS.UserModel.IRow;

namespace ExcelReport.Driver.NPOI
{
    public class Row : IRow
    {
        public NpoiRow NpoiRow { get; private set; }

        public Row(NpoiRow npoiRow)
        {
            NpoiRow = npoiRow;
        }

        public ICell this[int columnIndex] => NpoiRow.GetCell(columnIndex).GetAdapter();

        public IEnumerator<ICell> GetEnumerator()
        {
            foreach (NpoiCell npoiCell in NpoiRow)
            {
                yield return npoiCell.GetAdapter();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return NpoiRow;
        }

        public ICell CopyCell(int sourceIndex, int targetIdnex)
        {
            var sourceCell = NpoiRow.GetCell(sourceIndex);
            var targetCell = NpoiRow.CopyCell(sourceIndex, targetIdnex);

            if (null != sourceCell.CellStyle)
            {
                targetCell.CellStyle = sourceCell.CellStyle;
            }
            targetCell.SetCellType(sourceCell.CellType);
            ICell cell = targetCell.GetAdapter();
            return cell;
        }
    }
}