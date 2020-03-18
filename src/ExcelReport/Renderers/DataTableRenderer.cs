using ExcelReport.Contexts;
using ExcelReport.Driver;
using ExcelReport.Exceptions;
using ExcelReport.Extends;
using ExcelReport.Meta;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelReport.Renderers
{
    public class DataTableRenderer : Named, IElementRenderer
    {
        protected DataTable DataSource { get; set; }

        public DataTableRenderer(string name,DataTable dataSource)
        {
            Name = name;
            DataSource = dataSource;
        }
        public void Render(SheetContext sheetContext)
        {
            DataTabler dataTabler = sheetContext.WorksheetContainer.DataTablers[Name];

            if(DataSource.IsNull())
            {
                throw new RenderException($"The DataSource of DataTableRenderer <[{dataTabler.Name}] is empty");
            }

            sheetContext.FillDataTablerLayout(dataTabler, () =>
            {
                RenderDataTableHender(sheetContext, dataTabler);
                RenderDataTableBody(sheetContext, dataTabler);
            });
        }

        private void RenderDataTableHender(SheetContext sheetContext,DataTabler dataTabler)
        {
            var dataTableHeaderLocation = dataTabler.HeaderStart;

            var columns = DataSource.Columns;
            int columnCount = columns.Count;
            int rowHeaderLocation = dataTableHeaderLocation.RowIndex;
            int columnHeaderLocation = dataTableHeaderLocation.ColumnIndex + columns.Count - 1;

            for (int startH = dataTableHeaderLocation.ColumnIndex; startH < columnHeaderLocation; startH++)
            {
                ICell cell = sheetContext.GetCell(rowHeaderLocation, startH, copyBeforeColumnCell: true);
                if (null == cell)
                {
                    throw new RenderException($"DataTable[{dataTabler.Name}],cell[{rowHeaderLocation},{startH}] is null");
                }
            }
            sheetContext.MergeCells(rowHeaderLocation, rowHeaderLocation, dataTableHeaderLocation.ColumnIndex, columnHeaderLocation);

            ICell merageCell = sheetContext.GetCell(dataTableHeaderLocation);
            merageCell.Value = DataSource.TableName.CastTo<string>();
        }

        private void RenderDataTableBody(SheetContext sheetContext,DataTabler dataTabler)
        {
            var dataTableDataLocation = dataTabler.DataStart;

            var columns = DataSource.Columns;
            var rows = DataSource.Rows;
            int columnIndex = 0;
            int rowIndex = 0;

            foreach (DataColumn dataColumn in columns)
            {
                rowIndex = 0;
                foreach (DataRow dataRow in rows)
                {
                    int rowLocation = dataTableDataLocation.RowIndex + rowIndex;
                    int columnLocation = dataTableDataLocation.ColumnIndex + columnIndex;
                    ICell cell = sheetContext.GetCell(rowLocation, columnLocation, copyBeforeColumnCell: true);
                    if (null == cell)
                    {
                        throw new RenderException($"DataTable [{dataTabler.Name}],cell[{rowLocation},{columnLocation}] is null");
                    }
                    cell.Value = dataRow[dataColumn].CastTo<string>();
                    rowIndex += 1;
                }
                columnIndex += 1;
            }
        }

        public int SortNum(SheetContext sheetContext)
        {
            DataTabler EDataTableRender = sheetContext.WorksheetContainer.DataTablers[Name];

            if (EDataTableRender.HeaderStart.IsNull())
            {
                throw new TemplateException($"DataTableRenderer <[{EDataTableRender.Name}] header non-existent.");
            }
            if (EDataTableRender.DataStart.IsNull())
            {
                throw new TemplateException($"DataTableRenderer <[{EDataTableRender.Name}] data non-existent.");
            }

            EDataTableRender.HenderEnd = new Location(EDataTableRender.HeaderStart.RowIndex, EDataTableRender.HeaderStart.ColumnIndex + DataSource.Columns.Count - 1);
            EDataTableRender.DataEnd = new Location(EDataTableRender.DataStart.RowIndex + DataSource.Rows.Count - 1, EDataTableRender.DataStart.ColumnIndex + DataSource.Columns.Count - 1);

            return EDataTableRender.HeaderStart.RowIndex;
        }
    }
}
