using System.Collections.Generic;

namespace ExcelReport.Driver
{
    public interface IRow : IEnumerable<ICell>, IAdapter
    {
        ICell this[int columnIndex]
        {
            get;
        }
        ICell CopyCell(int sourceIndex, int targetIdnex);
    }
}