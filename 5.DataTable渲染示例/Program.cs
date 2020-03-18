using ExcelReport;
using ExcelReport.Driver.NPOI;
using ExcelReport.Renderers;
using System;
using System.Data;

namespace _5.DataTable渲染示例
{
    class Program
    {
        static void Main(string[] args)
        {
            Configurator.Put(".xls", new WorkbookLoader());

            try
            { 
                var dataTable = GetDataTable();

                ExportHelper.ExportToLocal(@"Template\Template.xls", "out.xls",
                             new SheetRenderer("访视视图",
                                    new DataTableRenderer("Plants", dataTable)
                                )
                    );
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("finished!");
        }
        private static DataTable GetDataTable()
        {
            DataTable dt = new DataTable("试验流程图");
            dt.Columns.Add(new DataColumn("研究阶段", typeof(string)));
            dt.Columns.Add(new DataColumn("数据集", typeof(string)));
            dt.Columns.Add(new DataColumn("筛选期\r\n(第-14~1天)", typeof(string)));
            dt.Columns.Add(new DataColumn("第一周期\r\n(第-1天)", typeof(string)));

            var row = dt.NewRow();
            row["研究阶段"] = "是否重复访视";
            row["数据集"] = "";
            row["筛选期\r\n(第-14~1天)"] = "否";
            row["第一周期\r\n(第-1天)"] = "否";
            dt.Rows.Add(row);

            var row1 = dt.NewRow();
            row1["研究阶段"] = "访视编号";
            row1["数据集"] = "";
            row1["筛选期\r\n(第-14~1天)"] = "1";
            row1["第一周期\r\n(第-1天)"] = "2";
            dt.Rows.Add(row1);

            return dt;
        }
    }
}
