using FastStream.office;
using FastStream.socket;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            EecelTool tool = new EecelTool();
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("col1");
            dataTable.Columns.Add("col2");
            dataTable.Columns.Add("col3");
            for (int i = 0; i < 100; i++)
            {
                var row = dataTable.NewRow();
                row[0] = "jin";
                row[1] = "yu";
                row[2] = 345;
            }
            tool.ExportTable("sss.xlsx", new List<DataTable>() { dataTable });

        }
    }
}
