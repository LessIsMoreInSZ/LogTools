using LogTools;
using System.Data;

namespace ConsoleTestApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello, World!");
            Launcher launcher = new Launcher();
            for (int i = 0; i < 50; i++)
            {
                DataTable dataTable = new DataTable();
                dataTable.TableName = "Test" + i.ToString();
                dataTable.Columns.Add("Test1");
                dataTable.Columns.Add("Test2");

                for (int j = 0; j < 1000; j++)
                {
                    DataRow row = dataTable.NewRow();
                    row[0] = "1";
                    row[1] = "2";
                    dataTable.Rows.Add(row);
                }

                //launcher.CurrentTable = dataTable;
                launcher.DataTableQueue.Add(dataTable);
            }

            Console.ReadLine();
        }
    }
}
