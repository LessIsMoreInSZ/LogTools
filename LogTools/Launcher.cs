using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LogTools
{
    public class Launcher
    {
        // 用于记录日志的数据表
        private DataTable _dataTable = new DataTable();
        public DataTable CurrentTable 
        { 
            get { return _dataTable; } 
            set { _dataTable = value; }
        }

        public BlockingCollection<DataTable> DataTableQueue = new BlockingCollection<DataTable>();
        private CancellationTokenSource ctsQueue = new CancellationTokenSource();
        private CancellationTokenSource ctsIo = new CancellationTokenSource();

        public Launcher()
        {
            //DataIntoQueue();
            WriteIO();
        }

        public void DataIntoQueue()
        {
            Task.Run(() =>
            {
                var token = ctsQueue.Token;
                while (!token.IsCancellationRequested)
                {
                    if(CurrentTable.Rows.Count!=0)
                    {
                        DataTableQueue.Add(CurrentTable);
                        Console.WriteLine("in:" + CurrentTable.TableName);
                        //CurrentTable.Clear();
                    }
                }
            });
        }

        public void WriteIO()
        {
            Task.Run(() =>
            {
                var token = ctsIo.Token;
                while (!token.IsCancellationRequested)
                {
                    DataTable data= DataTableQueue.Take();
                    if (data!=null)
                    {
                        Console.WriteLine("out:" + data.TableName);
                        ExeclHelper.TableToExcel(data, "D:\\Test\\" + data.TableName + ".xls");
                        Thread.Sleep(100);
                    }
                }
            });
        }

    }
}
