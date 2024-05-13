using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace LogTools
{
    public class ExeclHelper
    {
        public static List<T> DataTableToList<T>(DataTable dt) where T : new()
        {
            var list = new List<T>();
            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    T t = new T();
                    var properties = t.GetType().GetProperties();
                    foreach (var property in properties)
                    {
                        if (dt.Columns.Contains(property.Name)) //&& dr[property.Name] != DBNull.Value
                        {
                            if (!property.CanWrite) continue;
                            var value = dr[property.Name];
                            if (value != null && value.ToString() != string.Empty)
                            {
                                if (property.PropertyType == typeof(string))
                                    property.SetValue(t, value.ToString(), null);
                                if (property.PropertyType == typeof(Int32))
                                    property.SetValue(t, Int32.Parse(value.ToString()), null);
                                if (property.PropertyType == typeof(Double))
                                    property.SetValue(t, double.Parse(value.ToString()), null);
                            }
                        }
                    }

                    //if (dr[0].ToString()!=string.Empty) //在dataTable处判断空行
                    list.Add(t);
                    Trace.WriteLine(list.Count);
                }
            }
            catch (Exception ex)
            {

                throw;
            }

            return list;
        }


        public static DataTable ListToDataTable<T>(List<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                    if (values[i].GetType() == typeof(string) && values[i].ToString() == "")
                    {
                        values[i] = " ";
                    }
                }

                if (!(values[3].ToString().Equals("0") && values[4].ToString().Equals("0") && values[7].ToString().Equals("0")))
                    tb.Rows.Add(values);
            }

            return tb;
        }

        public static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        public static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="fileName">文件所在路径</param>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public static List<T> ExcelToDataTable<T>(string fileName, string sheetName, bool isFirstRowColumn) where T : new()
        {
            IWorkbook workbook = null;
            FileStream fs = null;
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    if (firstRow == null)
                    {
                        return null;
                    }
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);

                            if (cell != null)
                            {
                                cell.SetCellType(CellType.String);
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                                dataRow[j] = cell.DateCellValue.ToString();
                            else if (cell != null && cell.CellType == CellType.Formula)
                            {
                                dataRow[j] = row.GetCell(j).NumericCellValue;
                            }
                            else
                            {
                                //dataRow[j] = row.GetCell(j).ToString();
                                if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            //Trace.WriteLine($"i:{i}+j:{j}");
                        }

                        // Execl中去除空行，pinindex为空直接弃用
                        if (!string.IsNullOrEmpty(dataRow[0].ToString()))
                            data.Rows.Add(dataRow);
                    }
                }

                DataTable dt = data;
                return DataTableToList<T>(data);
                //return data;
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception: " + ex.Message);
                throw;
            }
        }


        /// <summary>
        /// Datable导出成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file">导出路径(包括文件名与扩展名)</param>
        public static void TableToExcel(DataTable dt, string file)
        {
            bool result = false;
            IWorkbook workbook;
            FileStream fs = null;
            IRow row;
            ISheet sheet;
            ICell cell;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    // 新建工作簿对象
                    workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet1的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    //设置列头
                    row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    //using (fs = File.CreateText(file))
                    //{
                    //    workbook.Write(fs);//向打开的这个xls文件中写入数据
                    //    result = true;
                    //}

                    //写入文件
                    FileStream xlsfile = new FileStream(file, FileMode.Create);
                    workbook.Write(xlsfile);
                    xlsfile.Close();
                }
                //return result;
            }
            catch (Exception ex)
            {
                fs?.Close();
                //return false;
            }

        }

        public static void TwoSheetToExecl(DataTable dt1, DataTable dt2, string file)
        {
            try
            {
                IWorkbook workbook;
                string fileExt = Path.GetExtension(file).ToLower();
                if (fileExt == ".xlsx")
                {
                    //using (FileStream streamExcel = File.OpenRead(file))
                    //{
                    //    streamExcel.Position = 0;
                    //    workbook = new XSSFWorkbook(streamExcel);
                    //}
                    workbook = new XSSFWorkbook();
                }
                else if (fileExt == ".xls")
                {
                    using (FileStream streamExcel = File.OpenRead(file))
                        workbook = new HSSFWorkbook();
                }
                else { workbook = null; }
                if (workbook == null) { return; }

                ISheet sheet1 = workbook.GetSheet("Sheet1");
                if (sheet1 == null)
                {
                    sheet1 = workbook.CreateSheet("Sheet1");
                }

                //表头
                IRow row = sheet1.CreateRow(0);
                for (int i = 0; i < dt1.Columns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(dt1.Columns[i].ColumnName);
                }

                //数据  
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    IRow row1 = sheet1.CreateRow(i + 1);
                    for (int j = 0; j < dt1.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt1.Rows[i][j].ToString());
                    }
                }

                ISheet sheet2 = workbook.GetSheet("Sheet2");
                if (sheet2 == null)
                {
                    sheet2 = workbook.CreateSheet("Sheet2");
                }

                //数据  
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    IRow row1 = sheet2.CreateRow(i);
                    for (int j = 0; j < dt2.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt2.Rows[i][j].ToString());
                    }
                }

                //转为字节数组  
                MemoryStream stream = new MemoryStream();
                workbook.Write(stream, true);
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch (Exception ex)
            {
                //throw ex;
            }
        }

        public static void VerticalTableToExcel(DataTable dt, string file)
        {
            try
            {
                IWorkbook workbook;
                string fileExt = Path.GetExtension(file).ToLower();
                if (fileExt == ".xlsx")
                {
                    using (FileStream streamExcel = File.OpenRead(file))
                        workbook = new XSSFWorkbook(streamExcel);
                }
                else if (fileExt == ".xls")
                {
                    using (FileStream streamExcel = File.OpenRead(file))
                        workbook = new HSSFWorkbook(streamExcel);
                }
                else { workbook = null; }
                if (workbook == null) { return; }

                ISheet sheet = workbook.GetSheet("Sheet2");
                if (sheet == null)
                {
                    sheet = workbook.CreateSheet("Sheet2");
                }
                //表头  
                //IRow row = sheet.CreateRow(0);
                //for (int i = 0; i < dt.Columns.Count; i++)
                //{
                //    ICell cell = row.CreateCell(i);
                //    cell.SetCellValue(dt.Columns[i].ColumnName);
                //}

                //数据  
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row1 = sheet.CreateRow(i);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }

                //转为字节数组  
                MemoryStream stream = new MemoryStream();
                workbook.Write(stream, true);
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch (Exception ex)
            {
                //throw ex;
            }

        }

        /// <summary>
        /// 获取sheet表对应的DataTable
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="strMsg"></param>
        /// <returns></returns>
        private static DataTable GetSheetDataTable(ISheet sheet, out string strMsg)
        {
            strMsg = "";
            DataTable dt = new DataTable();
            string sheetName = sheet.SheetName;
            int startIndex = 0;// sheet.FirstRowNum;
            int lastIndex = sheet.LastRowNum;

            //最大列数
            int cellCount = 0;
            IRow maxRow = sheet.GetRow(0);
            for (int i = startIndex; i <= lastIndex; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null && cellCount < row.LastCellNum)
                {
                    cellCount = row.LastCellNum;
                    maxRow = row;
                }
            }
            //列名设置
            try
            {
                //maxRow.LastCellNum = 12 // L
                for (int i = 0; i < cellCount; i++)//maxRow.FirstCellNum
                {
                    dt.Columns.Add(Convert.ToChar(((int)'A') + i).ToString());
                    //DataColumn column = new DataColumn("Column" + (i + 1).ToString());
                    //dt.Columns.Add(column);
                }
            }
            catch
            {
                strMsg = "工作表" + sheetName + "中无数据";
                return null;
            }
            //数据填充
            for (int i = startIndex; i <= lastIndex; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow drNew = dt.NewRow();
                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < row.LastCellNum; ++j)
                    {
                        if (row.GetCell(j) != null)
                        {
                            ICell cell = row.GetCell(j);
                            switch (cell.CellType)
                            {
                                case CellType.Blank:
                                    drNew[j] = "";
                                    break;
                                case CellType.Numeric:
                                    short format = cell.CellStyle.DataFormat;
                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                        drNew[j] = cell.DateCellValue;
                                    else
                                        drNew[j] = cell.NumericCellValue;
                                    if (cell.CellStyle.DataFormat == 177 || cell.CellStyle.DataFormat == 178 || cell.CellStyle.DataFormat == 188)
                                        drNew[j] = cell.NumericCellValue.ToString("#0.00");
                                    break;
                                case CellType.String:
                                    drNew[j] = cell.StringCellValue;
                                    break;
                                case CellType.Formula:
                                    try
                                    {
                                        drNew[j] = cell.NumericCellValue;
                                        if (cell.CellStyle.DataFormat == 177 || cell.CellStyle.DataFormat == 178 || cell.CellStyle.DataFormat == 188)
                                            drNew[j] = cell.NumericCellValue.ToString("#0.00");
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            drNew[j] = cell.StringCellValue;
                                        }
                                        catch { }
                                    }
                                    break;
                                default:
                                    drNew[j] = cell.StringCellValue;
                                    break;
                            }
                        }
                    }
                }
                dt.Rows.Add(drNew);
            }
            return dt;
        }

    }
}
