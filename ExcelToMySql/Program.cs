using System;
using System.Text;
using System.IO;
using System.Data;
using excel2json;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace ExcelToMySql
{
    class Program
    {

        static void Export(Options options)
        {
            string excelPath = options.ExcelPath;
            int header = options.HeaderRows;
            string fileName = Path.GetFileNameWithoutExtension(excelPath);
            String strMsg = "";
            DataSet dt = ExcelToDataSet("遍历数据.xlsx", out strMsg);
            DataSet book = dt;


            // 数据检测
            if (book.Tables.Count < 1)
            {
                throw new Exception("Excel文件中没有找到Sheet");
            }

            // 取得数据
            for (int mark = 0; mark < book.Tables.Count; mark++)
            {
                DataTable sheet = book.Tables[mark];
                if (sheet.Rows.Count <= 0)
                {
                    continue;
                    //throw new Exception("Excel Sheet中没有数据");
                }

                //-- 确定编码
                Encoding cd = new UTF8Encoding(false);
                if (options.Encoding != "utf8-nobom")
                {
                    foreach (EncodingInfo ei in Encoding.GetEncodings())
                    {
                        Encoding e = ei.GetEncoding();
                        if (e.EncodingName == options.Encoding)
                        {
                            cd = e;
                            break;
                        }
                    }
                }

                //-- 导出SQL文件
                if (options.sqlite || options.mysql)
                {
                    SQLExporter exporter = new SQLExporter(sheet, header);
                    if (string.IsNullOrEmpty(options.SQLPath))
                    {
                        options.SQLPath = string.Format("{0}/{1}", options.WorkOut, fileName);
                    }
                    options.SQLPath = sheet.TableName.ToString();
                    exporter.SaveToFile(options, cd, sheet.TableName.ToString());
                }
            }
            //}
        }

        /// <summary>
        /// Excel转换成DataTable（.xls）
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string filePath)
        {
            var dt = new DataTable();
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var hssfworkbook = new HSSFWorkbook(file);
                var sheet = hssfworkbook.GetSheetAt(0);
                for (var j = 0; j < 5; j++)
                {
                    dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
                }
                var rows = sheet.GetRowEnumerator();
                while (rows.MoveNext())
                {
                    var row = (HSSFRow)rows.Current;
                    var dr = dt.NewRow();
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);
                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            switch (cell.CellType)
                            {
                                case CellType.Blank:
                                    dr[i] = "[null]";
                                    break;
                                case CellType.Boolean:
                                    dr[i] = cell.BooleanCellValue;
                                    break;
                                case CellType.Numeric:
                                    dr[i] = cell.ToString();
                                    break;
                                case CellType.String:
                                    dr[i] = cell.StringCellValue;
                                    break;
                                case CellType.Error:
                                    dr[i] = cell.ErrorCellValue;
                                    break;
                                case CellType.Formula:
                                    try
                                    {
                                        dr[i] = cell.NumericCellValue;
                                    }
                                    catch
                                    {
                                        dr[i] = cell.StringCellValue;
                                    }
                                    break;
                                default:
                                    dr[i] = "=" + cell.CellFormula;
                                    break;
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        /// <summary>
        /// Excel转换成DataSet（.xlsx/.xls）
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="strMsg"></param>
        /// <returns></returns>
        public static DataSet ExcelToDataSet(string filePath, out string strMsg)
        {
            strMsg = "";
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string fileType = Path.GetExtension(filePath).ToLower();
            string fileName = Path.GetFileName(filePath).ToLower();
            try
            {
                ISheet sheet = null;
                int sheetNumber = 0;
                // 加载Excel文件
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (fileType == ".xlsx")
                    {
                        // 2007版本
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        sheetNumber = workbook.NumberOfSheets;
                        for (int i = 0; i < sheetNumber; i++)
                        {
                            string sheetName = workbook.GetSheetName(i);
                            sheet = workbook.GetSheet(sheetName);
                            if (sheet != null)
                            {
                                dt = GetSheetDataTable(sheet, out strMsg);
                                if (dt != null)
                                {
                                    dt.TableName = sheetName.Trim();
                                    ds.Tables.Add(dt);
                                }
                                else
                                {
                                    //MessageBox.Show("Sheet数据获取失败，原因：" + strMsg);
                                }
                            }
                        }
                    }
                    else if (fileType == ".xls")
                    {
                        // 2003版本
                        HSSFWorkbook workbook = new HSSFWorkbook(fs);
                        sheetNumber = workbook.NumberOfSheets;
                        for (int i = 0; i < sheetNumber; i++)
                        {
                            string sheetName = workbook.GetSheetName(i);
                            sheet = workbook.GetSheet(sheetName);
                            if (sheet != null)
                            {
                                dt = GetSheetDataTable(sheet, out strMsg);
                                if (dt != null)
                                {
                                    dt.TableName = sheetName.Trim();
                                    ds.Tables.Add(dt);
                                }
                                else
                                {
                                    //MessageBox.Show("Sheet数据获取失败，原因：" + strMsg);
                                }
                            }
                        }
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {
                strMsg = ex.Message;
                return null;
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
            int lastIndex = sheet.PhysicalNumberOfRows;
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
                for (int i = 0; i < maxRow.LastCellNum; i++)//maxRow.FirstCellNum
                {
                    String temp = maxRow.GetCell(i).StringCellValue;
                    dt.Columns.Add(temp);
                }
            }
            catch
            {
                strMsg = "工作表" + sheetName + "中无数据";
                return null;
            }

            //数据填充
            for (int i = 1; i <= lastIndex; i++)
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


        static void Main(string[] args)
        {
            if (args.Length <= 0)
            {
                Console.WriteLine("传参个数不正确,请检测数据传参！");
                return;
            }

            String busType = args[0].ToString();
            if (busType.Length <= 0)
            {
                Console.WriteLine("传参值正确,请检测数据传参！");
                return;
            }


            Options option = new Options();
            option.HeaderRows = 1;
            option.WorkOut = "ExportMySql";
            option.mysql = true;
            option.SQLPath = "test_type";
            String filePath = @"遍历数据.xlsx";
            option.ExcelPath = filePath;

            if (busType.Contains("NeManager"))
                option.DbName = "unm2000nemanager";

            if (busType.Contains("TP"))
                option.DbName = "unm2000autotp";

            if (busType.Contains("VC"))
                option.DbName = "unm2000autovc";

            if (busType.Contains("OTN"))
                option.DbName = "unm2000autootn";

            if (busType.Contains("PTN"))
                option.DbName = "unm2000autoptn";

            if (Directory.Exists(option.WorkOut) == false)
                Directory.CreateDirectory(option.WorkOut);

            if (File.Exists(filePath))
                Export(option);
            else
                Console.WriteLine("Excel文件不存在！");
        }
    }
}
