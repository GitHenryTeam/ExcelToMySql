using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace excel2json
{
    class SQLExporter
    {
		enum RowType
		{
			INTEGER = 0,
			text,
		}
		RowType[] _types = null;
        DataTable m_sheet;
        int m_headerRows;

        /// <summary>
        /// 初始化内部数据
        /// </summary>
        /// <param name="sheet">Excel读取的一个表单</param>
        /// <param name="headerRows">表头有几行</param>
        public SQLExporter(DataTable sheet, int headerRows)
        {
            m_sheet = sheet;
            m_headerRows = headerRows;
			DataColumnCollection columns = sheet.Columns;
			if (columns != null && columns.Count > 0)
			{
				_types = new RowType[columns.Count];
				int firstDataRow = m_headerRows - 1;
				for (int i = firstDataRow; i < sheet.Rows.Count; i++)
				{
					DataRow row = sheet.Rows[i];
					for (int j = 0; j < columns.Count; j ++ )
					{
						if (_types[j] == RowType.text)
						{
							continue;
						}

						DataColumn column = columns[j];
						object value = row[column];
						Type type = value.GetType();
						if (type == typeof(double))
						{
							double num = (double)value;
							if ((int)num != num)
							{
								_types[j] = RowType.text;
							}
						}
						else if (type == typeof(string))
						{
							int var = 0;
							string str = value.ToString();
							if (string.IsNullOrEmpty(str))
							{
								continue;
							}
							else if (int.TryParse((string)value, out var) == false || var.ToString() != (string)value)
							{
								_types[j] = RowType.text;
							}
						}
					}
				}
			}
        }

        /// <summary>
        /// 转换成SQL字符串，并保存到指定的文件
        /// </summary>
        /// <param name="filePath">存盘文件</param>
        /// <param name="encoding">编码格式</param>
		public void SaveToFile(Options options, Encoding encoding, string tableName)
        {
            //-- 转换成SQL语句
			string mysql_header = null;
			string sqlite_header = null;
            GetTabelStructSQL(m_sheet, options.DbName,tableName, out sqlite_header, out mysql_header);
            string tabelContent = GetTableContentSQL(m_sheet, tableName);

            if(!Directory.Exists(options.WorkOut))
                Directory.CreateDirectory(options.WorkOut);

			bool[] build = { options.sqlite, options.mysql};
			string[] headers = { sqlite_header, mysql_header};
            string[] filePaths = { options.WorkOut + "\\" + options.SQLPath + "_sqlite.sql", options.WorkOut + "\\" + options.SQLPath + ".sql" };

			for (int i = 0; i < 2; i ++ )
			{
				//-- 保存文件
				if(build[i])
				{
					string filePath = filePaths[i];
                    if (File.Exists(filePath))
                        File.Delete(filePath);

					using (FileStream file = new FileStream(filePath, FileMode.Append, FileAccess.Write))
					{
						using (TextWriter writer = new StreamWriter(file, encoding))
						{
							writer.Write(headers[i]);
							writer.WriteLine();
							writer.Write(tabelContent);
							writer.WriteLine();
						}
					}
				}
			}
        }

        /// <summary>
        /// 将表单内容转换成INSERT语句
        /// </summary>
        private string GetTableContentSQL(DataTable sheet, string tabelName)
        {
            StringBuilder sbContent = new StringBuilder();
            StringBuilder sbNames = new StringBuilder();
            StringBuilder sbValues = new StringBuilder();

            //-- 字段名称列表
            foreach (DataColumn column in sheet.Columns)
            {
                if (column.ToString().Length>1)
                { 
                    sbNames.Append(column.ToString());
                    sbNames.Append(", ");
                }
            }

            //-- 逐行转换数据
            int firstDataRow = m_headerRows - 1;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++ )
            {
                DataRow row = sheet.Rows[i];
				sbValues.Remove(0, sbValues.Length);

				int j = 0;
				DataColumnCollection columns = sheet.Columns;
				for (j = 0; j < columns.Count; j ++ )
				{
					DataColumn column = columns[j];
                    object value = row[column];

                    if (String.IsNullOrEmpty(value.ToString()) && value.ToString().Length>1 && (j == 0))//处理第一列填并且若为空 直接跳过 后面列值为空或不填的情况
                    {
                        continue;
                    }
                    
					if (sbValues.Length > 0)
					{
						sbValues.Append(", ");
					}
					
					Type type = value.GetType();
					if (type == typeof(System.DBNull))
					{
						if(j == 0)
						{
							break;
						}
                        sbValues.Append(@"'NULL'");
					}
					else
					{
						sbValues.AppendFormat("'{0}'", value.ToString());
					}
				}

				if(j > 0)
				{
                    if (sbValues.Length > 0)
                    {
                        sbValues.Append(",'NULL','NULL'");
                        sbContent.AppendFormat("INSERT INTO `{0}` VALUES({1});\n", tabelName, sbValues.ToString());
                    }
				}
            }
            sbContent = sbContent.Replace(@"\",@"\\");
            return sbContent.ToString();
        }

        /// <summary>
        /// 根据表头构造CREATE TABLE语句
        /// </summary>
        private void GetTabelStructSQL(DataTable sheet, string dbName, string tabelName, out string sqlite_header, out string mysql_header)
        {
			sqlite_header = "";
			mysql_header = "";

            // sqlite 删除表并重新创建
            StringBuilder sqlite = new StringBuilder();
            sqlite.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", dbName);
            sqlite.AppendFormat("CREATE TABLE `{0}` (\n", dbName);

            //mysql 创建库并使用状态
			StringBuilder mysql = new StringBuilder();
            mysql.AppendFormat("CREATE DATABASE IF NOT EXISTS `{0}`;\n", dbName);
            mysql.AppendFormat("\r\n");
            mysql.AppendFormat("USE `{0}`;\n", dbName);
            mysql.AppendFormat("\r\n");
            mysql.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", tabelName);
            mysql.AppendFormat("\r\n");  
			mysql.AppendFormat("CREATE TABLE IF NOT EXISTS `{0}` (\n", tabelName);

            //遍历列名
			DataColumn column = null;
			DataColumnCollection columns = sheet.Columns;
			for (int i = 0; i < columns.Count; i ++ )
			{
				column = columns[i];
				string filedName = column.ToString();

                if (filedName.Contains("Column")!=true)
                {
                    sqlite.AppendFormat("`{0}` TEXT,\n", filedName);
                    mysql.AppendFormat("`{0}` TEXT,\n", filedName);
                }
			}
            //补充Tag Result 列
            mysql.AppendFormat("`{0}` TEXT,\n", "Tag");
            mysql.AppendFormat("`{0}` TEXT \n", "Result");
            mysql.AppendFormat("\r\n");

            sqlite.AppendLine("\n);");
			mysql.AppendFormat("\n)  ENGINE=InnoDB DEFAULT CHARSET=utf8;\r\n");

			sqlite_header = sqlite.ToString();
			mysql_header = mysql.ToString();
        }
    }
}
