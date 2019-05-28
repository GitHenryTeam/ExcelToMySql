using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace excel2json
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class CSVExporter
    {
		static string CheckContent(string content)
		{
			if(string.IsNullOrEmpty(content))
			{
				return content;
			}
			content = content.TrimEnd();
			content = content.Replace("\"", "\"\"");
			if(content.Contains(","))
			{
				content = "\"" + content + "\"";
			}
			return content;
		}

        static public void SaveToCSVFile(DataTable sheet, string filePath)
		{
			if (sheet.Columns.Count <= 0)
				return;
			if (sheet.Rows.Count <= 0)
				return;

			using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
			{
				using (TextWriter writer = new StreamWriter(file, new UTF8Encoding(true)))
				{
					int rowsCount = sheet.Rows.Count;
					DataColumnCollection columns = sheet.Columns;
					int columnsCount = columns.Count;
					int maxIndex = columnsCount - 1;
					for (int i = 0; i < columns.Count; i++)
					{
						string v = CheckContent(columns[i].ToString());
						writer.Write(v);
						writer.Write(i < maxIndex ? "," : "\r\n");
					}

					for (int i = 0; i < rowsCount; i++)
					{
						DataRow row = sheet.Rows[i];
						for (int j = 0; j < columnsCount; j++)
						{
							string v = CheckContent(row[columns[j]].ToString());
							writer.Write(v);
							writer.Write(j < maxIndex ? "," : "\r\n");
						}
					}
					writer.Flush();
					writer.Close();
				}
			}
		}
    }
}
