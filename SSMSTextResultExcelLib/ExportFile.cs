using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using SSMSTextResultExcelLib;

namespace SSMSTextResultExcelLib
{
    public class ExportFile
    {
        public void ExcelResult(string inputString, string path)
        {            
            var formatter = new TextFormatter();

            var ds = formatter.TotalResultFormatDS(inputString, true);

            using (var p = new ExcelPackage())
            {
                foreach (DataTable table in ds.Tables)
                {
                    var ws = p.Workbook.Worksheets.Add(table.TableName);
                    //ws.Cells[1,1].Value

                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        ws.Cells[1, col + 1].Value = table.Columns[col].ColumnName;
                        ws.Cells[1, col + 1].Style.Font.Bold = true;
                        ws.Column(col + 1).Width = 20;
                    }

                    for (int row = 0; row < table.Rows.Count; row++)
                    {
                        for (int col = 0; col < table.Columns.Count; col++)
                        {
                            int tempNumber;
                            string val = table.Rows[row][col].ToString().Trim();

                            bool success = Int32.TryParse(val, out tempNumber);

                            if (success)
                            {
                                ws.Cells[row + 2, col + 1].Value = tempNumber;
                            } else
                            {
                                ws.Cells[row + 2, col + 1].Value = val;
                            }
                            
                        }
                    }

                    //ws.Cells["A1"].Value = "This is cell A1";
                }

                p.SaveAs(new FileInfo(path));
            }
        }

        public string TextResult()
        {
            return "";
        }
    }
}
