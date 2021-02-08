using System;
using System.Data;

namespace PDF
{
    public class Excel
    {
        public Spire.Xls.Workbook workbook1;
        public Spire.Xls.Worksheet sheet;
        public Excel()
        {
            workbook1 = new Spire.Xls.Workbook();       
        }
        public void Save(string fileName)
        {
            string tmpFile = System.IO.Path.GetTempFileName().Replace(".tmp", ".xlsx");
            workbook1.SaveToFile(Environment.CurrentDirectory + "\\" + fileName + "", Spire.Xls.FileFormat.Version2013);
        }
        public void Generate(string[] columns, DataTable table)
        {
       
            sheet = workbook1.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.1;
            sheet.PageSetup.BottomMargin = 0.1;


            sheet.PageSetup.Orientation = Spire.Xls.PageOrientationType.Portrait;

            int myrow = 1;
            int let = 1;
            foreach (var col in columns)
            {
                var range = GetExcelColumnName(let) + "" + myrow + ":" + GetExcelColumnName(let) + myrow;
                sheet.Range[range].Merge();
                sheet.Range[range].Style.HorizontalAlignment = Spire.Xls.HorizontalAlignType.Center;
                sheet.Range[range].Style.WrapText = true;
                sheet.Range[range].Style.Font.IsBold = true;
                sheet.Range[range].Style.VerticalAlignment = Spire.Xls.VerticalAlignType.Center;
                sheet.Range[range].BorderAround(Spire.Xls.LineStyleType.Thin);
                sheet.Range[range].Text = col;
                let++;
            }
            myrow++;
            let = 1;
            for (int i = 0; i < table.Rows.Count; i++)
            {

                for (int j = 0; j < table.Columns.Count; j++)
                {
                    var range = GetExcelColumnName(let) + "" + myrow + ":" + GetExcelColumnName(let) + myrow;
                    sheet.Range[range].Merge();
                    sheet.Range[range].Style.HorizontalAlignment = Spire.Xls.HorizontalAlignType.Center;
                    sheet.Range[range].Style.WrapText = true;
                    sheet.Range[range].Style.VerticalAlignment = Spire.Xls.VerticalAlignType.Center;
                    sheet.Range[range].BorderAround(Spire.Xls.LineStyleType.Thin);
                    sheet.Range[range].Text = table.Rows[i][j].ToString();
                    let++;
                }
                let = 1;
                myrow++;

            }
        }
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
