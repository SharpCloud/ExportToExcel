using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExportToExcel
{
    public class ExcelWriter
    {
        private readonly _Application _excel;
        private readonly _Workbook _workbook;

        public ExcelWriter()
        {
            _excel = new Application
            {
                SheetsInNewWorkbook = 1,
                Visible = true
            };

            _workbook = _excel.Workbooks.Add(Missing.Value);
        }

        public void AddSheet(string sheetName, string[,] data)
        {
            SetUserControl(false);

            var worksheet = _workbook.Sheets.Add();
            worksheet.Name = sheetName;

            var columnCount = data.GetLength(0);
            var rowCount = data.GetLength(1);

            for (var c = 0; c < columnCount; c++)
            {
                for (var r = 0; r < rowCount; r++)
                {
                    worksheet.Cells[c + 1, r + 1] = data[c, r];
                }
            }

            SetUserControl(true);
        }

        private void SetUserControl(bool allowUserControl)
        {
            _excel.ScreenUpdating = allowUserControl;
            _excel.UserControl = allowUserControl;
        }
    }
}
