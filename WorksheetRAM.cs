using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Excel.Helpers
{
    public class WorksheetRAM
    {
        private Microsoft.Office.Interop.Excel.Application Application;

        public WorksheetRAM(Microsoft.Office.Interop.Excel.Application excel)
        {
            Application = excel;
        }

        public Microsoft.Office.Interop.Excel.Worksheet GetEndSheet()
        {
            return Application.Worksheets.get_Item(Application.Worksheets.Count);
        }

        public void WriteOnNewSheet(DataTable table)
        {
            Microsoft.Office.Interop.Excel.Worksheet endSheet = GetEndSheet();
            Microsoft.Office.Interop.Excel.Worksheet newSheet = Application.Sheets.Add(Type.Missing, endSheet, Type.Missing, Type.Missing);
            WriteOn(table, newSheet, newSheet.Cells[1, 1], true);
        }

        //
        // rngFrom 始点セル
        //
        public void WriteOn(DataTable table, Microsoft.Office.Interop.Excel.Worksheet newSheet, Microsoft.Office.Interop.Excel.Range rngFrom, bool hasHeader)
        {
            Debug.Assert(table.Rows.Count > 0);

            // 対象シートのセル
            Microsoft.Office.Interop.Excel.Range xlCells = newSheet.Cells;
            // 終点セル
            Microsoft.Office.Interop.Excel.Range rngTo = xlCells[(rngFrom.Row - 1) + table.Rows.Count + (hasHeader ? 1 : 0), (rngFrom.Column - 1) + table.Columns.Count];
            // 貼り付け範囲作成
            Microsoft.Office.Interop.Excel.Range rngTarget = newSheet.get_Range(rngFrom, rngTo);
            // 配列を張り付け
            rngTarget.NumberFormatLocal = "@";
            rngTarget.Value = Arrays.Factory2D.Convert(table, hasHeader);
        }

        private static readonly string FN_SHEET_NAME = "sheet_name";
        public DataTable CreateSheetNameTbl()
        {
            var table = new DataTable();
            table.Columns.AddRange(new DataColumn[] {
                new DataColumn(FN_SHEET_NAME, typeof(string)){ Caption = Application.ActiveWorkbook.Name + " シート名" }, 
            });
            table.PrimaryKey = new DataColumn[] { table.Columns[FN_SHEET_NAME] };
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in Application.Worksheets)
            {
                var row = table.NewRow();
                row[FN_SHEET_NAME] = sheet.Name;
                table.Rows.Add(row);
            }
            return table;
        }
    }
}
