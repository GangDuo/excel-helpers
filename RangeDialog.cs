using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Excel.Helpers
{
    public class RangeDialog
    {
        private Microsoft.Office.Interop.Excel.Application Application;

        public string Prompt { get; set; }
        public bool IsCanceled { get; private set; }
        public Microsoft.Office.Interop.Excel.Range Data { get; private set; }

        public RangeDialog(Microsoft.Office.Interop.Excel.Application excel)
        {
            Application = excel;
            Prompt = "JANが入力されたセルを選択してください。\nCtrlキーで複数セルを選択できます。";
        }

        public void Show()
        {
            var ret = Application.InputBox(Prompt,
                "範囲選択",
                Application.ActiveCell.get_Address(Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1),
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            IsCanceled = ret is bool;
            if (IsCanceled)
            {
                // ユーザがキャンセルボタンを選択した
                return;
            }
            Debug.Assert(ret is Microsoft.Office.Interop.Excel.Range);
            Data = ret;
        }
    }
}
