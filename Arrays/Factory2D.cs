using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Excel.Helpers.Arrays
{
    public class Factory2D
    {
        private DataTable Source;

        public bool HasHeader { get; set; }

        public Factory2D(DataTable source)
        {
            Source = source;
            HasHeader = true;
        }

        public object[,] Header()
        {
            object[,] array2D = new object[1, Source.Columns.Count];
            // ヘッダ名を設定
            for (var i = 0; i < Source.Columns.Count; ++i)
            {
                DataColumn item = Source.Columns[i];
                array2D[0, i] = item.Caption;
            }
            return array2D;
        }

        public object[,] Payload()
        {
            object[,] array2D = new object[Source.Rows.Count, Source.Columns.Count];
            // データを設定
            for (int i_r = 0; i_r < Source.Rows.Count; i_r++)
            {
                DataRow item = Source.Rows[i_r];
                for (int i_c = 0; i_c < Source.Columns.Count; i_c++)
                {
                    array2D[i_r, i_c] = item.ItemArray[i_c];
                }
            }
            return array2D;
        }

        public object[,] Generate()
        {
            object[,] array2D = new object[Source.Rows.Count + (HasHeader ? 1 : 0), Source.Columns.Count];
#if false
            var row = 0;
            if (HasHeader)
            {
                // ヘッダ名を設定
                for (var i = 0; i < Source.Columns.Count; ++i)
                {
                    DataColumn item = Source.Columns[i];
                    array2D[0, i] = item.Caption;
                }
                ++row;
            }
            // データを設定
            for (int i_r = 0; i_r < Source.Rows.Count; i_r++)
            {
                DataRow item = Source.Rows[i_r];
                for (int i_c = 0; i_c < Source.Columns.Count; i_c++)
                {
                    array2D[row, i_c] = item.ItemArray[i_c];
                }
                ++row;
            }
#else
            var i = 0;
            if (HasHeader)
            {
                var h = Header();
                Array.Copy(h, 0, array2D, 0, h.Length);
                i = h.Length;
            }
            var d = Payload();
            Array.Copy(d, 0, array2D, i, d.Length);
#endif
            return array2D;
        }

        public static object[,] Convert(DataTable table, bool hasHeader = true)
        {
            return new Factory2D(table)
            {
                HasHeader = hasHeader
            }
            .Generate();
        }
    }
}
