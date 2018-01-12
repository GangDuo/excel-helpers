using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;

namespace Excel.Helpers.UnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var table = new DataTable();
            table.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("id", Type.GetType("System.Int32")),
                new DataColumn("name", Type.GetType("System.String"))
            });
            for (int i = 0; i <= 2; i++)
            {
                var row = table.NewRow();
                row["id"] = i;
                row["name"] = "name " + i;
                table.Rows.Add(row);
            }
            var xs = Arrays.Factory2D.Convert(table);
            //System.Diagnostics.Debug.WriteLine(xs);
        }
    }
}
