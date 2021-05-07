using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace GenericExporter
{
    public class Exporter : IExporter
    {
        public byte[] Export<T>(
        List<T> rows,
        ExportType exportType = ExportType.Excel,
        string[] headers = null,
        Func<T, object[]> formatterFunc = null)
        {
            using (var ms = GetStream(rows, exportType, headers, formatterFunc))
            {
                return ms.ToArray();
            }
        }
        public MemoryStream GetStream<T>(
        List<T> rows,
        ExportType exportType = ExportType.Excel,
        string[] headers = null,
        Func<T, object[]> formatterFunc = null)
        {
            var ms = new MemoryStream();
            switch (exportType)
            {
                case ExportType.Excel:
                    using (var excel = ToExcel<T>(rows, headers, formatterFunc))
                        excel.SaveAs(ms);
                    break;
            }
            ms.Position = 0;
            return ms;
        }
        public XLWorkbook ToExcel<T>(
            List<T> rows,
            string[] headers = null,
            Func<T, object[]> formatterFunc = null)
        {
            var workbook = new XLWorkbook();
            var row = 1;
            var col = 1;
            var worksheet = workbook.Worksheets.Add("Export");
            var pis = typeof(T).GetProperties();
            foreach (var h in headers ?? pis.Select(pi => pi.Name).ToArray())
            {
                worksheet.Cell(row, col++).Value = h;
            }
            foreach (var o in rows)
            {
                row++;
                col = 1;
                var values = formatterFunc == null ?
                    pis.Select(pi => pi.GetValue(o)) : formatterFunc(o);
                foreach (var v in values)
                {
                    worksheet.Cell(row, col++).Value = v;
                }
            }
            worksheet.Columns().AdjustToContents();
            return workbook;
        }
    }
    public enum ExportType
    {
        Excel
    }
}
