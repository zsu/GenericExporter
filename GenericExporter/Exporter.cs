using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;

namespace GenericExporter
{
    public class Exporter : IExporter
    {
        public byte[] Export<T>(
        IEnumerable<T> rows,
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
        IEnumerable<T> rows,
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
            IEnumerable<T> rows,
            string[] headers = null,
            Func<T, object[]> formatterFunc = null)
        {
            bool isDynamic=false;
            var workbook = new XLWorkbook();
            var row = 1;
            var col = 1;
            var worksheet = workbook.Worksheets.Add("Export");
            if(rows==null || rows.Count()==0)
            {
                return workbook;
            }
            var properties = typeof(T).GetProperties();
            string[] names=properties.Select(x=>x.Name).ToArray();
            if(names?.Length==0)
            {
                isDynamic = true;
                var dynamicProperties = GetPropertiesForDynamic(rows.FirstOrDefault());
                names = dynamicProperties.Select(x => x.Key).ToArray();
            }
            foreach (var h in headers ?? names)
            {
                worksheet.Cell(row, col++).Value = h;
            }
            foreach (var item in rows)
            {
                row++;
                col = 1;
                var values = formatterFunc == null ?
                    (isDynamic? GetPropertiesForDynamic(rows.FirstOrDefault()).Select(x=>x.Value): properties.Select(x => x.GetValue(item))) : formatterFunc(item);
                foreach (var v in values)
                {
                    worksheet.Cell(row, col++).Value = v;
                }
            }
            worksheet.Columns().AdjustToContents();
            return workbook;
        }
        public Dictionary<string, object> GetPropertiesForDynamic(dynamic dynamicToGetPropertiesFor)
        {
            JObject attributesAsJObject = (JObject)JToken.FromObject(dynamicToGetPropertiesFor);
            Dictionary<string, object> properties = attributesAsJObject.ToObject<Dictionary<string, object>>();
            return properties;
        }
    }
    public enum ExportType
    {
        Excel
    }
}
