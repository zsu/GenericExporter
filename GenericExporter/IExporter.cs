using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace GenericExporter
{
    public interface IExporter
    {
        byte[] Export<T>(IEnumerable<T> rows, ExportType exportType = ExportType.Excel, string[] headers = null, bool freezeHeader = true, Func<T, object[]> formatterFunc = null);
        MemoryStream GetStream<T>(IEnumerable<T> rows, ExportType exportType = ExportType.Excel, string[] headers = null, bool freezeHeader = true, Func<T, object[]> formatterFunc = null);
        XLWorkbook ToExcel<T>(IEnumerable<T> rows, string[] headers = null, bool freezeHeader = true, Func<T, object[]> formatterFunc = null);
    }
}