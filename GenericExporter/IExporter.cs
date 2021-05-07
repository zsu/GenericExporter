using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace GenericExporter
{
    public interface IExporter
    {
        byte[] Export<T>(List<T> rows, ExportType exportType = ExportType.Excel, string[] headers = null, Func<T, object[]> formatterFunc = null);
        MemoryStream GetStream<T>(List<T> rows, ExportType exportType = ExportType.Excel, string[] headers = null, Func<T, object[]> formatterFunc = null);
        XLWorkbook ToExcel<T>(List<T> rows, string[] headers = null, Func<T, object[]> formatterFunc = null);
    }
}