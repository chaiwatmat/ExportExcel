using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace ExportExcelService.IServices
{
    public interface IExcelExportable
    {
        string TemplateName { get; }
        bool IsMatch(string templateName);
        ExcelPackage GetExcelPackage(FileInfo file);
    }
}
