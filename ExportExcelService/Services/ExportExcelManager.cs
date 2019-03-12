using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using System.Web;
using Domain.Model;
using ExportExcelService.IServices;

namespace ExportExcelService.Services
{
    public class ExportExcelManager
    {
        readonly List<IExcelExportable> templates = new List<IExcelExportable>
        {
            new ExportStaffScoreService()
        };

        public ExcelPackage GetExcelPackage(string fileName, FileInfo file)
        {
            foreach(var template in templates)
            {
                if (template.IsMatch(fileName))
                {
                    return template.GetExcelPackage(file);
                }
            }

            return null;
        }
    }
}