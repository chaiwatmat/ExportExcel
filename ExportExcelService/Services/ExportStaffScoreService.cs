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
    public class ExportStaffScoreService : IExcelExportable
    {
        public string TemplateName => "staff_score.xlsx";

        public ExcelPackage GetExcelPackage(FileInfo file)
        {
            var data = GetDataFromService();
            
            var excelPackage = new ExcelPackage(file);
            var ws = excelPackage.Workbook.Worksheets[1];
            var rowNumber = 2;

            foreach (var d in data)
            {
                ws.Cells["A" + rowNumber].Value = d.No;
                ws.Cells["B" + rowNumber].Value = d.Name;
                ws.Cells["C" + rowNumber].Value = d.Score;

                rowNumber++;
            }

            return excelPackage;
        }

        public bool IsMatch(string templateName)
        {
            return templateName == TemplateName;
        }

        private IEnumerable<StaffScore> GetDataFromService()
        {
            var arrayData = new StaffScore[] {
                    new StaffScore { No = 1, Name = "Mr.A", Score = 10 },
                    new StaffScore { No = 2, Name = "Mr.B", Score = 20 },
                    new StaffScore { No = 3, Name = "Mr.C", Score = 15 },
                    new StaffScore { No = 4, Name = "Mr.D", Score = 30 }
            };

            return arrayData;
        }
    }
}