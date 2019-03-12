using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;


using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;

using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExportExcel.Controllers
{
    public class ExportController : ApiController
    {
        // GET api/<controller>
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        [HttpGet]
        [Route("api/Export/download")]
        public void Download()
        {
            var data = GetDataFromService();
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Data");
                var rowNumber = 2;

                foreach(var d in data)
                {
                    ws.Cells["A" + rowNumber].Value = d.No;
                    ws.Cells["B" + rowNumber].Value = d.Name;
                    ws.Cells["C" + rowNumber].Value = d.Score;

                    rowNumber++;
                }


                var fileName = "ExcellData.xlsx";
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var headerKey = "content-disposition";
                var headerValue = string.Format("attachment;  filename={0}", fileName);

                HttpContext.Current.Response.ContentType = contentType;
                HttpContext.Current.Response.AddHeader(headerKey, headerValue);
                HttpContext.Current.Response.BinaryWrite(p.GetAsByteArray());
                HttpContext.Current.Response.End();
            }
        }

        private IEnumerable<StaffScore> GetDataFromService()
        {
            var arrayData = new StaffScore[] {
                    new StaffScore { No = 1, Name = "Mr.A", Score = 10 },
                    new StaffScore { No = 2, Name = "Mr.B", Score = 20 },
                    new StaffScore { No = 3, Name = "Mr.C", Score = 15 }
            };

            return arrayData;
        }
    }

    public class StaffScore
    {
        public int No { get; set; }
        public string Name { get; set; }
        public int Score { get; set; }
    }
}