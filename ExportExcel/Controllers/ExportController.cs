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

using Domain.Model;
using ExportExcelService.Services;

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
        [Route("api/export/download")]
        public void Download(string name)
        {
            var fileNameWithExtension = name + ".xlsx";
            var exportExcelManager = new ExportExcelManager();
            var path = HttpContext.Current.Server.MapPath("~/Templates/" + fileNameWithExtension);
            var file = new FileInfo(path);

            var exportExcelPackage = exportExcelManager.GetExcelPackage(fileNameWithExtension, file);

            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var headerKey = "content-disposition";
            var headerValue = string.Format("attachment;  filename={0}", fileNameWithExtension);

            HttpContext.Current.Response.ContentType = contentType;
            HttpContext.Current.Response.AddHeader(headerKey, headerValue);
            HttpContext.Current.Response.BinaryWrite(exportExcelPackage.GetAsByteArray());
            HttpContext.Current.Response.End();
        }
    }
}