using System;
using System.Web.Http;
using System.IO;
using System.Web;
using ExportExcelService.Services;

namespace ExportExcel.Controllers
{
    public class ExportController : ApiController
    {
        [HttpGet]
        [Route("api/export/{name}")]
        public void Download(string name)
        {
            try
            {
                var fileNameWithExtension = name + ".xlsx";
                var exportExcelManager = new ExportExcelManager();
                var path = HttpContext.Current.Server.MapPath("~/Templates/" + fileNameWithExtension);
                var file = new FileInfo(path);
                var bytes = exportExcelManager.GetExcelPackage(fileNameWithExtension, file);

                Export(fileNameWithExtension, bytes);
            }
            catch (Exception)
            {

            }
        }

        private void Export(string fileNameWithExtension, byte[] bytes)
        {
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var headerKey = "content-disposition";
            var headerValue = string.Format("attachment;  filename={0}", fileNameWithExtension);

            HttpContext.Current.Response.ContentType = contentType;
            HttpContext.Current.Response.AddHeader(headerKey, headerValue);
            HttpContext.Current.Response.BinaryWrite(bytes);
            HttpContext.Current.Response.End();
        }
    }
}