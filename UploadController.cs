using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Aspose.Pdf;
using System.Net;

namespace mypdfdoc.Controllers
{
    public class UploadController : Controller
    {
        // GET: Upload
        public ActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public ActionResult UploadFile()
        {
            return View();
        }
        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                
                if (file.ContentLength > 0)
                {
                    Document pdfDocument = new Document(file.InputStream);
                    DocSaveOptions saveOptions = new DocSaveOptions();
                    // Specify the output format as DOCX
                    saveOptions.Format = DocSaveOptions.DocFormat.DocX;
                    saveOptions.Mode = DocSaveOptions.RecognitionMode.Flow;

                    // Set the Horizontal proximity as 2.5
                    saveOptions.RelativeHorizontalProximity = 2.5f;
                    // Save document in docx format
                    //MemoryStream stream = new MemoryStream();
                    pdfDocument.Save(stream, saveOptions);
                    /*
                    string _FileName = Path.GetFileName(file.FileName);
                    string _path = Path.Combine(Server.MapPath("~/UploadedFiles"), _FileName);
                    file.SaveAs(_path);
                    */
                    

                }
                ViewBag.Message = "File Uploaded Successfully!!";
                //return stream;
               // FileStream file1 = new FileStream(@"C:\Users\dasdib\Desktop\b.docx", FileMode.Create, FileAccess.Write);
                //stream.WriteTo(file1);
                //file1.Close();
                //stream.Close();
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", file.FileName.Replace("pdf","docx"));
                //return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
            catch(Exception e)
            {
                ViewBag.Message = "File upload failed!!";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
        }

        [HttpPost]
        public ActionResult FromURLFile()
        {
            MemoryStream stream;
            MemoryStream stream2 = new MemoryStream();
            try
            {
                string number1 = Request.Form["txtNumber1"];
                using (WebClient webClient = new WebClient())
                {
                    byte[] dataBytes = webClient.DownloadData(number1);
                    stream = new MemoryStream(dataBytes);
                }
                Document pdfDocument = new Document(stream);
                DocSaveOptions saveOptions = new DocSaveOptions();
                saveOptions.Format = DocSaveOptions.DocFormat.DocX;
                saveOptions.Mode = DocSaveOptions.RecognitionMode.Flow;
                saveOptions.RelativeHorizontalProximity = 2.5f;
                pdfDocument.Save(stream2, saveOptions);
                ViewBag.Message = "File Uploaded Successfully!!";
                var a = Guid.NewGuid();
                return File(stream2.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document","fromurl"+a+".docx");
            }
            catch (Exception e)
            {
                ViewBag.Message = "File upload failed!!";
                return File(stream2.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
        }
        //[HttpGet]
        //public ActionResult FromURLFile()
        //{
        //    return View("UploadFile");
        //}
    }
}