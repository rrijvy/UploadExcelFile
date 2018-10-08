using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace UploadExcelFile.Controllers
{
    public class UploadController : Controller
    {
        private readonly IHostingEnvironment _environment;

        public UploadController(IHostingEnvironment environment)
        {
            _environment = environment;
        }

        public IActionResult Index(byte[] file)
        {
            return View();
        }

        //[HttpPost]
        //public IActionResult Index(IFormFile file)
        //{
        //    string folderName = "Upload";
        //    string webRootPath = _environment.WebRootPath;
        //    string newPath = Path.Combine(webRootPath, folderName);
        //    StringBuilder sb = new StringBuilder();
        //    if (!Directory.Exists(newPath))
        //        Directory.CreateDirectory(newPath);
        //    if (file.Length > 0)
        //    {
        //        string sFileExtension = Path.GetExtension(file.FileName).ToLower();
        //        ISheet sheet;
        //        string fullPath = Path.Combine(newPath, file.FileName);
        //        using (var stream = new FileStream(fullPath, FileMode.Create))
        //        {
        //            file.CopyTo(stream);
        //            stream.Position = 0;
        //            if (sFileExtension == ".xls")
        //            {
        //                HSSFWorkbook hssfwb = new HSSFWorkbook(stream);
        //                sheet = hssfwb.GetSheetAt(0);
        //            }
        //            else
        //            {
        //                XSSFWorkbook xssfwb = new XSSFWorkbook(stream);
        //                sheet = xssfwb.GetSheetAt(0);
        //            }
        //            IRow headerRow = sheet.GetRow(0);
        //            int cellCount = headerRow.LastCellNum;
                    
        //        }
        //    }
        //    return View();
        //}


        //[HttpPost]
        //public IActionResult Index(IList<IFormFile> files)
        //{
        //    foreach (IFormFile item in files)
        //    {
        //        string fileName = ContentDispositionHeaderValue.Parse(item.ContentDisposition).FileName.Trim('"');
        //        fileName = this.EnsureFilename(fileName);
        //        using(FileStream fileStream = System.IO.File.Create(this.GetPath(fileName)))
        //        {
        //        }
        //    }

        //    return this.Content("Successfull.");
        //}

        //private string GetPath(string fileName)
        //{
        //    string path = _environment.WebRootPath + "\\Upload\\";
        //    if (!Directory.Exists(path))
        //    {
        //        Directory.CreateDirectory(path);
        //    }
        //    return path + fileName;
        //}

        //private string EnsureFilename(string fileName)
        //{
        //    if (fileName.Contains("\\"))
        //        fileName = fileName.Substring(fileName.LastIndexOf("\\") + i);
        //    return fileName;
        //}
    }
}