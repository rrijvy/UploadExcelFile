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
using UploadExcelFile.Models;

namespace UploadExcelFile.Controllers
{
    public class UploadController : Controller
    {
        private readonly IHostingEnvironment _environment;

        public UploadController(IHostingEnvironment environment)
        {
            _environment = environment;
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


        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(IFormFile files)
        {

            List<Student> list = new List<Student>();


            string fileName = ContentDispositionHeaderValue.Parse(files.ContentDisposition).FileName.Trim('"');
            string sFileExtension = Path.GetExtension(files.FileName).ToLower();
            ISheet sheet;

            //fileName = this.EnsureFilename(fileName);

            using (var stream = new FileStream(GetPath(fileName), FileMode.Create))
            {
                files.CopyTo(stream);
                if (sFileExtension == ".xls")
                {
                    HSSFWorkbook hssfwb = new HSSFWorkbook(stream);
                    sheet = hssfwb.GetSheetAt(0);
                }
                else
                {
                    XSSFWorkbook xssfwb = new XSSFWorkbook(stream);
                    sheet = xssfwb.GetSheetAt(0);
                }
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        list.Add()
                    }
                }

            }



            return this.Content("Successfull.");
        }

        private string GetPath(string fileName)
        {
            string path = _environment.WebRootPath + "\\Upload\\";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path + fileName;
        }

        //private string EnsureFilename(string fileName)
        //{
        //    if (fileName.Contains("\\"))
        //        fileName = fileName.Substring(fileName.LastIndexOf("\\") + 1);

        //    return fileName;
        //}
    }
}