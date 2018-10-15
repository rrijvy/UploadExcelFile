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
using OfficeOpenXml;
using UploadExcelFile.Data;
using UploadExcelFile.Models;

namespace UploadExcelFile.Controllers
{
    public class UploadController : Controller
    {
        private readonly IHostingEnvironment _environment;
        private readonly ApplicationDbContext _context;

        public UploadController(IHostingEnvironment environment,
                               ApplicationDbContext context)
        {
            _environment = environment;
            _context = context;
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



            //fileName = this.EnsureFilename(fileName);

            using (var stream = new FileStream(GetPath(fileName), FileMode.Create))
            {
                files.CopyTo(stream);
                var ep = new ExcelPackage(new FileInfo(GetPath(fileName)));
                var wp = ep.Workbook.Worksheets["sheet1"];

                string[,] cellValue = new string[wp.Dimension.End.Row - 1, wp.Dimension.End.Column];
                List<Student> students = new List<Student>(cellValue.Length);

                for (int i = wp.Dimension.Start.Row + 1; i <= wp.Dimension.End.Row; i++)
                {
                    var student = new Student();

                    for (int j = wp.Dimension.Start.Column; j <= wp.Dimension.End.Column; j++)
                    {
                        cellValue[i - 2, j - 1] = wp.Cells[i, j].Value.ToString();
                    }
                }

                for (int i = 0; i < wp.Dimension.End.Row - 1; i++)
                {
                    for (int j = 0; j < wp.Dimension.End.Column; j = j + 3)
                    {
                        students.Add(new Student { Name = cellValue[i, j], Age = int.Parse(cellValue[i, j + 1]), Gender = cellValue[i, j + 2] });
                    }
                }

                foreach (var item in students)
                {
                    _context.Students.Add(item);
                }
                _context.SaveChanges();
                return Json(students);
            }




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