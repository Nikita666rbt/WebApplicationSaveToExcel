using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApplication4.Models;

namespace WebApplication4.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View(new Organization());
        }

        [HttpPost]
        public ActionResult Save(Organization organization)
        {
            if (ModelState.IsValid)
            {
                string filePath = Server.MapPath("~/App_Data/Организации.xlsx");
                SaveToExcel(filePath, organization);
                TempData["Message"] = "Данные успешно сохранены!";
                return RedirectToAction("Index");
            }
            return View("Index", organization);
        }

        public void SaveToExcel(string filePath, Organization organization) 
        {
            FileInfo fileInfo = new FileInfo(filePath);
            string tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileName(filePath));//временный файл чтобы избежать конфликта по незакрытию
            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet;
                if (fileInfo.Exists)
                {
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        package.Load(stream);
                        worksheet = package.Workbook.Worksheets[0];
                    }
                }
                else//если у нас пустой файл
                {
                    worksheet = package.Workbook.Worksheets.Add("Организации");
                    worksheet.Cells[1, 1].Value = "Наименование организации";
                    worksheet.Cells[1, 2].Value = "Юридический адрес";
                    worksheet.Cells[1, 3].Value = "Телефон";
                    worksheet.Cells[1, 4].Value = "Email";
                }

                var rowCount = worksheet.Dimension?.Rows ?? 0;
                worksheet.Cells[rowCount + 1, 1].Value = organization.Name;
                worksheet.Cells[rowCount + 1, 2].Value = organization.LegalAddress;
                worksheet.Cells[rowCount + 1, 3].Value = organization.Phone;
                worksheet.Cells[rowCount + 1, 4].Value = organization.Email;

                package.SaveAs(new FileInfo(tempFilePath));
            }
            System.IO.File.Delete(filePath); 
            System.IO.File.Move(tempFilePath, filePath);
        }
    }
}