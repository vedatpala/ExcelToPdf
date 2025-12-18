using Microsoft.AspNetCore.Mvc;
using Spire.Xls;
using System.Diagnostics;
using WebApplicationExcelToPdfConvert.Models;

namespace WebApplicationExcelToPdfConvert.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;
        private readonly ILogger<HomeController> _logger;

        public HomeController(IWebHostEnvironment hostingEnvironment, ILogger<HomeController> logger)
        {
            _hostingEnvironment = hostingEnvironment;
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ConvertToPdf(ExcelFileModel model)
        {
            if (!ModelState.IsValid)
            {
                return Json(new { success = false, message = "Geçersiz dosya formatý." });
            }

            string filePath = null;
            string pdfPath = null;
            string pdfFileName = null;

            try
            {
                var uploadsFolder = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }

                var uniqueFileName = $"{Guid.NewGuid()}_{Path.GetFileNameWithoutExtension(model.ExcelFile.FileName)}";
                filePath = Path.Combine(uploadsFolder, $"{uniqueFileName}{Path.GetExtension(model.ExcelFile.FileName)}");
                pdfFileName = $"{uniqueFileName}.pdf";
                pdfPath = Path.Combine(uploadsFolder, pdfFileName);

                // Save the uploaded file
                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    await model.ExcelFile.CopyToAsync(fileStream);
                }

                // Convert to PDF using Spire.XLS
                try
                {
                    using (Workbook workbook = new Workbook())
                    {
                        workbook.LoadFromFile(filePath);
                        workbook.SaveToFile(pdfPath, FileFormat.PDF);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "PDF dönüþtürme hatasý");
                    return Json(new { success = false, message = "Excel dosyasý PDF'e dönüþtürülürken bir hata oluþtu." });
                }

                // Verify the PDF was created
                if (!System.IO.File.Exists(pdfPath))
                {
                    return Json(new { success = false, message = "PDF oluþturulamadý." });
                }

                // Return success with file information
                var pdfUrl = $"/uploads/{pdfFileName}";
                return Json(new
                {
                    success = true,
                    pdfUrl = pdfUrl,
                    fileName = pdfFileName
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya iþleme hatasý");
                return Json(new { success = false, message = "Bir hata oluþtu: " + ex.Message });
            }
            finally
            {
                // Clean up the original Excel file if it exists
                try
                {
                    if (filePath != null && System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Dosya temizleme hatasý");
                    // Don't fail the request if cleanup fails
                }
            }
        }
    }
}
