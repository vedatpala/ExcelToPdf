using System.ComponentModel.DataAnnotations;

namespace WebApplicationExcelToPdfConvert.Models
{
    public class ExcelFileModel
    {
        [Required(ErrorMessage = "Lütfen bir Excel dosyası seçin.")]
        [Display(Name = "Excel Dosyası")]
        public IFormFile ExcelFile { get; set; }
    }
}
