using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PDFConvertor.Services;
using Spire.Doc;
using Spire.Xls;


namespace PDFConvertor.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : ControllerBase
    {
        private readonly ILogger<ExportController> _logger;
     
        public ExportController(ILogger<ExportController> logger)
        {
            _logger = logger;            
        }

        
        [HttpPost]
        [Route("WordToPdf")]
        public async Task<IActionResult> WordToPdf(IFormFile _file)
        {
            DocumentConvertor doc = new DocumentConvertor();
            FileDocument _Document = new FileDocument();
            string ContentType, attachmentName = "";
            byte[] fileBytes = null;
        
            ContentType = _file.ContentType;
            attachmentName = _file.FileName;
            using (var memoryStream = new MemoryStream())
            {
                await _file.CopyToAsync(memoryStream);
                fileBytes = memoryStream.ToArray();
            }           
            _Document.format = Spire.Doc.FileFormat.Docx;
            _Document.filename = attachmentName;
            _Document._file = fileBytes;
            doc.ExportDocToPdf(ref _Document);
            if (_Document.IsSuccess == true)
            {
             fileBytes = System.IO.File.ReadAllBytes(_Document.returnPath);
            }
            _Document._file = fileBytes;
            return File(fileBytes, "application/pdf", _Document.returnPath + ".pdf");           
        }
        [HttpPost]
        [Route("XlsToPdf")]
        public async Task<IActionResult> XlsToPdf(IFormFile _file)
        {
            DocumentConvertor doc = new DocumentConvertor();
            FileDocument _Document = new FileDocument();
            string ContentType, attachmentName = "";
            byte[] fileBytes = null;

            ContentType = _file.ContentType;
            attachmentName = _file.FileName;
            using (var memoryStream = new MemoryStream())
            {
                await _file.CopyToAsync(memoryStream);
                fileBytes = memoryStream.ToArray();
            }            
            _Document.filename = attachmentName;
            _Document._file = fileBytes;
            doc.ExportXlsToPdf(ref _Document);
            if (_Document.IsSuccess == true)
            {
                fileBytes = System.IO.File.ReadAllBytes(_Document.returnPath);
            }
            return File(fileBytes, "application/pdf", _Document.returnPath + ".pdf");
        }
        [HttpPost]
        [Route("ImageToPdf")]
        public async Task<IActionResult> ImageToPdf(IFormFile _file)
        {
            DocumentConvertor doc = new DocumentConvertor();
            FileDocument _Document = new FileDocument();
            string ContentType, attachmentName = "";
            byte[] fileBytes = null;

            ContentType = _file.ContentType;
            attachmentName = _file.FileName;
            using (var memoryStream = new MemoryStream())
            {
                await _file.CopyToAsync(memoryStream);
                fileBytes = memoryStream.ToArray();
            }
            _Document.filename = attachmentName;
            _Document._file = fileBytes;
            doc.ImageToPdf(ref _Document);
            if (_Document.IsSuccess == true)
            {
                fileBytes = System.IO.File.ReadAllBytes(_Document.returnPath);
            }
            return File(fileBytes, "application/pdf", _Document.returnPath + ".pdf");
        }
    }
}
