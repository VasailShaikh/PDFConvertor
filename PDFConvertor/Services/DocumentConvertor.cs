using Spire.Doc;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Spire.Presentation;
using Spire.Xls;
using System.Drawing;

namespace PDFConvertor.Services
{
    public class DocumentConvertor: IWordToPdf, IxlsToPdf
    {
        private const string _rootpath = "TempFiles\\";
        public string _filename = "Tempfile_" + DateTime.Now.ToShortDateString();
        public string _pdfFilePath = "";
        public string _tempFilePath = "";
        public FileDocument ExportDocToPdf(ref FileDocument _Document)
        {
            Document document = new Document();
            try
            {
                if (_Document._file != null)
                {
                    if (string.IsNullOrWhiteSpace(_Document.filename))
                    {
                        _Document.filename = _filename;
                    }
                    _tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), _rootpath , _Document.filename);
                    _pdfFilePath = Path.GetFileName(_Document.filename).Replace(".docx","").Replace(".doc","").Replace(".html","").Replace(".htm", "");
                    using (var fs = new FileStream(_tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(_Document._file, 0, _Document._file.Length);
                    }
                    document.LoadFromFile(_tempFilePath);
                    document.SaveToFile(_pdfFilePath, Spire.Doc.FileFormat.PDF);
                    DeleteTempFiles(_tempFilePath);
                    _Document.returnPath = _pdfFilePath;
                    _Document.IsSuccess = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Document;
        }

        public FileDocument ExportXlsToPdf(ref FileDocument _Document)
        {
            try
            {
                Workbook workbook = new Workbook();
                if (_Document._file != null)
                {
                    if (string.IsNullOrWhiteSpace(_Document.filename))
                    {
                        _Document.filename = _filename;
                    }
                    _tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), _rootpath, _Document.filename);
                    _pdfFilePath = Path.GetFileName(_Document.filename).Replace(".xls", "").Replace(".xlsx", "").Replace(".csv", "");
                    using (var fs = new FileStream(_tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(_Document._file, 0, _Document._file.Length);
                    }
                    //Load a .xlsx (or .xls) file
                    workbook.LoadFromFile(_tempFilePath);
                    //Set worksheets to fit to page when converting
                    workbook.ConverterSetting.SheetFitToPage = true;
                    //Save to PDF
                    workbook.SaveToFile(_pdfFilePath, Spire.Xls.FileFormat.PDF);
                    DeleteTempFiles(_tempFilePath);
                    _Document.returnPath = _pdfFilePath;
                    _Document.IsSuccess = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Document;
        }

        public FileDocument ImageToPdf(ref FileDocument _Document)
        {
            try
            {
                PdfDocument doc = new PdfDocument();
                doc.PageSettings.SetMargins(0);
                if (_Document._file != null)
                {
                    if (string.IsNullOrWhiteSpace(_Document.filename))
                    {
                        _Document.filename = _filename;
                    }
                    _tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), _rootpath, _Document.filename);
                    _pdfFilePath = Path.GetFileName(_tempFilePath).Replace("jpeg","").Replace("jpg","").Replace("png","");
                    using (var fs = new FileStream(_tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(_Document._file, 0, _Document._file.Length);
                    }
                }
             
                Image image = Image.FromFile(_tempFilePath);
                using (image) 
                {                  
                float width = image.PhysicalDimension.Width;
                float height = image.PhysicalDimension.Height;
           
                PdfPageBase page = doc.Pages.Add(new SizeF(width, height));              
                PdfImage pdfImage = PdfImage.FromImage(image);                
                page.Canvas.DrawImage(pdfImage, 0, 0, pdfImage.Width, pdfImage.Height);                
                doc.SaveToFile(_pdfFilePath, Spire.Pdf.FileFormat.PDF);
                }
                DeleteTempFiles(_tempFilePath);
                _Document.returnPath = _pdfFilePath;
                _Document.IsSuccess = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        return _Document;
        }

        public FileDocument PpToPdf(ref FileDocument _Document)
        {
            try
            {
                //Create a Presentation object
                Presentation presentation = new Presentation();
                if (_Document._file != null)
                {
                    if (string.IsNullOrWhiteSpace(_Document.filename))
                    {
                        _Document.filename = _filename;
                    }
                    _tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), _rootpath, _Document.filename);
                    _pdfFilePath = Path.GetFileName(_tempFilePath).Replace("ppt", "").Replace("pptx", "");
                    using (var fs = new FileStream(_tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(_Document._file, 0, _Document._file.Length);
                    }
                }
                using (presentation)
                { 
                   
                presentation.LoadFromFile(_tempFilePath);
                //SaveToPdfOption saveToPdfOption = presentation.SaveToPdfOption;              
                //saveToPdfOption.PdfSecurity.Encrypt("open-psd", "permission-psd", Spire.Pdf.Security.PdfPermissionsFlags.None, Spire.Pdf.Security.PdfEncryptionKeySize.Key128Bit);
                //Save to PDF
                presentation.SaveToFile(_pdfFilePath, Spire.Presentation.FileFormat.PDF);
                }
                
                _Document.returnPath = _pdfFilePath;
                _Document.IsSuccess = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Document;
        }

        public void DeleteTempFiles(string _tempPath)
        {
            if (File.Exists(_tempPath))
            {
                File.Delete(_tempPath);
            }
        }
    }

 

    public class FileDocument
    {
        public byte[]? _file { get; set; }
        public Spire.Doc.FileFormat format { get; set; }
        public string? filename { get; set; }
        public bool? IsSuccess  { get; set; } 
        public string? returnPath { get; set; }
    }
}
