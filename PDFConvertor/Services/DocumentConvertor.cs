using Spire.Doc;


namespace PDFConvertor.Services
{
    public class DocumentConvertor
    {
        private const string _rootpath = "TempFiles\\";
        public string _filename = "Tempfile_" + DateTime.Now.ToShortDateString();
        public string _pdfFilePath = "";
        public string _sourceFilePath = "";
        public FileDocument ExportToPdf(ref FileDocument _Document)
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
                    _sourceFilePath = Path.Combine(Directory.GetCurrentDirectory(), _rootpath , _Document.filename);
                    _pdfFilePath = Path.GetFileName(_Document.filename).Replace(".docx","").Replace(".doc","").Replace(".html","").Replace(".htm", "");
                    using (var fs = new FileStream(_sourceFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(_Document._file, 0, _Document._file.Length);
                    }
                    document.LoadFromFile(_sourceFilePath);
                    document.SaveToFile(_pdfFilePath, FileFormat.PDF);
                    if (File.Exists(_sourceFilePath))
                    {
                        File.Delete(_sourceFilePath);
                    }
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

        
    }



    public class FileDocument
    {
        public byte[]? _file { get; set; }
        public FileFormat format { get; set; }
        public string? filename { get; set; }
        public bool? IsSuccess  { get; set; } 
        public string? returnPath { get; set; }
    }
}
