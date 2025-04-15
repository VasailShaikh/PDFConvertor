namespace PDFConvertor.Services
{
    public interface IWordToPdf
    {
        FileDocument ExportDocToPdf(ref FileDocument _Document);

    }
}
