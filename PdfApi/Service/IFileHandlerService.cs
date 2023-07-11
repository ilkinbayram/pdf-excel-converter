namespace PdfApi.Service;

public interface IFileHandlerService
{
    bool ExtractTablesFromPdfToExcel(byte[] fileBytes, string fileNameParam);
}
