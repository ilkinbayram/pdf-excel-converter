using Microsoft.AspNetCore.Mvc;
using PdfApi.Service;

namespace PdfApi.Controllers;

[ApiController]
[Route("files")]
public class FilesController : ControllerBase
{
    private readonly IFileHandlerService _fileHandlerService;

    public FilesController(IFileHandlerService fileHandlerService)
    {
        _fileHandlerService = fileHandlerService;
    }

    [HttpPost("add-pdf")]
    public async Task<IActionResult> Index(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("No file is uploaded");
        }

        // Dosya byte array'a çevrilir
        using var memoryStream = new MemoryStream();
        await file.CopyToAsync(memoryStream);
        byte[] fileBytes = memoryStream.ToArray();

        string fileNameParam = Path.GetFileNameWithoutExtension(file.FileName);

        // PDF dosyasını Excel'e çevir
        var result = _fileHandlerService.ExtractTablesFromPdfToExcel(fileBytes, fileNameParam);

        if (result)
        {
            return Ok("File has been successfully converted to Excel and saved.");
        }
        else
        {
            return BadRequest("File conversion failed.");
        }
    }
}
