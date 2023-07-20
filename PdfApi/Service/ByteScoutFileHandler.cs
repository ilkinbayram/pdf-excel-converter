using Bytescout.PDFExtractor;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;
using PdfApi.Service;

public class ByteScoutFileHandler : IFileHandlerService
{
    public bool ExtractTablesFromPdfToExcel(byte[] fileBytes, string fileNameParam)
    {
        //Create a temporary file from byte array
        string tempFilePath = Path.Combine(Path.GetTempPath(), fileNameParam);
        File.WriteAllBytes(tempFilePath, fileBytes);

        // Ensure the ExcelDosyalar folder exists
        string outputDirectory = Path.Combine("C:", "ExcelDosyalar");
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        // Output file path
        string outputFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(fileNameParam) + ".xlsx");

        // Create Bytescout.PDFExtractor.XLSExtractor instance
        XLSExtractor extractor = new XLSExtractor();
        extractor.RegistrationName = "demo";
        extractor.RegistrationKey = "demo";

        // Load the document from the temporary file path
        extractor.LoadDocumentFromFile(tempFilePath);

        // Uncomment this line if you need all pages converted into a single worksheet:
        //extractor.PageToWorksheet = false;

        // Set the output format to XLSX
        extractor.OutputFormat = SpreadseetOutputFormat.XLSX;
    
        // Save the spreadsheet to file
        extractor.SaveToXLSFile(outputFilePath);

        // Cleanup
        extractor.Dispose();

        // Clean up the temporary file
        File.Delete(tempFilePath);

        // Check if file is created successfully
        if (File.Exists(outputFilePath))
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    
    public void MoveDataInExcel(string filePath)
    {
        // Ensure EPPlus considers the file as coming from a trusted source
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet1 = package.Workbook.Worksheets[0]; // first worksheet

            // Move the data for the first sheet
            MoveCellData(worksheet1, "A2", "A1");
            MoveCellData(worksheet1, "A3", "A2");
            MoveCellData(worksheet1, "A4", "A3");
            MoveCellData(worksheet1, "A5", "A4");
            MoveCellData(worksheet1, "A6", "A5");
            MoveCellData(worksheet1, "A7", "A6");
            MoveCellData(worksheet1, "A8", "A7");
            MoveCellData(worksheet1, "2", "A1");
            MoveCellData(worksheet1, "3", "A2");
            MoveCellData(worksheet1, "4", "A3");
            MoveCellData(worksheet1, "5", "A4");
            MoveCellData(worksheet1, "6", "A5");
            MoveCellData(worksheet1, "7", "A6");
            MoveCellData(worksheet1, "8", "A7");
            // ... Continue for all the cells for the first sheet

            // Add a new worksheet
            var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");

            // Move the data for the second sheet
            MoveCellData(worksheet2, "A9", "A1");
            MoveCellData(worksheet2, "B9", "B1");
            // ... Continue for all the cells for the second sheet

            // Clear the remaining cells
            ClearRemainingCells(worksheet1);
            ClearRemainingCells(worksheet2);

            // Save the changes
            package.Save();
        }
    }

    public void MoveCellData(ExcelWorksheet worksheet, string fromCell, string toCell)
    {
        var value = worksheet.Cells[fromCell].Value;
        worksheet.Cells[toCell].Value = value;
        worksheet.Cells[fromCell].Clear();
    }

    public void ClearRemainingCells(ExcelWorksheet worksheet)
    {
        // Assuming you want to clear all cells after the last row and column you've changed
        for (int i = 1; i <= worksheet.Dimension.Rows; i++)
        {
            for (int j = 1; j <= worksheet.Dimension.Columns; j++)
            {
                // Don't clear if this cell is already empty
                if (worksheet.Cells[i, j].Value == null)
                    continue;

                // Clear the cell
                worksheet.Cells[i, j].Clear();
            }
        }
    }
}
