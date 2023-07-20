using Bytescout.PDFExtractor;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;

namespace PdfApi.Service;

public class AtbFileHandler : IFileHandlerService
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
            MoveDataInExcel(outputFilePath);
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

            var w1_A1 = GetCleanCellData(worksheet1,"A2");
            var w1_A2 = GetCleanCellData(worksheet1,"A3");
            var w1_A3 = GetCleanCellData(worksheet1,"A4");
            var w1_A4 = GetCleanCellData(worksheet1,"A5");
            var w1_A5 = GetCleanCellData(worksheet1,"A6");
            var w1_A6 = GetCleanCellData(worksheet1,"A7");
            var w1_A7 = GetCleanCellData(worksheet1,"A8");
            var w1_B1 = GetCleanCellData(worksheet1,"C2");
            var w1_B2 = GetCleanCellData(worksheet1,"C3");
            var w1_B3 = GetCleanCellData(worksheet1,"C4");
            var w1_B4 = GetCleanCellData(worksheet1,"C5");
            var w1_B5 = GetCleanCellData(worksheet1,"C6");
            var w1_B6 = GetCleanCellData(worksheet1,"C7");
            var w1_B7 = GetCleanCellData(worksheet1,"C8");
            
            var w2_A1 = GetCleanCellData(worksheet1,"A9");
            var w2_B1 = GetCleanCellData(worksheet1,"B9");
            var w2_C1 = GetCleanCellData(worksheet1,"C9");
            var w2_D1 = GetCleanCellData(worksheet1,"D9");
            var w2_E1 = GetCleanCellData(worksheet1,"E9");
            var w2_F1 = GetCleanCellData(worksheet1,"F9");
            var w2_G1 = GetCleanCellData(worksheet1,"G9");
            var w2_H1 = GetCleanCellData(worksheet1,"H9");
            
            var w2_A2 = GetCleanCellData(worksheet1,"A12");
            var w2_B2 = GetCleanCellData(worksheet1,"B12");
            var w2_C2 = GetCleanCellData(worksheet1,"E20");
            var w2_D2 = GetCleanCellData(worksheet1,"D12");
            var w2_E2 = GetCleanCellData(worksheet1,"E12");
            var w2_F2 = GetCleanCellData(worksheet1,"F12");
            var w2_G2 = GetCleanCellData(worksheet1,"G12");
            var w2_H2 = $"{GetCleanCellData(worksheet1,"H10")} {GetCleanCellData(worksheet1,"H11")} {GetCleanCellData(worksheet1,"H13")} {GetCleanCellData(worksheet1,"H14")}";
            var w2_A3 = GetCleanCellData(worksheet1,"A17");
            var w2_B3 = GetCleanCellData(worksheet1,"B17");
            var w2_C3 = GetCleanCellData(worksheet1,"E20");
            var w2_D3 = GetCleanCellData(worksheet1,"D17");
            var w2_E3 = GetCleanCellData(worksheet1,"E17");
            var w2_F3 = GetCleanCellData(worksheet1,"F17");
            var w2_G3 = GetCleanCellData(worksheet1,"G17");
            var w2_H3 = $"{GetCleanCellData(worksheet1,"H15")} {GetCleanCellData(worksheet1,"H16")} {GetCleanCellData(worksheet1,"H18")} {GetCleanCellData(worksheet1,"H19")}";
            
            package.Workbook.Worksheets.Delete(worksheet1);

            var ws1 = package.Workbook.Worksheets.Add("Table-1");
            var ws2 = package.Workbook.Worksheets.Add("Table-2");

            ws1.Cells["A1"].Value = w1_A1;
            ws1.Cells["A2"].Value = w1_A2;
            ws1.Cells["A3"].Value = w1_A3;
            ws1.Cells["A4"].Value = w1_A4;
            ws1.Cells["A5"].Value = w1_A5;
            ws1.Cells["A6"].Value = w1_A6;
            ws1.Cells["A7"].Value = w1_A7;
            ws1.Cells["B1"].Value = w1_B1;
            ws1.Cells["B2"].Value = w1_B2;
            ws1.Cells["B3"].Value = w1_B3;
            ws1.Cells["B4"].Value = w1_B4;
            ws1.Cells["B5"].Value = w1_B5;
            ws1.Cells["B6"].Value = w1_B6;
            ws1.Cells["B7"].Value = w1_B7;

            ws2.Cells["A1"].Value = w2_A1;
            ws2.Cells["B1"].Value = w2_B1;
            ws2.Cells["C1"].Value = w2_C1;
            ws2.Cells["D1"].Value = w2_D1;
            ws2.Cells["E1"].Value = w2_E1;
            ws2.Cells["F1"].Value = w2_F1;
            ws2.Cells["G1"].Value = w2_G1;
            ws2.Cells["H1"].Value = w2_H1;
            ws2.Cells["A2"].Value = w2_A2;
            ws2.Cells["B2"].Value = w2_B2;
            ws2.Cells["C2"].Value = w2_C2;
            ws2.Cells["D2"].Value = w2_D2;
            ws2.Cells["E2"].Value = w2_E2;
            ws2.Cells["F2"].Value = w2_F2;
            ws2.Cells["G2"].Value = w2_G2;
            ws2.Cells["H2"].Value = w2_H2;
            ws2.Cells["A3"].Value = w2_A3;
            ws2.Cells["B3"].Value = w2_B3;
            ws2.Cells["C3"].Value = w2_C3;
            ws2.Cells["D3"].Value = w2_D3;
            ws2.Cells["E3"].Value = w2_E3;
            ws2.Cells["F3"].Value = w2_F3;
            ws2.Cells["G3"].Value = w2_G3;
            ws2.Cells["H3"].Value = w2_H3;
            
            // Save the changes
            package.Save();
        }
    }

    private string GetCleanCellData(ExcelWorksheet worksheet, string address)
    {
        try
        {
            var result = worksheet.Cells[address].Value.ToString();
            return result;
        }
        catch (Exception e)
        {
            return string.Empty;
        }
    }
}

