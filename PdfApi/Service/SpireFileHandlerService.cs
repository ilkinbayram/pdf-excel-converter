using System;
using Spire.Pdf;
using Spire.Pdf.Utilities;
using Spire.Xls;
using doxml = DocumentFormat.OpenXml.Packaging;
using doxmlspr = DocumentFormat.OpenXml.Spreadsheet;

namespace PdfApi.Service;

public class SpireFileHandlerService : IFileHandlerService
{
    public bool ExtractTablesFromPdfToExcel(byte[] fileBytes, string fileNameParam)
    {
        if (fileBytes == null || fileBytes.Length == 0)
            return false;

        try
        {
            using (var memoryStream = new MemoryStream(fileBytes))
            {
                PdfDocument pdf = new PdfDocument(memoryStream);

                PdfTableExtractor extractor = new PdfTableExtractor(pdf);
                
                PdfTable[] pdfTables = extractor.ExtractTable(0);

                if (pdfTables != null && pdfTables.Length > 0)
                {
                    for (int tableNum = 0; tableNum < pdfTables.Length; tableNum++)
                    {
                        Workbook wb = new Workbook();
                        
                        wb.Worksheets.Clear();

                        String sheetName = String.Format("Table - {0}", tableNum + 1);
                        Worksheet sheet = wb.Worksheets.Add(sheetName);
                        
                        for (int rowNum = 0; rowNum < pdfTables[tableNum].GetRowCount(); rowNum++)
                        {
                            for (int colNum = 0; colNum < pdfTables[tableNum].GetColumnCount(); colNum++)
                            {
                                String text = pdfTables[tableNum].GetText(rowNum, colNum);
                                
                                sheet.Range[rowNum + 1, colNum + 1].Text = text;
                            }
                        }

                        sheet.AllocatedRange.AutoFitColumns();

                        string fileName = String.Format("{0}-ExportedExcel-{1}.xlsx", fileNameParam, tableNum + 1);
                        if (!Directory.Exists("C:\\ExcelDosyalar"))
                        {
                            Directory.CreateDirectory("C:\\ExcelDosyalar");
                        }
                        string filePath = Path.Combine("C:\\ExcelDosyalar", fileName);
                        wb.SaveToFile(filePath, ExcelVersion.Version2016);
                        
                        RemoveLastSheet(filePath);
                    }
                }

                return true;
            }
        }
        catch (Exception ex)
        {
            return false;
        }
    }
    
    private void RemoveLastSheet(string filePath)
    {
        using (doxml.SpreadsheetDocument document = doxml.SpreadsheetDocument.Open(filePath, true))
        {
            doxml.WorkbookPart workbookPart = document.WorkbookPart;

            doxmlspr.Sheets sheets = workbookPart.Workbook.GetFirstChild<doxmlspr.Sheets>();

            int numSheets = sheets.Count();

            string relId = sheets.Elements<doxmlspr.Sheet>().ElementAt(numSheets - 1).Id;

            doxmlspr.Sheet theSheet = workbookPart.Workbook.Descendants<doxmlspr.Sheet>().Where(s => s.Id.Value.Equals(relId)).First();
            theSheet.Remove();

            workbookPart.Workbook.Save();
        }
    }
}
