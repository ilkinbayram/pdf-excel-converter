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

                //Create a PdfTableExtractor instance
                PdfTableExtractor extractor = new PdfTableExtractor(pdf);
                //Extract tables from the first page
                PdfTable[] pdfTables = extractor.ExtractTable(0);

                //If any tables are found
                if (pdfTables != null && pdfTables.Length > 0)
                {
                    //Loop through the tables
                    for (int tableNum = 0; tableNum < pdfTables.Length; tableNum++)
                    {
                        //Create a Workbook object,
                        Workbook wb = new Workbook();
                        //Remove default worksheets
                        wb.Worksheets.Clear();

                        //Add a worksheet to workbook
                        String sheetName = String.Format("Table - {0}", tableNum + 1);
                        Worksheet sheet = wb.Worksheets.Add(sheetName);
                        //Loop through the rows in the current table
                        for (int rowNum = 0; rowNum < pdfTables[tableNum].GetRowCount(); rowNum++)
                        {
                            //Loop through the columns in the current table
                            for (int colNum = 0; colNum < pdfTables[tableNum].GetColumnCount(); colNum++)
                            {
                                //Extract data from the current table cell
                                String text = pdfTables[tableNum].GetText(rowNum, colNum);
                                //Insert data into a specific cell
                                sheet.Range[rowNum + 1, colNum + 1].Text = text;
                            }
                        }

                        //Auto fit column width
                        sheet.AllocatedRange.AutoFitColumns();

                        //Save the workbook to an Excel file
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
            // Retrieve a reference to the workbook part.
            doxml.WorkbookPart workbookPart = document.WorkbookPart;

            // Get the Sheets element from the workbook.
            doxmlspr.Sheets sheets = workbookPart.Workbook.GetFirstChild<doxmlspr.Sheets>();

            // Determine the number of sheets.
            int numSheets = sheets.Count();

            // Get the relationship id of the last sheet
            string relId = sheets.Elements<doxmlspr.Sheet>().ElementAt(numSheets - 1).Id;

            // Remove the sheet reference from the workbook.
            doxmlspr.Sheet theSheet = workbookPart.Workbook.Descendants<doxmlspr.Sheet>().Where(s => s.Id.Value.Equals(relId)).First();
            theSheet.Remove();

            // Save the workbook.
            workbookPart.Workbook.Save();
        }
    }
}
