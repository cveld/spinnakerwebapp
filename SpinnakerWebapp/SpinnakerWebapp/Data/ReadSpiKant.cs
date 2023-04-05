using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpinnakerWebapp.Entities;
using System.Text.RegularExpressions;

namespace SpinnakerWebapp.Data
{
    public class ReadSpiKant
    {
        public List<Baan> Read()
        {
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(@".\Data\spi kant.xlsx", false))
            {
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                var result = new List<Baan>();
                var baan = new Baan();

                foreach (Row row in rows) 
                {
                    var rowindex = (row.RowIndex - 2) % 4;
                    if (rowindex == 0)
                    {
                        baan = new Baan();
                        result.Add(baan);
                    }
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        Cell cell = row.Descendants<Cell>().ElementAt(i);
                        var columnIndexOptional = GetColumnIndex(cell.CellReference);
                        if (columnIndexOptional == null) continue;
                        int columnIndex = (int)columnIndexOptional; 
                        var cellValue = GetCellValue(spreadSheetDocument, cell);

                        switch (rowindex)
                        {
                            case 0:
                                if (columnIndex == 2)
                                {
                                    baan.Windrichting = cellValue;
                                }
                                break;
                            case 1:
                                if (columnIndex > 2)
                                {
                                    if (!String.IsNullOrEmpty(cellValue))
                                    {
                                        baan.Boeien.Add(new Boei
                                        {
                                            Name = cellValue,
                                            OriginalColumnIndex = columnIndex
                                        });
                                    }
                                }
                                break;
                            case 2:
                                if (columnIndex > 2)
                                {
                                    var boei = baan.Boeien.FirstOrDefault(b => b.OriginalColumnIndex == columnIndex);
                                    if (boei != null)
                                    {
                                        boei.WindAngle = Int32.Parse(cellValue);
                                    }
                                }
                                break;
                        } 
                    }
                }

                return result;
            }
        }

        private static int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return null;
            }

            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber + 1;
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue?.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }

        }
    }
}