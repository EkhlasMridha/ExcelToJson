using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Web;

namespace ExelParser.Services
{
    public class ExcelService
    {
        public ExcelService()
        {
        }

        public DataTable ReadExcel(string sheetName)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "ExcelSheets", "Paint_Calculator-1.xlsx");
            DataTable dataTable = new DataTable();

            SpreadsheetDocument doc = OpenSpreaSheet(path);

            var workSheet = GetWorkSheetByName(sheetName, doc);

            DataTable data = GetWorksheetToDataTable(workSheet,doc);

            return data;
        }

        public Worksheet GetWorkSheetByName(string sheetname, SpreadsheetDocument doc)
        {
            WorkbookPart workbookPart = doc.WorkbookPart;
            string relationId = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.Equals(sheetname))?.Id;

            Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(relationId)).Worksheet;

            return worksheet;
        }

        public static SpreadsheetDocument OpenSpreaSheet(string path)
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false);

            return doc;
        }

        public DataTable GetWorksheetToDataTable(Worksheet worksheet, SpreadsheetDocument doc)
        {
            WorkbookPart workbookPart = doc.WorkbookPart;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            DataTable dataTable = new DataTable();
            int numOfColumn = sheetData.ElementAt(0).ChildElements.Count();

            for(int rcnt = 0; rcnt < sheetData.ChildElements.Count; ++rcnt)
            {
                List<string> rowList = new List<string>();
                var cnt = sheetData.ElementAt(rcnt).ChildElements.Count();
                for (int rcnt2 = 0; rcnt2 < numOfColumn; ++rcnt2)
                {
                    Cell currentCell = (Cell)sheetData.ElementAt(rcnt).ChildElements.ElementAt(rcnt2);

                    if(currentCell.DataType != null)
                    {
                        if(currentCell.DataType == CellValues.SharedString)
                        {
                            int id;
                            if(Int32.TryParse(currentCell.InnerText, out id))
                            {
                                SharedStringItem sharedStringItem = workbookPart.SharedStringTablePart.
                                                                    SharedStringTable.Elements<SharedStringItem>()
                                                                    .ElementAt(id);
                                if(sharedStringItem.Text != null)
                                {
                                    if(rcnt == 0)
                                    {
                                        dataTable.Columns.Add(sharedStringItem.Text.Text);
                                    }
                                    else
                                    {
                                        rowList.Add(sharedStringItem.Text.Text);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if(rcnt != 0)
                        {
                            rowList.Add(currentCell.InnerText);
                        }
                    }
                }
                if(rcnt != 0)
                {
                    dataTable.Rows.Add(rowList.ToArray());
                }
            }

            return dataTable;
        }
           
    }
}
