using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace ExcelCosmo
{
    class ExcelReaderCreator
    {
        /// <summary>
        /// Read Spreadsheet data.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static ICollection<CosmoSheet> ReadSpreadsheetCosmo(string filePath)
        {
            var excel = new ExcelQueryFactory(filePath);
            var lstCosmo = (from ico in excel.WorksheetNoHeader() select new CosmoSheet
                {
                    CompleteName = ico[0],
                    Address = ico[1],
                    CityStateZip = ico[2]
                }).ToList();
            return lstCosmo;
        }

        /// <summary>
        /// Creating Customized Excel Sheet.
        /// </summary>
        /// <param name="filePath">The File path</param>
        public static void CreateComosExcel(string filePath)
        {
            var cosmoSheetRowsCollection = ReadSpreadsheetCosmo(filePath);

            // Creating excel package
            ExcelPackage _ExcelPackage = new ExcelPackage();
            // Adding worksheet name
            ExcelWorksheet _ExcelWorksheet = _ExcelPackage.Workbook.Worksheets.Add("0");
            _ExcelWorksheet.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            _ExcelWorksheet.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
            _ExcelWorksheet.View.FreezePanes(2, 8);

            int rowIndex = 1;
            // Creating Header Row into Excel Sheet
            CreateHeader(_ExcelWorksheet, ref rowIndex);

            CreateData(_ExcelWorksheet, ref rowIndex, cosmoSheetRowsCollection);
            //Generate A File with Random name
            string file = "Cosmetologist" + "_" + Guid.NewGuid().ToString() + ".xlsx";
            using (MemoryStream ms = new MemoryStream())
            {
                _ExcelPackage.SaveAs(ms);
                File.WriteAllBytes(file, ms.ToArray());
            }
        }

        /// <summary>
        /// Creating Data Row
        /// </summary>
        /// <param name="_ExcelWorksheet">Excel Worksheet</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="cosmoSheetRowsCollection">Cosmo Sheet Rows Collection</param>
        private static void CreateData(ExcelWorksheet _ExcelWorksheet, ref int rowIndex, ICollection<CosmoSheet> cosmoSheetRowsCollection)
        {
            int colIndex = 1;
            object _ManualEntry = string.Empty;
            object _TotalEntryTime = string.Empty;
            foreach (var r in cosmoSheetRowsCollection)
            {
                colIndex = 1;
                rowIndex++;

                string[] completeNameArrary = r.CompleteName.Split(',');
                string[] firstLastNameArrary = completeNameArrary.Length > 1 ? completeNameArrary[1].Trim().Split(' ') : new string[0];
                string[] cityStateZipArrary = r.CityStateZip.Split(' ');
                foreach (string columnName in ColumnNames())
                {
                    var cell = _ExcelWorksheet.Cells[rowIndex, colIndex];
                    RowFormatting(cell);

                    switch (colIndex)
                    {
                        case 1: // First Name
                            cell.Value = firstLastNameArrary != null && firstLastNameArrary.Length > 0 ? firstLastNameArrary[0].Trim() : string.Empty;
                            break;
                        case 2: // Middle Name
                            cell.Value = firstLastNameArrary != null && firstLastNameArrary.Length > 1 ? firstLastNameArrary[1].Trim() : string.Empty;
                            break;
                        case 3: // Last Name
                            cell.Value = completeNameArrary != null && completeNameArrary.Length > 0 ? completeNameArrary[0].Trim() : string.Empty;
                            break;
                        case 4: // Address
                            cell.Value = !string.IsNullOrEmpty(r.Address) ? r.Address.Trim() : string.Empty;
                            break;
                        case 5: // City
                            cell.Value = cityStateZipArrary != null && cityStateZipArrary.Length > 0 ? cityStateZipArrary[0].Trim() : string.Empty;
                            break;
                        case 6: // State
                            cell.Value = cityStateZipArrary != null && cityStateZipArrary.Length > 1 ? cityStateZipArrary[1].Trim() : string.Empty;
                            break;
                        case 7: // Zip
                            cell.Value = cityStateZipArrary != null && cityStateZipArrary.Length > 2 ? cityStateZipArrary[2].Trim() : string.Empty;
                            break;
                    }

                    colIndex++;
                }
            }

        }

        /// <summary>
        /// Creating Header of Worksheet - First Row
        /// </summary>
        /// <param name="_ExcelWorksheet">Excel Worksheet</param>
        /// <param name="rowIndex">Row Index</param>
        private static void CreateHeader(ExcelWorksheet _ExcelWorksheet, ref int rowIndex)
        {
            int colIndex = 1;

            foreach (string columnName in ColumnNames())
            {
                var cell = _ExcelWorksheet.Cells[rowIndex, colIndex];
                HeaderFormatting(cell);
                cell.Value = columnName;

                switch (colIndex)
                {
                    case 1:
                        // FirstName- string
                        cell.AutoFitColumns(11.00);
                        break;
                    case 2:
                        // MiddleName- string
                        cell.AutoFitColumns(48.86);
                        break;
                    case 3:
                        // LastName - string
                        cell.AutoFitColumns(16.14);
                        break;
                    case 4:
                        // Address - string
                        cell.AutoFitColumns(10.43);
                        break;
                    case 5:
                        // City - string
                        cell.AutoFitColumns(11.29);
                        break;
                    case 6:
                        // State - string
                        cell.AutoFitColumns(9);
                        break;
                    case 7:
                        // Zip - String
                        cell.AutoFitColumns(45.14);
                        break;
                }

                colIndex++;
            }
        }

        /// <summary>
        /// Header Formatting
        /// </summary>
        /// <param name="rng">Cell Range</param>
        private static void HeaderFormatting(ExcelRange rng)
        {
            rng.Style.Font.Bold = true;
            rng.Style.Font.Size = 11;
            rng.Style.Font.Name = "Calibri";
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;    //Set Pattern for the background to Solid
            rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            rng.Style.Font.Color.SetColor(Color.Black);
            rng.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            rng.Style.WrapText = true;
            rng.Style.Locked = true;
        }

        /// <summary>
        /// Row Formatting
        /// </summary>
        /// <param name="rng">Cell Range</param>
        private static void RowFormatting(ExcelRange rng)
        {
            rng.Style.Font.Bold = false;
            rng.Style.Font.Size = 11;
            rng.Style.Font.Name = "Calibri";
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid; 
            rng.Style.Fill.BackgroundColor.SetColor(Color.White);
            rng.Style.Font.Color.SetColor(Color.Black);
            rng.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.General;
            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            rng.Style.WrapText = false;
            rng.Style.Locked = true;
        }

        /// <summary>
        /// List of Header Column Names
        /// </summary>
        /// <returns>List of Header Column Names</returns>
        private static List<string> ColumnNames()
        {
            List<string> lst = new List<string>();
            lst.Add("FirstName");
            lst.Add("MiddleName");
            lst.Add("LastName");
            lst.Add("Address");
            lst.Add("City");
            lst.Add("State");
            lst.Add("Zip");
            return lst;
        }
    }
}
