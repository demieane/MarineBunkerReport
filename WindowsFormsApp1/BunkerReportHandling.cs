using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using Newtonsoft.Json;

/// <summary>
/// Summary description for Class1
/// </summary>
namespace WindowsFormsApp1
{ 
    public class BunkerReportHandling
    {
        public int index = new int();

        public void OverviewSheet(string FilePath, string[] AssetNames)
        {
            using (var BunkerReport = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = BunkerReport.Workbook.Worksheets.Add("Reports");
                //Set the column 1 width to 50
                double columnWidth = 15;

                index = 3;
                worksheet.Cells[index, 1].Value = "VESSELS";
                worksheet.Cells[index , 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[index , 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                
                foreach (string asset in AssetNames)
                {
                    worksheet.Column(1).Width = columnWidth + 5;
                    worksheet.Cells[index + 1,1].Value = asset;
                    

                    //worksheet.Cells[index + 1, 1].Style.Font.Bold = true;
                    worksheet.Cells[index + 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[index + 1, 1].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                    index = index + 1;

                }

                BunkerReport.SaveAs(new FileInfo(FilePath));
            }         
        }

        public string[] NewEntry(string FilePath, string AssetName)
        {
            FileInfo File = new FileInfo(FilePath);
            //if (File.Exists)
            //{

            using (var BunkerReport = new ExcelPackage(File))
            {
                // In the first worksheet
                ExcelWorksheet worksheet = BunkerReport.Workbook.Worksheets[1];


                var rowCnt = worksheet.Dimension.End.Row;
                var colCnt = worksheet.Dimension.End.Column;

                    Console.WriteLine("HELP row : {0}", rowCnt);
                    Console.WriteLine("HELP col : {0}", colCnt);

                // Find in which row the asset is
                var query =
                        from cell in worksheet.Cells[1, 1, rowCnt, colCnt]
                        where cell.Value?.ToString() == AssetName
                        select cell;

                object first = query.Cast<object>().First();

                //Console.WriteLine("FOUND {0}", CellCoordinates);
                string Coordinates = first.ToString();
                Console.WriteLine("Coordinates {0}", Coordinates);

                string[] cellID = Coordinates.Split('A');
                Console.WriteLine("cellID[0] {0}", cellID[0]);
                Console.WriteLine("cellID[1] {0}", cellID[1]);

                int numVal = Int32.Parse(cellID[1]);
                var lastRowCell1 = worksheet.Cells.Last(c => c.Start.Row == numVal);
                Console.WriteLine(lastRowCell1.Address);


                //worksheet.Cells[2, 3].Value = worksheet.Cells[Coordinates].Value;

                BunkerReport.Save();
                string[] Output = { Coordinates, lastRowCell1.Address};
                return Output;
            }
            //}
         
        }
        


        public void NewBunkerReport(string FilePath)
        {
        //
        // TODO: Add constructor logic here
        //

        // Create new bunker report excel file
        using (var NewBunkerReport = new ExcelPackage())
        {
            DateTime thisDay = DateTime.Today;
            Console.WriteLine(thisDay.ToString());

            // Add a new worksheet to the empty workbook
            ExcelWorksheet worksheet = NewBunkerReport.Workbook.Worksheets.Add("Fleet_Data");

            //worksheet.Cells.AutoFitColumns();
            //worksheet.Column(1).BestFit = true;

            //Set the column 1 width to 50
            double columnWidth = 15;

            //worksheet.Column(1).Width = columnWidth + 5;
            worksheet.Column(2).Width = columnWidth + 5;
            worksheet.Column(3).Width = columnWidth - 3 ;
            worksheet.Column(4).Width = columnWidth;
            worksheet.Column(5).Width = columnWidth;
            worksheet.Column(6).Width = columnWidth;

                worksheet.Cells[1, 1].Value = "BUNKER REPORT DATASHEET";

            //Ok now format the values;
            using (var range = worksheet.Cells[1, 1, 1, 6])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);
                //range.Style.Font.Color.SetColor(Color.White);
            }

            using (var range = worksheet.Cells[2, 1, 3, 6])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                //range.Style.Font.Color.SetColor(Color.White);
            }


            worksheet.Cells[1, 1, 1, 6].Merge = true; //Merge columns start and end range
            worksheet.Cells[1, 1, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

            worksheet.Cells[2, 1, 3, 1].Merge = true; //Merge columns start and end range
            worksheet.Cells[2, 1, 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

            worksheet.Cells[2, 2, 3, 2].Merge = true; //Merge columns start and end range
            worksheet.Cells[2, 2, 3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

            worksheet.Cells[2, 3, 3, 3].Merge = true; //Merge columns start and end range
            worksheet.Cells[2, 3, 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

            worksheet.Cells[2, 4, 3, 4].Merge = true; //Merge columns start and end range
            worksheet.Cells[2, 4, 3, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

            worksheet.Cells[2, 5, 3, 5].Merge = true; //Merge columns start and end range
            worksheet.Cells[2, 5, 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

            worksheet.Cells[2, 6, 3, 6].Merge = true; //Merge columns start and end range
            worksheet.Cells[2, 6, 3, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center


            //Add the headers
            worksheet.Cells[2, 1].Value = "Registry";
            worksheet.Cells[2, 2].Value = "Vessel";
            worksheet.Cells[2, 3].Value = "Position";
            worksheet.Cells[2, 4].Value = "Entry Date";
            worksheet.Cells[2, 5].Value = "FO Difference [MT]";
            worksheet.Cells[2, 6].Value = "DO Difference [MT]";

            worksheet.Column(5).Style.WrapText = true;
            worksheet.Column(6).Style.WrapText = true;



            //string filePath = @"C:\Users\tech006\Desktop\Praktiki2018\BunkerReport\NewBunkerReport.xlsx";
            NewBunkerReport.SaveAs(new FileInfo(FilePath));
        }
    }

        public void UpdateWorksheet(string FilePath, string WorkSheetName)
        {
            FileInfo File = new FileInfo(FilePath);
            if (File.Exists)
            {
                using (var BunkerReport = new ExcelPackage(File))
                {

                    // Add a new worksheet to the empty workbook
                    ExcelWorksheet worksheetNew = BunkerReport.Workbook.Worksheets.Add(WorkSheetName);
                    /*
                    //Set the column 1 width to 50
                    double columnWidth = 15;
                    //worksheet.Column(1).Width = columnWidth - 15;
                    worksheetNew.Column(2).Width = columnWidth + 5;
                    worksheetNew.Column(3).Width = columnWidth - 5;
                    worksheetNew.Column(4).Width = columnWidth ;
                    worksheetNew.Column(5).Width = columnWidth;
                    worksheetNew.Column(6).Width = columnWidth;
                    */

                    worksheetNew.Cells[1, 1].Value = "BUNKER REPORT DATASHEET";

                    //Ok now format the values;
                    using (var range = worksheetNew.Cells[1, 1, 1, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);
                        //range.Style.Font.Color.SetColor(Color.White);
                    }

                    using (var range = worksheetNew.Cells[2, 1, 3, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                        //range.Style.Font.Color.SetColor(Color.White);
                    }


                    worksheetNew.Cells[1, 1, 1, 6].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[1, 1, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                    worksheetNew.Cells[2, 1, 3, 1].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[2, 1, 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                    worksheetNew.Cells[2, 2, 3, 2].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[2, 2, 3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                    worksheetNew.Cells[2, 3, 3, 3].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[2, 3, 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                    worksheetNew.Cells[2, 4, 3, 4].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[2, 4, 3, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                    worksheetNew.Cells[2, 5, 3, 5].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[2, 5, 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                    worksheetNew.Cells[2, 6, 3, 6].Merge = true; //Merge columns start and end range
                    worksheetNew.Cells[2, 6, 3, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center


                    //Add the headers
                    worksheetNew.Cells[2, 1].Value = "Registry";
                    worksheetNew.Cells[2, 2].Value = "Vessel";
                    worksheetNew.Cells[2, 3].Value = "Position";
                    worksheetNew.Cells[2, 4].Value = "Report Date";
                    worksheetNew.Cells[2, 5].Value = "FO Difference [MT]";
                    worksheetNew.Cells[2, 6].Value = "DO Difference [MT]";

                    worksheetNew.Column(5).Style.WrapText = true;
                    worksheetNew.Column(6).Style.WrapText = true;

                    

                    BunkerReport.Save();

                    //string filePath = @"C:\Users\tech006\Desktop\Praktiki2018\BunkerReport\NewBunkerReport.xlsx";
                    //BunkerReport.SaveAs(new FileInfo(FilePath));
                }
            }
        }

    }
}