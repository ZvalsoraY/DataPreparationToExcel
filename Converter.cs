using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel.Charts;

namespace DataPreparationToExcel
{
    static class Converter
    {
        static public List<string> list = new List<string>();
        
        static void createExcelFile(string fileName) 
        {
            var listData = FileDoubleArrayList(fileName);
            var distanceDistinct = listData.Select(s => s[2]).Distinct().OrderBy(u => u);
            foreach (var dist in distanceDistinct) 
            {
                var massExcel = listData.Where(u => (Math.Abs(u[2] - dist) <= 0.01)).OrderBy(u => u[1]).ToList();
                generateExcel(massExcel, fileName, dist);
            }
        }

        static public void createExcelForListFiles(List<string> list)
        {
            foreach(var inputFileName in list)
            {
                createExcelFile(inputFileName);                
            }
        }

        
        static List<List<double>> FileDoubleArrayList(string fileName)
        {
            var stringArray = File.ReadAllLines(fileName).Select(x => x.Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries));
            var lLdoubleArray = new List<List<double>>();
            foreach (var linPer in stringArray)
            {
                lLdoubleArray.Add(linPer.Select(x => Math.Round(double.Parse(x, CultureInfo.InvariantCulture), 2)).ToList());
            }           
            return lLdoubleArray;
        }

        static void generateExcel(List<List<double>> inputArray, string fileNames, double axisCoordinateData = 0.0)
        {
            //string parth = $"D:\\test\\" + fileNames + $" dist {axisCoordinateData.ToString()}.xlsx";
            string parth = fileNames.Substring(0, fileNames.LastIndexOf('.')) + $" dist {axisCoordinateData.ToString()}.xlsx";
            using (var stream = new FileStream(parth, FileMode.Create, FileAccess.ReadWrite))
            {
                var wb = new XSSFWorkbook();
                //var wb = new ;
                var sheet = wb.CreateSheet($"Hsum dist {axisCoordinateData.ToString()}");
                //creating cell style for header
                var bStylehead = wb.CreateCellStyle();
                bStylehead.BorderBottom = BorderStyle.Thin;
                bStylehead.BorderLeft = BorderStyle.Thin;
                bStylehead.BorderRight = BorderStyle.Thin;
                bStylehead.BorderTop = BorderStyle.Thin;
                bStylehead.Alignment = HorizontalAlignment.Center;
                bStylehead.VerticalAlignment = VerticalAlignment.Center;
                bStylehead.FillBackgroundColor = HSSFColor.Green.Index;
                //var cellStyle =
                //var cellStyle = CreateCellStyleForHeader(wb);
                var Drawing = sheet.CreateDrawingPatriarch();
                //IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 8, 1, 18, 16);
                var anchor = Drawing.CreateAnchor(0, 0, 0, 0, 8, 1, 23, 16);
                var chart = Drawing.CreateChart(anchor);
                //IChart chart = Drawing.CreateChart(anchor);
                IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
                IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

                var chartData =
                        chart.ChartDataFactory.CreateLineChartData<double, double>();
                var lenCellRange = inputArray.Count + 1;
                IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf($"B2:B{lenCellRange}"));
                IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf($"G2:G{lenCellRange}"));
                //IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("G2:G20"));

                var series = chartData.AddSeries(xs, ys);
                series.SetTitle("Hsum");
                //chart.GetOrCreateLegend();
                chart.Plot(chartData, bottomAxis, leftAxis);

                var row = sheet.CreateRow(0);

                row.CreateCell(0, CellType.String).SetCellValue("x");
                row.CreateCell(1, CellType.String).SetCellValue("y");
                row.CreateCell(2, CellType.String).SetCellValue("z");
                row.CreateCell(3, CellType.String).SetCellValue("Hx");
                row.CreateCell(4, CellType.String).SetCellValue("Hy");
                row.CreateCell(5, CellType.String).SetCellValue("Hz");
                row.CreateCell(6, CellType.String).SetCellValue("Hsum");
                //row.Cells[0].CellStyle = bStylehead;
                //row.RowStyle = bStylehead;

                //filling the data
                var rowsCounter = 1;
                foreach (var rowData in inputArray)
                {
                    var rowD = sheet.CreateRow(rowsCounter++);
                    var dCounter = 0;
                    foreach (var d in rowData)
                    {
                        //rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(Double.Parse(d.ToString().Replace(@".", @",")));
                        //CultureInfo.InvariantCulture
                        rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(d);
                    }
                }
                wb.Write(stream);
                wb.Close();
            }
        }
    }
}
