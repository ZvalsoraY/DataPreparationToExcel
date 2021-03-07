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
        //string fileName = $"20tfiEMAT_20_ver_st3_WT10_sens3p5";
        static public List<string> list = new List<string>();
        
        static public void consoleWriteCheck(List<string> list)
        {
            foreach (var s in list)
                Console.Write(" " + s);
            //Console.Write(s);
            Console.WriteLine();
        }
        
        static void consoleWriteCheck(List<List<double>> writingArray)
        {
            foreach (var item in writingArray)
            {
                foreach (var s in item)
                    Console.Write(" " + s);
                //Console.Write(s);
                Console.WriteLine();
            }
        }

        static public void createExcel(List<string> list)
        {
            foreach(var inputFileName in list)
            {
                var doubleArrayListData = FileDoubleArrayList(inputFileName);
                List<List<double>> fileDataList = doubleArrayListData.Item1;
                foreach(var filterZ in doubleArrayListData.Item2)
                {
                    CreateExcel(sortByY(selectCoordZ(fileDataList, filterZ)), inputFileName, filterZ);
                }
                //CreateExcel(sortByY(selectCoordZ(fileDataList)), inputFileName);

            }
        }
        
        //static Tuple<List<List<double>>, List<double>> FileDoubleArrayList(string fileName)
        static (List<List<double>>, List<double>) FileDoubleArrayList(string fileName)
        {
            var lLdoubleArray = new List<List<double>>();
            var listFilterZ = new List<double>();
            var lines = File.ReadAllLines(fileName);
            for (int i = 0; i < lines.Length; i++)
            {
                var resString = new List<double>();

                string[] stringArray = lines[i].Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                foreach (var linPer in stringArray)
                {
                    resString.Add(Math.Round(double.Parse(linPer, CultureInfo.InvariantCulture), 2));
                }
                lLdoubleArray.Add(resString);
                listFilterZ.Add(resString[2]);
            }
            listFilterZ.Distinct();
            //consoleWriteCheck(lLdoubleArray);
            return (lLdoubleArray, listFilterZ);
        }

        static List<List<double>> selectCoordZ(List<List<double>> arraySelectZ, double zCoord = 0.0)
        {
            var zArray = new List<List<double>>();
            zArray.Clear();
            foreach (var tag in arraySelectZ)
            {
                if (Math.Abs(tag[2] - zCoord) <= 0.01) zArray.Add(tag);
            }
            //consoleWriteCheck(zArray);
            return zArray;
        }

        static List<List<double>> sortByY(List<List<double>> arraySortY)
        {
            var arrayByY = new List<List<double>>();
            arrayByY = arraySortY.OrderBy(l => l[1]).ToList();
            //consoleWriteCheck(resArray);
            return arrayByY;
        }

        static void CreateExcel(List<List<double>> inputArray, string fileNames, double axisCoordinateData = 0.0)
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
                        rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(Double.Parse(d.ToString().Replace(@".", @",")));
                    }
                }
                wb.Write(stream);
                wb.Close();
            }
        }
    }
}
