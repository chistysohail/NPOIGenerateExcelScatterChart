using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.SS.UserModel.Charts;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Create a new workbook
        IWorkbook workbook = new XSSFWorkbook();

        // Create the data sheet
        ISheet tableSheet = workbook.CreateSheet("Football Report");

        // Create a style for the header
        ICellStyle headerStyle = workbook.CreateCellStyle();
        headerStyle.Alignment = HorizontalAlignment.Center;
        headerStyle.VerticalAlignment = VerticalAlignment.Center;
        IFont headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        headerStyle.SetFont(headerFont);

        // Define background colors for different sections of the header
        ICellStyle greenStyle = workbook.CreateCellStyle();
        greenStyle.CloneStyleFrom(headerStyle);
        greenStyle.FillForegroundColor = IndexedColors.LightGreen.Index;
        greenStyle.FillPattern = FillPattern.SolidForeground;

        ICellStyle orangeStyle = workbook.CreateCellStyle();
        orangeStyle.CloneStyleFrom(headerStyle);
        orangeStyle.FillForegroundColor = IndexedColors.LightOrange.Index;
        orangeStyle.FillPattern = FillPattern.SolidForeground;

        ICellStyle blueStyle = workbook.CreateCellStyle();
        blueStyle.CloneStyleFrom(headerStyle);
        blueStyle.FillForegroundColor = IndexedColors.LightBlue.Index;
        blueStyle.FillPattern = FillPattern.SolidForeground;

        // Create the header rows
        IRow headerRow1 = tableSheet.CreateRow(0);
        IRow headerRow2 = tableSheet.CreateRow(1);
        IRow headerRow3 = tableSheet.CreateRow(2);

        // Define the merged regions (merged cells)
        tableSheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2)); // Merge first 3 cells for "League/Team/Player Name"
        tableSheet.AddMergedRegion(new CellRangeAddress(0, 0, 3, 4)); // Merge for "Goals Scored"
        tableSheet.AddMergedRegion(new CellRangeAddress(0, 0, 5, 6)); // Merge for "Assists"
        tableSheet.AddMergedRegion(new CellRangeAddress(1, 2, 0, 0)); // "League"
        tableSheet.AddMergedRegion(new CellRangeAddress(1, 2, 1, 1)); // "Team"
        tableSheet.AddMergedRegion(new CellRangeAddress(1, 2, 2, 2)); // "Player Name"
        tableSheet.AddMergedRegion(new CellRangeAddress(1, 1, 3, 4)); // Merge for "Goals Scored"
        tableSheet.AddMergedRegion(new CellRangeAddress(1, 1, 5, 6)); // Merge for "Assists"

        // Row 1 - Main headers
        ICell headerCell1 = headerRow1.CreateCell(0);
        headerCell1.SetCellValue("League / Team / Player Name");
        headerCell1.CellStyle = greenStyle;

        headerCell1 = headerRow1.CreateCell(3);
        headerCell1.SetCellValue("Goals Scored");
        headerCell1.CellStyle = orangeStyle;

        headerCell1 = headerRow1.CreateCell(5);
        headerCell1.SetCellValue("Assists");
        headerCell1.CellStyle = blueStyle;

        // Row 2 - Sub-headers (Merged cells for columns)
        ICell headerCell2 = headerRow2.CreateCell(0);
        headerCell2.SetCellValue("League");
        headerCell2.CellStyle = greenStyle;

        headerCell2 = headerRow2.CreateCell(1);
        headerCell2.SetCellValue("Team");
        headerCell2.CellStyle = greenStyle;

        headerCell2 = headerRow2.CreateCell(2);
        headerCell2.SetCellValue("Player Name");
        headerCell2.CellStyle = greenStyle;

        headerCell2 = headerRow2.CreateCell(3);
        headerCell2.SetCellValue("This Season");
        headerCell2.CellStyle = orangeStyle;

        headerCell2 = headerRow2.CreateCell(4);
        headerCell2.SetCellValue("Last Season");
        headerCell2.CellStyle = orangeStyle;

        headerCell2 = headerRow2.CreateCell(5);
        headerCell2.SetCellValue("This Season");
        headerCell2.CellStyle = blueStyle;

        headerCell2 = headerRow2.CreateCell(6);
        headerCell2.SetCellValue("Last Season");
        headerCell2.CellStyle = blueStyle;

        // Sample data for table
        string[] playerNames = { "Player 1", "Player 2", "Player 3", "Player 4" };
        double[] goalsScored = { 15, 18, 10, 20 };
        double[] assists = { 5, 7, 4, 6 };

        for (int i = 0; i < playerNames.Length; i++)
        {
            IRow row = tableSheet.CreateRow(i + 3); // Data starts after the header rows
            row.CreateCell(0).SetCellValue("League A");
            row.CreateCell(1).SetCellValue("Team X");
            row.CreateCell(2).SetCellValue(playerNames[i]);
            row.CreateCell(3).SetCellValue(goalsScored[i]); // Goals Scored This Season
            row.CreateCell(4).SetCellValue(goalsScored[i] * 1.1); // Goals Scored Last Season (dummy calculation)
            row.CreateCell(5).SetCellValue(assists[i]); // Assists This Season
            row.CreateCell(6).SetCellValue(assists[i] * 1.2); // Assists Last Season (dummy calculation)
        }

        // Auto-size columns for readability
        for (int i = 0; i < 7; i++)
        {
            tableSheet.AutoSizeColumn(i);
        }

        // Create a separate sheet for multiple charts
        ISheet chartSheet = workbook.CreateSheet("Performance Charts");

        // Create a drawing canvas on the chart sheet for multiple charts
        IDrawing drawing = chartSheet.CreateDrawingPatriarch();

        // Chart 1: Scatter Plot for Goals Scored This Season
        IClientAnchor anchor1 = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 10, 20);
        IChart chart1 = drawing.CreateChart(anchor1);
        IChartLegend legend1 = chart1.GetOrCreateLegend();
        legend1.Position = LegendPosition.TopRight;

        // Define chart data for first chart (Goals Scored vs Player Name)
        IChartDataSource<string> xs1 = DataSources.FromStringCellRange(tableSheet, new CellRangeAddress(3, 6, 2, 2)); // Player Names
        IChartDataSource<double> ys1 = DataSources.FromNumericCellRange(tableSheet, new CellRangeAddress(3, 6, 3, 3)); // Goals Scored

        IScatterChartData<string, double> scatterChartData1 = chart1.ChartDataFactory.CreateScatterChartData<string, double>();
        IScatterChartSeries<string, double> series1 = scatterChartData1.AddSeries(xs1, ys1);
        series1.SetTitle("Goals Scored This Season");

        IChartAxis bottomAxis1 = chart1.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
        IValueAxis leftAxis1 = chart1.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
        leftAxis1.Crosses = AxisCrosses.AutoZero;
        chart1.Plot(scatterChartData1);

        // Chart 2: Scatter Plot for Assists This Season
        IClientAnchor anchor2 = drawing.CreateAnchor(0, 0, 0, 0, 12, 0, 22, 20);
        IChart chart2 = drawing.CreateChart(anchor2);
        IChartLegend legend2 = chart2.GetOrCreateLegend();
        legend2.Position = LegendPosition.TopRight;

        // Define chart data for second chart (Assists vs Player Name)
        IChartDataSource<string> xs2 = DataSources.FromStringCellRange(tableSheet, new CellRangeAddress(3, 6, 2, 2)); // Player Names
        IChartDataSource<double> ys2 = DataSources.FromNumericCellRange(tableSheet, new CellRangeAddress(3, 6, 5, 5)); // Assists

        IScatterChartData<string, double> scatterChartData2 = chart2.ChartDataFactory.CreateScatterChartData<string, double>();
        IScatterChartSeries<string, double> series2 = scatterChartData2.AddSeries(xs2, ys2);
        series2.SetTitle("Assists This Season");

        IChartAxis bottomAxis2 = chart2.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
        IValueAxis leftAxis2 = chart2.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
        leftAxis2.Crosses = AxisCrosses.AutoZero;
        chart2.Plot(scatterChartData2);

        // Save the Excel file
        using (var fileData = new FileStream("Football_Report_with_Charts.xlsx", FileMode.Create))
        {
            workbook.Write(fileData);
        }

        Console.WriteLine("Football report with multi-row header and multiple charts generated successfully!");
    }
}
