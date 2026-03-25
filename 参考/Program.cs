using System.Data;
using Aspose.Cells;

const int TableHeaderRowIndex = 6;
const int MarkerRowIndex = 7;

TryLoadLicense();

var templateDirectory = Path.Combine(AppContext.BaseDirectory, "template");
var outputDirectory = Path.Combine(AppContext.BaseDirectory, "output");

Directory.CreateDirectory(templateDirectory);
Directory.CreateDirectory(outputDirectory);

var templatePath = Path.Combine(templateDirectory, "sales-order-template.xlsx");
var outputPath = Path.Combine(outputDirectory, $"sales-order-{DateTime.Now:yyyyMMddHHmmss}.xlsx");
var monthlyTemplatePath = Path.Combine(templateDirectory, "monthly-amount-template.xlsx");
var monthlyOutputPath = Path.Combine(outputDirectory, $"monthly-amount-{DateTime.Now:yyyyMMddHHmmss}.xlsx");
var matrixTemplatePath = Path.Combine(templateDirectory, "product-month-matrix-template.xlsx");
var matrixOutputPath = Path.Combine(outputDirectory, $"product-month-matrix-{DateTime.Now:yyyyMMddHHmmss}.xlsx");
var report = BuildSampleReport();
var monthlyReport = BuildMonthlyReport();
var rawTransactions = BuildRawTransactions();
var matrixReport = BuildProductMonthMatrixReportFromTransactions(rawTransactions);
var monthColumns = BuildMonthColumns(matrixReport.MonthHeaders);
var matrixTable = BuildProductMonthMatrixTable(matrixReport, monthColumns);

CreateTemplate(templatePath);
GenerateWorkbook(templatePath, outputPath, report);
CreateMonthlyTemplate(monthlyTemplatePath);
GenerateMonthlyWorkbook(monthlyTemplatePath, monthlyOutputPath, monthlyReport);
CreateProductMonthMatrixTemplate(matrixTemplatePath, monthColumns);
GenerateProductMonthMatrixWorkbook(matrixTemplatePath, matrixOutputPath, matrixTable);
PrintMatrixPreview(matrixOutputPath);

Console.WriteLine($"Template file: {templatePath}");
Console.WriteLine($"Output file: {outputPath}");
Console.WriteLine($"Monthly template file: {monthlyTemplatePath}");
Console.WriteLine($"Monthly output file: {monthlyOutputPath}");
Console.WriteLine($"Matrix template file: {matrixTemplatePath}");
Console.WriteLine($"Matrix output file: {matrixOutputPath}");

static void GenerateWorkbook(string templatePath, string outputPath, SalesReport report)
{
    var designer = new WorkbookDesigner
    {
        Workbook = new Workbook(templatePath)
    };

    designer.SetDataSource("Report", report);
    designer.SetDataSource("Items", report.Items);
    designer.Process();

    var workbook = designer.Workbook;
    var worksheet = workbook.Worksheets[0];
    var totalRowIndex = MarkerRowIndex + report.Items.Count;

    worksheet.Cells[$"D{totalRowIndex + 1}"].PutValue("Total");
    worksheet.Cells[$"E{totalRowIndex + 1}"].PutValue(report.Items.Sum(item => item.Amount));

    ApplyTotalRowStyle(worksheet, totalRowIndex);

    worksheet.AutoFitColumns();
    worksheet.AutoFitRows();

    workbook.Save(outputPath);
}

static void CreateTemplate(string templatePath)
{
    var workbook = new Workbook();
    var worksheet = workbook.Worksheets[0];
    worksheet.Name = "SalesOrder";

    worksheet.Cells.Merge(0, 0, 1, 5);
    worksheet.Cells["A1"].PutValue("Sales Order Report");
    worksheet.Cells["A3"].PutValue("Order No:");
    worksheet.Cells["B3"].PutValue("&=$Report.ReportNo");
    worksheet.Cells["D3"].PutValue("Created By:");
    worksheet.Cells["E3"].PutValue("&=$Report.CreatedBy");
    worksheet.Cells["A4"].PutValue("Customer:");
    worksheet.Cells["B4"].PutValue("&=$Report.CustomerName");
    worksheet.Cells["D4"].PutValue("Date:");
    worksheet.Cells["E4"].PutValue("&=$Report.ReportDate");
    worksheet.Cells["A6"].PutValue("Note:");
    worksheet.Cells["B6"].PutValue("Row 8 is the Smart Marker detail row and will expand automatically.");

    var headers = new[] { "No.", "Product", "Qty", "Unit Price", "Amount" };
    for (var column = 0; column < headers.Length; column++)
    {
        worksheet.Cells[TableHeaderRowIndex, column].PutValue(headers[column]);
    }

    worksheet.Cells["A8"].PutValue("&=Items.LineNo(copystyle)");
    worksheet.Cells["B8"].PutValue("&=Items.ProductName(copystyle)");
    worksheet.Cells["C8"].PutValue("&=Items.Quantity(numeric,copystyle)");
    worksheet.Cells["D8"].PutValue("&=Items.UnitPrice(numeric,copystyle)");
    worksheet.Cells["E8"].PutValue("&=Items.Amount(numeric,copystyle)");

    ApplyLayout(worksheet);
    workbook.Save(templatePath);
}

static void GenerateMonthlyWorkbook(string templatePath, string outputPath, MonthlyReport report)
{
    var designer = new WorkbookDesigner
    {
        Workbook = new Workbook(templatePath)
    };

    designer.SetDataSource("Report", report);
    designer.SetDataSource("MonthData", report.Months);
    designer.Process();

    var workbook = designer.Workbook;
    var worksheet = workbook.Worksheets[0];

    worksheet.AutoFitColumns();
    worksheet.AutoFitRows();

    workbook.Save(outputPath);
}

static void CreateMonthlyTemplate(string templatePath)
{
    var workbook = new Workbook();
    var worksheet = workbook.Worksheets[0];
    worksheet.Name = "MonthlyAmount";

    worksheet.Cells.Merge(0, 0, 1, 8);
    worksheet.Cells["A1"].PutValue("Monthly Amount Report");
    worksheet.Cells["A3"].PutValue("Customer:");
    worksheet.Cells["B3"].PutValue("&=$Report.CustomerName");
    worksheet.Cells["A4"].PutValue("Scenario:");
    worksheet.Cells["B4"].PutValue("Amount column expands to the right by month.");

    worksheet.Cells["A6"].PutValue("Metric");
    worksheet.Cells["B6"].PutValue("&=MonthData.Month(horizontal,shift,copystyle)");
    worksheet.Cells["A7"].PutValue("Amount");
    worksheet.Cells["B7"].PutValue("&=MonthData.Amount(horizontal,shift,copystyle,numeric)");

    worksheet.Cells["A9"].PutValue("Template note:");
    worksheet.Cells["B9"].PutValue("B6 and B7 are horizontal Smart Markers. New months create new columns automatically.");

    ApplyMonthlyLayout(worksheet);
    workbook.Save(templatePath);
}

static void GenerateProductMonthMatrixWorkbook(string templatePath, string outputPath, DataTable matrixTable)
{
    var designer = new WorkbookDesigner
    {
        Workbook = new Workbook(templatePath)
    };

    designer.SetDataSource(matrixTable);
    designer.Process();

    var workbook = designer.Workbook;
    var worksheet = workbook.Worksheets[0];

    worksheet.AutoFitColumns();
    worksheet.AutoFitRows();

    workbook.Save(outputPath);
}

static void CreateProductMonthMatrixTemplate(string templatePath, IReadOnlyList<MonthColumn> monthColumns)
{
    var workbook = new Workbook();
    var worksheet = workbook.Worksheets[0];
    worksheet.Name = "ProductMonthMatrix";

    worksheet.Cells.Merge(0, 0, 1, 10);
    worksheet.Cells["A1"].PutValue("Product Amount Matrix");
    worksheet.Cells["A3"].PutValue("Customer:");
    worksheet.Cells["B3"].PutValue("Shanghai Example Technology Co., Ltd.");
    worksheet.Cells["A4"].PutValue("Scenario:");
    worksheet.Cells["B4"].PutValue("Dynamic month columns are generated in the template, then filled by Smart Markers.");

    worksheet.Cells["A6"].PutValue("Product");
    worksheet.Cells["A7"].PutValue("&=MatrixTable.ProductName(copystyle)");

    for (var i = 0; i < monthColumns.Count; i++)
    {
        var columnIndex = i + 1;
        worksheet.Cells[5, columnIndex].PutValue(monthColumns[i].HeaderText);
        worksheet.Cells[6, columnIndex].PutValue($"&=MatrixTable.{monthColumns[i].FieldName}(numeric,copystyle)");
    }

    worksheet.Cells["A9"].PutValue("Template note:");
    worksheet.Cells["B9"].PutValue("This is the reliable approach for a 2D matrix: dynamic template columns plus Smart Marker row expansion.");

    ApplyProductMonthMatrixLayout(worksheet, monthColumns.Count);
    workbook.Save(templatePath);
}

static void PrintMatrixPreview(string workbookPath)
{
    var workbook = new Workbook(workbookPath);
    var worksheet = workbook.Worksheets[0];

    Console.WriteLine("Matrix preview:");

    for (var row = 5; row <= 7; row++)
    {
        var values = new List<string>();
        for (var column = 0; column <= 5; column++)
        {
            values.Add(worksheet.Cells[row, column].StringValue);
        }

        Console.WriteLine(string.Join(" | ", values));
    }
}

static void ApplyLayout(Worksheet worksheet)
{
    worksheet.Cells.SetColumnWidth(0, 10);
    worksheet.Cells.SetColumnWidth(1, 26);
    worksheet.Cells.SetColumnWidth(2, 12);
    worksheet.Cells.SetColumnWidth(3, 14);
    worksheet.Cells.SetColumnWidth(4, 16);

    var titleStyle = worksheet.Cells["A1"].GetStyle();
    titleStyle.Font.Size = 16;
    titleStyle.Font.IsBold = true;
    titleStyle.HorizontalAlignment = TextAlignmentType.Center;
    titleStyle.VerticalAlignment = TextAlignmentType.Center;
    worksheet.Cells["A1"].SetStyle(titleStyle);

    ApplyLabelStyle(worksheet.Cells["A3"]);
    ApplyLabelStyle(worksheet.Cells["D3"]);
    ApplyLabelStyle(worksheet.Cells["A4"]);
    ApplyLabelStyle(worksheet.Cells["D4"]);
    ApplyLabelStyle(worksheet.Cells["A6"]);

    var dateStyle = worksheet.Cells["E4"].GetStyle();
    dateStyle.Custom = "yyyy-mm-dd";
    worksheet.Cells["E4"].SetStyle(dateStyle);

    for (var column = 0; column < 5; column++)
    {
        var headerCell = worksheet.Cells[TableHeaderRowIndex, column];
        var headerStyle = headerCell.GetStyle();
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.ForegroundColor = System.Drawing.Color.FromArgb(31, 78, 121);
        headerStyle.Font.Color = System.Drawing.Color.White;
        headerStyle.Font.IsBold = true;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        headerStyle.VerticalAlignment = TextAlignmentType.Center;
        SetBorder(headerStyle);
        headerCell.SetStyle(headerStyle);
    }

    for (var column = 0; column < 5; column++)
    {
        var markerCell = worksheet.Cells[MarkerRowIndex, column];
        var markerStyle = markerCell.GetStyle();
        markerStyle.VerticalAlignment = TextAlignmentType.Center;
        SetBorder(markerStyle);

        if (column == 2)
        {
            markerStyle.HorizontalAlignment = TextAlignmentType.Right;
            markerStyle.Custom = "0";
        }
        else if (column >= 3)
        {
            markerStyle.HorizontalAlignment = TextAlignmentType.Right;
            markerStyle.Custom = "#,##0.00";
        }

        markerCell.SetStyle(markerStyle);
    }

    worksheet.FreezePanes(7, 0, 7, 0);
}

static void ApplyMonthlyLayout(Worksheet worksheet)
{
    worksheet.Cells.SetColumnWidth(0, 16);
    worksheet.Cells.SetColumnWidth(1, 14);

    var titleStyle = worksheet.Cells["A1"].GetStyle();
    titleStyle.Font.Size = 16;
    titleStyle.Font.IsBold = true;
    titleStyle.HorizontalAlignment = TextAlignmentType.Center;
    worksheet.Cells["A1"].SetStyle(titleStyle);

    ApplyLabelStyle(worksheet.Cells["A3"]);
    ApplyLabelStyle(worksheet.Cells["A4"]);
    ApplyLabelStyle(worksheet.Cells["A9"]);

    for (var row = 5; row <= 6; row++)
    {
        var labelCell = worksheet.Cells[row, 0];
        var labelStyle = labelCell.GetStyle();
        labelStyle.Pattern = BackgroundType.Solid;
        labelStyle.ForegroundColor = System.Drawing.Color.FromArgb(217, 225, 242);
        labelStyle.Font.IsBold = true;
        labelStyle.HorizontalAlignment = TextAlignmentType.Center;
        SetBorder(labelStyle);
        labelCell.SetStyle(labelStyle);

        var markerCell = worksheet.Cells[row, 1];
        var markerStyle = markerCell.GetStyle();
        markerStyle.HorizontalAlignment = row == 5 ? TextAlignmentType.Center : TextAlignmentType.Right;
        markerStyle.Custom = row == 5 ? "@" : "#,##0.00";
        SetBorder(markerStyle);
        markerCell.SetStyle(markerStyle);
    }
}

static void ApplyProductMonthMatrixLayout(Worksheet worksheet, int monthCount)
{
    worksheet.Cells.SetColumnWidth(0, 22);
    worksheet.Cells.SetColumnWidth(1, 14);

    var titleStyle = worksheet.Cells["A1"].GetStyle();
    titleStyle.Font.Size = 16;
    titleStyle.Font.IsBold = true;
    titleStyle.HorizontalAlignment = TextAlignmentType.Center;
    worksheet.Cells["A1"].SetStyle(titleStyle);

    ApplyLabelStyle(worksheet.Cells["A3"]);
    ApplyLabelStyle(worksheet.Cells["A4"]);
    ApplyLabelStyle(worksheet.Cells["A9"]);

    var productHeader = worksheet.Cells["A6"];
    var productHeaderStyle = productHeader.GetStyle();
    productHeaderStyle.Pattern = BackgroundType.Solid;
    productHeaderStyle.ForegroundColor = System.Drawing.Color.FromArgb(31, 78, 121);
    productHeaderStyle.Font.Color = System.Drawing.Color.White;
    productHeaderStyle.Font.IsBold = true;
    productHeaderStyle.HorizontalAlignment = TextAlignmentType.Center;
    SetBorder(productHeaderStyle);
    productHeader.SetStyle(productHeaderStyle);

    var productCell = worksheet.Cells["A7"];
    var productCellStyle = productCell.GetStyle();
    SetBorder(productCellStyle);
    productCell.SetStyle(productCellStyle);

    for (var column = 1; column <= monthCount; column++)
    {
        var monthHeader = worksheet.Cells[5, column];
        var monthHeaderStyle = monthHeader.GetStyle();
        monthHeaderStyle.Pattern = BackgroundType.Solid;
        monthHeaderStyle.ForegroundColor = System.Drawing.Color.FromArgb(31, 78, 121);
        monthHeaderStyle.Font.Color = System.Drawing.Color.White;
        monthHeaderStyle.Font.IsBold = true;
        monthHeaderStyle.HorizontalAlignment = TextAlignmentType.Center;
        SetBorder(monthHeaderStyle);
        monthHeader.SetStyle(monthHeaderStyle);

        var amountCell = worksheet.Cells[6, column];
        var amountCellStyle = amountCell.GetStyle();
        amountCellStyle.HorizontalAlignment = TextAlignmentType.Right;
        amountCellStyle.Custom = "#,##0.00";
        SetBorder(amountCellStyle);
        amountCell.SetStyle(amountCellStyle);
    }
}

static void ApplyTotalRowStyle(Worksheet worksheet, int totalRowIndex)
{
    var labelCell = worksheet.Cells[totalRowIndex, 3];
    var amountCell = worksheet.Cells[totalRowIndex, 4];

    var labelStyle = labelCell.GetStyle();
    labelStyle.Font.IsBold = true;
    labelStyle.HorizontalAlignment = TextAlignmentType.Right;
    SetBorder(labelStyle);
    labelCell.SetStyle(labelStyle);

    var amountStyle = amountCell.GetStyle();
    amountStyle.Font.IsBold = true;
    amountStyle.HorizontalAlignment = TextAlignmentType.Right;
    amountStyle.Custom = "#,##0.00";
    SetBorder(amountStyle);
    amountCell.SetStyle(amountStyle);
}

static void ApplyLabelStyle(Cell cell)
{
    var style = cell.GetStyle();
    style.Font.IsBold = true;
    cell.SetStyle(style);
}

static void SetBorder(Style style)
{
    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
}

static SalesReport BuildSampleReport()
{
    var items = new List<SalesItem>
    {
        new() { LineNo = 1, ProductName = "27-inch Monitor", Quantity = 3, UnitPrice = 1399.00m },
        new() { LineNo = 2, ProductName = "Wireless Keyboard", Quantity = 5, UnitPrice = 199.00m },
        new() { LineNo = 3, ProductName = "USB-C Dock", Quantity = 2, UnitPrice = 329.50m },
        new() { LineNo = 4, ProductName = "Office Chair", Quantity = 1, UnitPrice = 899.00m }
    };

    return new SalesReport
    {
        ReportNo = "SO-20260325-001",
        CustomerName = "Shanghai Example Technology Co., Ltd.",
        ReportDate = new DateTime(2026, 3, 25),
        CreatedBy = "Aspose.Cells Demo",
        Items = items
    };
}

static MonthlyReport BuildMonthlyReport()
{
    return new MonthlyReport
    {
        CustomerName = "Shanghai Example Technology Co., Ltd.",
        Months = new List<MonthAmount>
        {
            new() { Month = "2026-01", Amount = 1200m },
            new() { Month = "2026-02", Amount = 1560m },
            new() { Month = "2026-03", Amount = 1410m },
            new() { Month = "2026-04", Amount = 1780m },
            new() { Month = "2026-05", Amount = 1990m }
        }
    };
}

static ProductMonthMatrixReport BuildProductMonthMatrixReportFromTransactions(IReadOnlyList<ProductMonthTransaction> transactions)
{
    var monthHeaders = transactions
        .Select(transaction => transaction.Month)
        .Distinct(StringComparer.Ordinal)
        .OrderBy(month => month, StringComparer.Ordinal)
        .ToList();

    var rows = transactions
        .GroupBy(transaction => transaction.ProductName, StringComparer.Ordinal)
        .OrderBy(group => group.Key, StringComparer.Ordinal)
        .Select(group =>
        {
            var amountLookup = group.ToDictionary(item => item.Month, item => item.Amount, StringComparer.Ordinal);
            var amounts = monthHeaders
                .Select(month => amountLookup.TryGetValue(month, out var amount) ? amount : 0m)
                .ToArray();

            return new ProductMonthMatrixRow
            {
                ProductName = group.Key,
                Amounts = amounts
            };
        })
        .ToList();

    return new ProductMonthMatrixReport
    {
        CustomerName = "Shanghai Example Technology Co., Ltd.",
        MonthHeaders = monthHeaders,
        Rows = rows
    };
}

static List<ProductMonthTransaction> BuildRawTransactions()
{
    return new List<ProductMonthTransaction>
    {
        new() { ProductName = "Monitor", Month = "2026-01", Amount = 1200m },
        new() { ProductName = "Monitor", Month = "2026-02", Amount = 1560m },
        new() { ProductName = "Monitor", Month = "2026-03", Amount = 1410m },
        new() { ProductName = "Monitor", Month = "2026-04", Amount = 1780m },
        new() { ProductName = "Monitor", Month = "2026-05", Amount = 1990m },
        new() { ProductName = "Keyboard", Month = "2026-01", Amount = 300m },
        new() { ProductName = "Keyboard", Month = "2026-02", Amount = 420m },
        new() { ProductName = "Keyboard", Month = "2026-03", Amount = 380m },
        new() { ProductName = "Keyboard", Month = "2026-04", Amount = 450m },
        new() { ProductName = "Keyboard", Month = "2026-05", Amount = 470m },
        new() { ProductName = "Dock", Month = "2026-01", Amount = 980m },
        new() { ProductName = "Dock", Month = "2026-02", Amount = 1100m },
        new() { ProductName = "Dock", Month = "2026-03", Amount = 1050m },
        new() { ProductName = "Dock", Month = "2026-04", Amount = 1160m },
        new() { ProductName = "Dock", Month = "2026-05", Amount = 1230m }
    };
}

static List<MonthColumn> BuildMonthColumns(IReadOnlyList<string> monthHeaders)
{
    var columns = new List<MonthColumn>();

    for (var i = 0; i < monthHeaders.Count; i++)
    {
        columns.Add(new MonthColumn
        {
            HeaderText = monthHeaders[i],
            FieldName = $"M{i + 1:00}"
        });
    }

    return columns;
}

static DataTable BuildProductMonthMatrixTable(ProductMonthMatrixReport report, IReadOnlyList<MonthColumn> monthColumns)
{
    var table = new DataTable("MatrixTable");
    table.Columns.Add("ProductName", typeof(string));

    foreach (var monthColumn in monthColumns)
    {
        table.Columns.Add(monthColumn.FieldName, typeof(decimal));
    }

    foreach (var row in report.Rows)
    {
        var dataRow = table.NewRow();
        dataRow["ProductName"] = row.ProductName;

        for (var i = 0; i < monthColumns.Count; i++)
        {
            dataRow[monthColumns[i].FieldName] = row.Amounts[i];
        }

        table.Rows.Add(dataRow);
    }

    return table;
}

static void TryLoadLicense()
{
    var candidatePaths = new[]
    {
        Path.Combine(AppContext.BaseDirectory, "Aspose.Cells.lic"),
        Path.Combine(Directory.GetCurrentDirectory(), "Aspose.Cells.lic")
    };

    foreach (var path in candidatePaths)
    {
        if (!File.Exists(path))
        {
            continue;
        }

        var license = new License();
        license.SetLicense(path);
        Console.WriteLine($"Loaded license: {path}");
        return;
    }

    Console.WriteLine("Aspose.Cells license was not found. The output will use evaluation mode.");
}

public sealed class SalesReport
{
    public string ReportNo { get; init; } = string.Empty;

    public string CustomerName { get; init; } = string.Empty;

    public DateTime ReportDate { get; init; }

    public string CreatedBy { get; init; } = string.Empty;

    public List<SalesItem> Items { get; init; } = new();
}

public sealed class SalesItem
{
    public int LineNo { get; init; }

    public string ProductName { get; init; } = string.Empty;

    public int Quantity { get; init; }

    public decimal UnitPrice { get; init; }

    public decimal Amount => Quantity * UnitPrice;
}

public sealed class MonthlyReport
{
    public string CustomerName { get; init; } = string.Empty;

    public List<MonthAmount> Months { get; init; } = new();
}

public sealed class MonthAmount
{
    public string Month { get; init; } = string.Empty;

    public decimal Amount { get; init; }
}

public sealed class ProductMonthMatrixReport
{
    public string CustomerName { get; init; } = string.Empty;

    public List<string> MonthHeaders { get; init; } = new();

    public List<ProductMonthMatrixRow> Rows { get; init; } = new();
}

public sealed class ProductMonthMatrixRow
{
    public string ProductName { get; init; } = string.Empty;

    public decimal[] Amounts { get; init; } = Array.Empty<decimal>();
}

public sealed class ProductMonthTransaction
{
    public string ProductName { get; init; } = string.Empty;

    public string Month { get; init; } = string.Empty;

    public decimal Amount { get; init; }
}

public sealed class MonthColumn
{
    public string HeaderText { get; init; } = string.Empty;

    public string FieldName { get; init; } = string.Empty;
}
