
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;

using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;
using System.IO;
using System.Diagnostics; // For opening file





namespace MyDesktopApp.Pages
{




    public partial class Page2 : Page
    {
        private XLWorkbook workbook;
        private string selectedValue;
        private IXLWorksheet worksheet;
         private string selectedMonth;
         private string selectedYear;

private double ToDouble(IXLCell cell)
{
    double result = 0;
    double.TryParse(cell.GetValue<string>(), out result);
    return result;
}




public Page2(XLWorkbook uploadedWorkbook, string selectedValue, string selectedMonth, string selectedYear)
{
    InitializeComponent();
    
    this.workbook = uploadedWorkbook;
    this.selectedValue = selectedValue;
this.selectedMonth =selectedMonth;
this.selectedYear = selectedYear;
    this.worksheet = workbook.Worksheet(1);
    Console.WriteLine(" Page2 with Selected Value: " + selectedValue + " and Date from: "+ workbook);
    LoadEmployeeNames();
}


private void LoadEmployeeNames()
{
    // Find the "Name of Employee" column
    var employeeColumn = worksheet
        .CellsUsed()
        .FirstOrDefault(c => c.GetString().Trim()
        .Equals("Name of Employee", StringComparison.OrdinalIgnoreCase));

    if (employeeColumn == null)
    {
        MessageBox.Show("Column 'Name of Employee' not found in Excel.");
        return;
    }

    int nameCol = employeeColumn.Address.ColumnNumber;

    List<string> employeeNames = new List<string>();

    // Loop through each row (skip header)
    foreach (var row in worksheet.RowsUsed().Skip(1))
    {
        // Read first 6 columns (A to F) – adjust count if needed
        bool allEmpty = true;
        int totalColumns = worksheet.ColumnsUsed().Count();

        for (int col = 1; col <=totalColumns; col++)   
        {
            string value = row.Cell(col).GetString().Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                allEmpty = false;
                break;
            }
        }

        // ❌ Skip row if first 6 columns are empty
        if (allEmpty)
            continue;

        // Get employee name
        string empName = row.Cell(nameCol).GetString().Trim();

        if (!string.IsNullOrWhiteSpace(empName)|| !string.IsNullOrEmpty(empName))
            employeeNames.Add(Clean(empName));
    }

    // Remove duplicates and bind
    EmployeeComboBox.ItemsSource = employeeNames
        .Distinct()
        .OrderBy(n => n)
        .ToList();
}










   private void EmployeeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Console.WriteLine("EmployeeComboBox_SelectionChanged triggered");
            if (EmployeeComboBox.SelectedItem == null)
                return;

            string selectedName = EmployeeComboBox.SelectedItem.ToString();
            Console.WriteLine("Selected Employee: " + selectedName);
            DisplayEmployeeDetails(selectedName);
        }




        private void DisplayEmployeeDetails(string employeeName)
        {
            DetailsPanel.Children.Clear();

            // Get header row (Row 5 in your sheet)
            var headers = worksheet.Row(5).CellsUsed().Select(c => c.GetString().Trim()).ToList();

            // Try to find the employee row (starting after header + category row)
            IXLRow employeeRow = null;
            for (int row = 7; row <= worksheet.LastRowUsed().RowNumber(); row++)
            {
                var cell = worksheet.Cell(row, 2);

                // Handle merged cells
                var mergedRange = cell.MergedRange();
                string name = mergedRange != null
                    ? Clean(mergedRange.FirstCell().GetString())
                    : Clean(cell.GetString());

                if (string.Equals(name, Clean(employeeName), StringComparison.OrdinalIgnoreCase))
                {
                    employeeRow = worksheet.Row(row);
                    break;
                }
            }

            if (employeeRow != null)
            {
                for (int i = 0; i < headers.Count; i++)
                {
                    string header = headers[i];
                    string value = employeeRow.Cell(i + 1).GetString();

                    TextBlock tb = new TextBlock
                    {
                        Text = $"{header}: {value}",
                        Margin = new Thickness(0, 5, 0, 5),
                        FontSize = 14
                    };

                    DetailsPanel.Children.Add(tb);
                }
            }
            else
            {
                DetailsPanel.Children.Add(new TextBlock { Text = "Employee details not found." });
            }
        }









        private void GoToPage1_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Navigate(new Page1());
        }

        /// <summary>
        /// Removes normal and non-breaking spaces from text.
        /// </summary>
        private string Clean(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            // Replace non-breaking spaces (char 160) with normal ones, then trim
            var cleaned = input.Replace('\u00A0', ' ').Trim();

            // Optionally remove all whitespace characters (safer for Excel oddities)
            cleaned = new string(cleaned.Where(c => !char.IsControl(c)).ToArray());

            return cleaned;
        }






// private void GeneratePDF_Click(object sender, RoutedEventArgs e)
// {
//     try
//     {
//         if (EmployeeComboBox.SelectedItem == null)
//         {
//             MessageBox.Show("Please select an employee first.");
//             return;
//         }

//         string employeeName = EmployeeComboBox.SelectedItem.ToString();
//         // string month = fromMonthYear.ToString("MMMM");
//         // string year = fromMonthYear.Year.ToString();
//         // string fromDate = fromMonthYear.ToString("dd.MM.yyyy");
//         // string toDate = toMonthYear.ToString("dd.MM.yyyy");

//         // Find employee row
//         IXLRow employeeRow = FindEmployeeRow(employeeName);
//         if (employeeRow == null)
//         {
//             MessageBox.Show("Employee details not found.");
//             return;
//         }

//         // Get salary data based on employee type
//         SalaryData salaryData = GetSalaryData(selectedValue, employeeRow);

//         // Generate PDF
//         GenerateSalarySlipPDF(employeeName,  salaryData);

//         MessageBox.Show("PDF generated successfully!");
//     }
//     catch (Exception ex)
//     {
//         MessageBox.Show("Error: " + ex.Message);
//     }
// }
private void GeneratePDF_Click(object sender, RoutedEventArgs e)
{
    try
    {
        if (EmployeeComboBox.SelectedItem == null)
        {
            MessageBox.Show("Please select an employee first.");
            return;
        }
        string selectedName = EmployeeComboBox.SelectedItem.ToString();

        // string employeeName = EmployeeComboBox.SelectedItem.ToString();

        // string selectedYear = YearComboBox.SelectedItem?.ToString();
        // string selectedMonth = MonthComboBox.SelectedItem?.ToString();

        if (selectedYear == null || selectedMonth == null)
        {
            MessageBox.Show("Select Year and Month first.");
            return;
        }

        // Find employee row
        IXLRow employeeRow = FindEmployeeRow(selectedName);
        if (employeeRow == null)
        {
            MessageBox.Show("Employee details not found.");
            return;
        }

        // Get salary data
        SalaryData salaryData = GetSalaryData(selectedValue, employeeRow);

        // Pass year and month to PDF generator
        GenerateSalarySlipPDF(selectedName, salaryData);

        MessageBox.Show("PDF generated successfully!"+ selectedYear + selectedMonth+ selectedName);
    }
    catch (Exception ex)
    {
        MessageBox.Show("Error: " + ex.Message);
    }
}

private IXLRow FindEmployeeRow(string employeeName)
{
    for (int row = 7; row <= worksheet.LastRowUsed().RowNumber(); row++)
    {
        var cell = worksheet.Cell(row, 2);
        string name = Clean(cell.GetString());
        
        if (string.Equals(name, Clean(employeeName), StringComparison.OrdinalIgnoreCase))
        {
            return worksheet.Row(row);
        }
    }
    return null;
}

private SalaryData GetSalaryData(string staffType, IXLRow employeeRow)
{
    var data = new SalaryData();

    // Common columns for all staff types
    data.BasicPay = ToDouble(employeeRow.Cell("H"));
    data.GrossSalary = ToDouble(employeeRow.Cell("I"));
    data.PfEmployee = ToDouble(employeeRow.Cell("J"));
    data.PfEmployer = ToDouble(employeeRow.Cell("K"));
    data.EsiEmployee = ToDouble(employeeRow.Cell("L"));
    data.EsiEmployer = ToDouble(employeeRow.Cell("M"));
    data.EmployeeClub = ToDouble(employeeRow.Cell("N"));

    // Staff type-specific mappings
    switch (staffType)
    {
        case "Canteen Staffs":
            data.OverTime = ToDouble(employeeRow.Cell("G"));
            data.OverTimeSalary = ToDouble(employeeRow.Cell("P"));
            data.MonthlySalary = ToDouble(employeeRow.Cell("O"));
            data.TakeHome = ToDouble(employeeRow.Cell("Q"));
            data.CTC = ToDouble(employeeRow.Cell("R"));
            data.ProfessionalTax = ToDouble(employeeRow.Cell("V"));
            data.TotalDeductions = ToDouble(employeeRow.Cell("N"));
            break;

        case "Social Investigators":
        case "Social investigators":
        case "social investigators":
            data.OverTime = ToDouble(employeeRow.Cell("F"));
            data.OverTimeSalary = ToDouble(employeeRow.Cell("P"));
            data.MonthlySalary = ToDouble(employeeRow.Cell("O"));
            data.TakeHome = ToDouble(employeeRow.Cell("R"));
            data.CTC = ToDouble(employeeRow.Cell("S"));
            data.ProfessionalTax = ToDouble(employeeRow.Cell("Q"));
            data.TotalDeductions = ToDouble(employeeRow.Cell("N"));
            break;

        case "Administrative Assistant":
            data.OverTime = ToDouble(employeeRow.Cell("F"));
            data.OverTimeSalary = ToDouble(employeeRow.Cell("P"));
            data.MonthlySalary = ToDouble(employeeRow.Cell("O"));
            data.TakeHome = ToDouble(employeeRow.Cell("R"));
            data.CTC = ToDouble(employeeRow.Cell("S"));
            data.ProfessionalTax = ToDouble(employeeRow.Cell("Q"));
            data.TotalDeductions = ToDouble(employeeRow.Cell("Z"));
            break;

        case "Electrician and Plumber":
            data.OverTime = ToDouble(employeeRow.Cell("F"));
            data.OverTimeSalary = ToDouble(employeeRow.Cell("P"));
            data.MonthlySalary = ToDouble(employeeRow.Cell("O"));
            data.TakeHome = ToDouble(employeeRow.Cell("R"));
            data.CTC = ToDouble(employeeRow.Cell("S"));
            data.ProfessionalTax = ToDouble(employeeRow.Cell("Q"));
            data.HasBasicPay = false;
            break;

        case "Ward Assistants":
            data.BasicPay = ToDouble(employeeRow.Cell("I"));
            data.GrossSalary = ToDouble(employeeRow.Cell("J"));
            data.OverTime = ToDouble(employeeRow.Cell("G"));
            data.MonthlySalary = ToDouble(employeeRow.Cell("P"));
            data.OverTimeSalary = ToDouble(employeeRow.Cell("Q"));
            data.TakeHome = ToDouble(employeeRow.Cell("S"));
            data.CTC = ToDouble(employeeRow.Cell("R"));
            data.PfEmployee = ToDouble(employeeRow.Cell("K"));
            data.PfEmployer = ToDouble(employeeRow.Cell("L"));
            data.EsiEmployee = ToDouble(employeeRow.Cell("M"));
            data.EsiEmployer = ToDouble(employeeRow.Cell("N"));
            data.EmployeeClub = ToDouble(employeeRow.Cell("O"));
            data.ProfessionalTax = ToDouble(employeeRow.Cell("R"));
            data.TotalDeductions = ToDouble(employeeRow.Cell("Z"));
            data.Others = ToDouble(employeeRow.Cell("U"));
            break;
    }

    return data;
}

private void GenerateSalarySlipPDF(string employeeName,  
     SalaryData salary)
{
    PdfDocument document = new PdfDocument();
    PdfPage page = document.AddPage();
    XGraphics gfx = XGraphics.FromPdfPage(page);

    // Define fonts
    XFont titleFont = new XFont("Arial", 18, XFontStyle.Bold);
    XFont boldFont = new XFont("Arial", 12, XFontStyle.Bold);
    XFont normalFont = new XFont("Arial", 11);
    XFont smallFont = new XFont("Arial", 10);

    double y = 30;
    double pageWidth = page.Width;
    double pageHeight = page.Height;
    double margin = 40;
    double contentWidth = pageWidth - (2 * margin);

    // ===== HEADER SECTION =====
    gfx.DrawString("SALARY SLIP", titleFont, XBrushes.Black,
        new XRect(0, y, pageWidth, 30), XStringFormats.TopCenter);
    y += 35;
    gfx.DrawString($"{selectedMonth?.ToUpper()} ", boldFont, XBrushes.Black,
        new XRect(0, y, pageWidth, 20), XStringFormats.TopCenter);
    y += 25;
    gfx.DrawString($"{selectedYear?.ToUpper()}", boldFont, XBrushes.Black,
        new XRect(0, y, pageWidth, 20), XStringFormats.TopCenter);
    y += 25;
  gfx.DrawLine(XPens.Black, margin, y, pageWidth - margin, y);
    y += 15;


    // Employee info
    gfx.DrawString($"Employee Name: {employeeName}", normalFont, XBrushes.Black,
        new XRect(margin, y, contentWidth / 2, 18), XStringFormats.TopLeft);
    // gfx.DrawString($"Period: {fromDate} to {toDate}", normalFont, XBrushes.Black,
    //     new XRect(margin + contentWidth / 2, y, contentWidth / 2, 18), XStringFormats.TopLeft);
    

        gfx.DrawString( $"Designation: {selectedValue}", normalFont,XBrushes.Black,
    new XRect(margin, y, contentWidth, 20), XStringFormats.TopRight);

    // Separator line
   y += 25;
    // ===== TABLE SECTION =====
    double tableX = margin;
    double tableWidth = contentWidth;
    double col1Width = tableWidth * 0.35;
    double col2Width = tableWidth * 0.15;
    double col3Width = tableWidth * 0.35;
    double col4Width = tableWidth * 0.15;
    double rowHeight = 25;

    // Table header
    DrawTableHeader(gfx, tableX, y, col1Width, col2Width, col3Width, col4Width, rowHeight, boldFont);
    y += rowHeight;

    // Table content
    y = DrawSalaryTable(gfx, tableX, y, col1Width, col2Width, col3Width, col4Width, 
        rowHeight, normalFont, salary);

    // Separator line
    y += 10;
    gfx.DrawLine(XPens.Black, margin, y, pageWidth - margin, y);
    y += 15;

    // ===== SUMMARY SECTION =====
    // 
    // double totalDeductions = salary.PfEmployee + salary.PfEmployer + salary.EsiEmployee + 
        // salary.EsiEmployer + salary.EmployeeClub + salary.ProfessionalTax;

    gfx.DrawString($"Total Monthly Salary: ₹ {salary.MonthlySalary:0}", boldFont, XBrushes.Black,
        new XRect(margin, y, contentWidth / 2, 20), XStringFormats.TopLeft);
    gfx.DrawString($"Total Deductions: ₹ {salary.TotalDeductions:0}", boldFont, XBrushes.Black,
        new XRect(margin + contentWidth / 2, y, contentWidth / 2, 20), XStringFormats.TopLeft);
    y += 25;

    // Net salary box
    XPen borderPen = new XPen(XColor.FromArgb(0, 0, 0), 2);
    gfx.DrawRectangle(XBrushes.LightGray, margin, y, contentWidth, 35);
    gfx.DrawRectangle(borderPen, margin, y, contentWidth, 35);
    
    gfx.DrawString($"Net Salary Released: ₹ {salary.TakeHome:0.00}", 
        new XFont("Arial", 14, XFontStyle.Bold), XBrushes.Black,
        new XRect(margin + 10, y + 5, contentWidth - 20, 25), XStringFormats.CenterLeft);
    
    y += 65;
    gfx.DrawString("Manager".ToUpper(), boldFont, XBrushes.Black,
        new XRect(margin + contentWidth / 2, y, contentWidth / 2, 20), XStringFormats.TopRight);
    y += 15;
    gfx.DrawString("Human Arm", boldFont, XBrushes.Black,
        new XRect(margin + contentWidth / 2, y, contentWidth / 2, 20), XStringFormats.TopRight);

    // Save PDF
//     // --------------------------
// 1. Read selected year & month
// --------------------------
// string selectedYear = YearComboBox.SelectedItem?.ToString();
// string selectedMonth = MonthComboBox.SelectedItem?.ToString();

if (string.IsNullOrEmpty(selectedYear) || string.IsNullOrEmpty(selectedMonth))
{
    MessageBox.Show("Please select both Year and Month.");
    return;
}

// --------------------------
// 2. Base folder in Documents
// --------------------------
string baseFolder = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
    "Payslip"
);

// --------------------------
// 3. YEAR folder
// --------------------------
string roleFolder = Path.Combine(baseFolder, selectedValue);
Directory.CreateDirectory(roleFolder);
string yearFolder = Path.Combine(roleFolder, selectedYear);
Directory.CreateDirectory(yearFolder);

// --------------------------
// 4. MONTH folder inside year
// --------------------------
string monthFolder = Path.Combine(yearFolder, selectedMonth);
Directory.CreateDirectory(monthFolder);

// --------------------------
// 5. Clean Employee Name (safe filename)
// --------------------------
string safeName = string.Join("_", employeeName.Split(Path.GetInvalidFileNameChars()));

// --------------------------
// 6. Create final PDF path
// --------------------------
string filePath = Path.Combine(
    monthFolder,
    $"{safeName}_SalarySlip_{selectedMonth}_{selectedYear}.pdf"
);
if (File.Exists(filePath))
{
    MessageBox.Show(
        "A salary slip for this employee already exists for the selected month and year.",
        "Duplicate File",
        MessageBoxButton.OK,
        MessageBoxImage.Warning
    );

    return; // 🚫 Stop saving (important)
}
// --------------------------
// 7. Save the generated PDF
// --------------------------
document.Save(filePath);

// --------------------------
// 8. Open the PDF after saving
// --------------------------
Process.Start(new ProcessStartInfo
{
    FileName = filePath,
    UseShellExecute = true
});

MessageBox.Show($"PDF saved successfully at:\n{filePath}");
}
private void DrawTableHeader(XGraphics gfx, double x, double y, double col1, double col2, 
    double col3, double col4, double rowH, XFont font)
{
    XBrush headerBrush = XBrushes.LightGray;
    XPen borderPen = XPens.Black;

    // Draw header cells
    gfx.DrawRectangle(headerBrush, x, y, col1, rowH);
    gfx.DrawRectangle(borderPen, x, y, col1, rowH);
    gfx.DrawString("Earnings", font, XBrushes.Black, new XRect(x + 5, y + 4, col1 - 10, rowH), 
        XStringFormats.CenterLeft);

    gfx.DrawRectangle(headerBrush, x + col1, y, col2, rowH);
    gfx.DrawRectangle(borderPen, x + col1, y, col2, rowH);
    gfx.DrawString("Amount", font, XBrushes.Black, new XRect(x + col1 + 5, y + 4, col2 - 10, rowH), 
        XStringFormats.CenterRight);

    gfx.DrawRectangle(headerBrush, x + col1 + col2, y, col3, rowH);
    gfx.DrawRectangle(borderPen, x + col1 + col2, y, col3, rowH);
    gfx.DrawString("Contributions", font, XBrushes.Black, 
        new XRect(x + col1 + col2 + 5, y + 4, col3 - 10, rowH), XStringFormats.CenterLeft);

    gfx.DrawRectangle(headerBrush, x + col1 + col2 + col3, y, col4, rowH);
    gfx.DrawRectangle(borderPen, x + col1 + col2 + col3, y, col4, rowH);
    gfx.DrawString("Amount", font, XBrushes.Black, 
        new XRect(x + col1 + col2 + col3 + 5, y + 4, col4 - 10, rowH), XStringFormats.CenterRight);
}

private double DrawSalaryTable(XGraphics gfx, double x, double y, double col1, double col2, 
    double col3, double col4, double rowH, XFont font, SalaryData salary)
{
    XPen borderPen = XPens.Black;

    void DrawRow(string label1, double val1, string label2, double val2)
    {
        gfx.DrawRectangle(borderPen, x, y, col1, rowH);
        gfx.DrawString(label1, font, XBrushes.Black, new XRect(x + 5, y + 4, col1 - 10, rowH), 
            XStringFormats.TopLeft);

        gfx.DrawRectangle(borderPen, x + col1, y, col2, rowH);
        gfx.DrawString(val1.ToString("0.00"), font, XBrushes.Black, 
            new XRect(x + col1 + 5, y + 4, col2 - 10, rowH), XStringFormats.TopRight);

        gfx.DrawRectangle(borderPen, x + col1 + col2, y, col3, rowH);
        gfx.DrawString(label2, font, XBrushes.Black, 
            new XRect(x + col1 + col2 + 5, y + 4, col3 - 10, rowH), XStringFormats.TopLeft);

        gfx.DrawRectangle(borderPen, x + col1 + col2 + col3, y, col4, rowH);
        gfx.DrawString(val2.ToString("0.00"), font, XBrushes.Black, 
            new XRect(x + col1 + col2 + col3 + 5, y + 4, col4 - 10, rowH), XStringFormats.TopRight);

        y += rowH;
    }

    // Draw rows
    DrawRow("Basic Pay", salary.BasicPay, "PF Employee", salary.PfEmployee);
    DrawRow("Gross Salary", salary.GrossSalary, "PF Employer", salary.PfEmployer);
    DrawRow("OverTime", salary.OverTime, "ESI Employee", salary.EsiEmployee);
    DrawRow("OverTime Salary", salary.OverTimeSalary, "ESI Employer", salary.EsiEmployer);
    DrawRow("Monthly Salary", salary.MonthlySalary, "Employee Club", salary.EmployeeClub);
    DrawRow("Others", salary.Others, "Professional Tax", salary.ProfessionalTax);
    DrawRow("CTC", salary.CTC, "Total Deductions", salary.TotalDeductions);

    return y;
}

// Helper class to hold salary data
public class SalaryData
{
    public double BasicPay { get; set; }
    public double GrossSalary { get; set; }
    public double OverTime { get; set; }
    public double OverTimeSalary { get; set; }
    public double MonthlySalary { get; set; }
    public double TakeHome { get; set; }
    public double CTC { get; set; }
    public double PfEmployee { get; set; }
    public double PfEmployer { get; set; }
    public double EsiEmployee { get; set; }
    public double EsiEmployer { get; set; }
    public double EmployeeClub { get; set; }
    public double ProfessionalTax { get; set; }
    public double TotalDeductions { get; set; }
    public double Others { get; set; }
    public bool HasBasicPay { get; set; } = true;
}
    }
}


