using System.Windows;
using System.Windows.Controls;
using System.IO;
using Microsoft.Win32;
using ClosedXML.Excel; //to read the excel file



namespace MyDesktopApp.Pages
{

    public partial class Page1 : Page
    {

        private string cellA2Value = string.Empty;
        private XLWorkbook uploadedWorkbook; // store the workbook  
 private XLWorkbook workbook;
        private string selectedValue;
        private IXLWorksheet worksheet;
         private string selectedMonth;
         private string selectedYear;




public Page1()
{
    InitializeComponent();
    
//     this.workbook = uploadedWorkbook;
//     this.selectedValue = selectedValue;
// this.selectedMonth =selectedMonth;
// this.selectedYear = selectedYear;
//     this.worksheet = workbook.Worksheet(1);
                OptionsComboBox.SelectionChanged += OptionsComboBox_SelectionChanged;
  RestoreSavedValues();
}
private void RestoreSavedValues()
{
    if (App.AppMemory.UploadedWorkbook != null) 
    {
        uploadedWorkbook = App.AppMemory.UploadedWorkbook;
        worksheet = uploadedWorkbook.Worksheet(1);

        // Assign previously selected values
        SelectComboItem(OptionsComboBox, App.AppMemory.SelectedCategory);
        SelectComboItem(MonthComboBox,App. AppMemory.SelectedMonth);
        SelectComboItem(YearComboBox,App. AppMemory.SelectedYear);


        // Recreate Excel text
        cellA2Value = worksheet.Cell("A2").GetString();
         string sheetName = worksheet.Name;

ExcelDetailsText.Text =
    $"\n\n\nExcel Name: {App.AppMemory.WorkbookPath ?? "No file loaded"}\n" +
    $"{worksheet.Cell("A1").GetString()}\n" +
    $"{cellA2Value}\n" +
    $"{worksheet.Cell("A3").GetString()}";
   
        // ExcelDetailsText.Text = $"\n\n\n{worksheet.Cell("A1").GetString()}\n{cellA2Value}\n{worksheet.Cell("A3").GetString()}";

         ValidateDropdownSelection(); // Enable button if valid
        GoToPage2Button.IsEnabled = true;
    }
}
private void SelectComboItem(ComboBox combo, string value)
{
    if (string.IsNullOrEmpty(value)) return;

    foreach (ComboBoxItem item in combo.Items)
    {
        if (item.Content.ToString() == value)
        {
            combo.SelectedItem = item;
            break;
        }
    }
}

        // public Page1()
        // {
        //     InitializeComponent();
        //     // Handle dropdown change

        // }

        




        // to go to page 2
        private void GoToPage2_Click(object sender, RoutedEventArgs e)
{




    Console.WriteLine("GoToPage2_Click triggered...........");
    

    if (OptionsComboBox?.SelectedItem is not ComboBoxItem selectedItem)
    {
        MessageBox.Show("Please select an option.");
        return;
    }

    string selectedValue = selectedItem.Content.ToString();

ComboBoxItem item = YearComboBox.SelectedItem as ComboBoxItem;
string selectedYear = item?.Content?.ToString() ?? "";

ComboBoxItem item1 = MonthComboBox.SelectedItem as ComboBoxItem;
string selectedMonth = item1?.Content?.ToString() ?? "";
    if (uploadedWorkbook == null)
    {
        MessageBox.Show("Please upload an Excel file first.");
        return;
    }

    if (!string.Equals(selectedValue?.Trim(), cellA2Value?.Trim(), StringComparison.OrdinalIgnoreCase))
    {
        MessageBox.Show(
            "Employee category does not match with the uploaded Excel!\n\n" +
            $"Selected: {selectedValue?.Trim()}\n" +
            $"Excel A2: {cellA2Value?.Trim()}\n\n" +
            "Please upload the correct file or select the matching category.",
            "Category Mismatch",
            MessageBoxButton.OK,
            MessageBoxImage.Warning
        );
        return;
    }
    
 
    // Save to AppMemory
  App.  AppMemory.UploadedWorkbook = uploadedWorkbook;
   App. AppMemory.SelectedCategory = selectedValue;
App.AppMemory.SelectedMonth= selectedMonth;
   App. AppMemory.SelectedYear = selectedYear;

    // Navigate to Page2
    this.NavigationService.Navigate(
        new Page2(uploadedWorkbook, selectedValue,selectedMonth,selectedYear)
    );
}

        
        private void UploadFile_Click(object sender, RoutedEventArgs e)
{
    OpenFileDialog openFileDialog = new OpenFileDialog();
    openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
    openFileDialog.Title = "Select Salary Excel File";

    if (openFileDialog.ShowDialog() != true)
        return;

    string filePath = openFileDialog.FileName;

    try
    {
        using (var tempWorkbook = new XLWorkbook(filePath))
        {
            var worksheet = tempWorkbook.Worksheet(1);
            string a2Value = worksheet.Cell("A2").GetValue<string>()?.Trim() ?? "";
            string selectedCategory = OptionsComboBox.SelectedItem is ComboBoxItem item
                ? item.Content.ToString().Trim()
                : null;

            if (!string.IsNullOrEmpty(selectedCategory))
            {
                if (string.Equals(selectedCategory, a2Value, StringComparison.OrdinalIgnoreCase))
                {
                    // PERFECT MATCH → Show your custom message
                    var result = MessageBox.Show(
                        "Employee category matches with uploaded Excel!\n\n" +
                        $"Selected: {selectedCategory}\n" +
                        $"Excel A2: {a2Value}\n\n" +
                        "Please press Continue to proceed.",
                        "Category Match Success!",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Information);

                    if (result != MessageBoxResult.OK)
                    {
                        return; // User pressed Cancel → don't load file
                    }

                    // User pressed Continue → Load file
                    LoadWorkbookAndProceed(filePath, a2Value);
                    return;
                }
                else
                {
                    // MISMATCH
                    MessageBox.Show(
                        $"Category Mismatch!\n\n" +
                        $"Selected: {selectedCategory}\n" +
                        $"Excel A2: {a2Value}\n\n" +
                        "Please upload the correct file.",
                        "Wrong File!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }
            else
            {
                // No category selected yet → just load and auto-select if possible
                LoadWorkbookAndProceed(filePath, a2Value);
            }
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Error: {ex.Message}", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

// Helper: Load file only after user confirms match
private void LoadWorkbookAndProceed(string filePath, string a2Value)
{
    uploadedWorkbook?.Dispose();
    uploadedWorkbook = new XLWorkbook(filePath);

    var worksheet = uploadedWorkbook.Worksheet(1);
    string a1 = worksheet.Cell("A1").GetValue<string>()?.Trim() ?? "";
    string a3 = worksheet.Cell("A3").GetValue<string>()?.Trim() ?? "";

    cellA2Value = a2Value;

    ExcelDetailsText.Text =
        $"{Path.GetFileName(filePath)}\n\n" +
        $"{a1}\n" +
        $"{a2Value}\n" +
        $"{a3}";
    // Auto-select ComboBox if A2 matches any item
    foreach (ComboBoxItem item in OptionsComboBox.Items)
    {
        if (string.Equals(item.Content.ToString().Trim(), a2Value, StringComparison.OrdinalIgnoreCase))
        { 
            OptionsComboBox.SelectedItem = item;
            break;
        }
    }

    ValidateDropdownSelection(); // Enable button
}








        // Handle dropdown selection changeprivate void MonthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
private void MonthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
{
    if (OptionsComboBox != null)
        OptionsComboBox.SelectedIndex = 0;

    ValidateDropdownSelection();
}

private void YearComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
{
    if (OptionsComboBox != null)
        OptionsComboBox.SelectedIndex = 0;

    ValidateDropdownSelection();
}

private void OptionsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
{
    ValidateDropdownSelection();
}


        // Validation logic
//         private void ValidateDropdownSelection()
//         {
//             if (OptionsComboBox?.SelectedItem is ComboBoxItem selectedItem && selectedItem != null)
//             {
//     string selectedValue = (OptionsComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString();
// Console.WriteLine("Validating selection. Selected Value: " + selectedValue + ", Excel A2 Value: " + cellA2Value+"month:"+MonthComboBox.SelectedItem.ToString()+" year:"+YearComboBox.SelectedItem.ToString());
//                 if (!string.IsNullOrWhiteSpace(selectedValue) && selectedValue == cellA2Value)
//                 {
//                     GoToPage2Button.IsEnabled = true;
//                 }
                
//                 else
//                 {   
//                     GoToPage2Button.IsEnabled = false;
//                 }
//             }
//         }

private void ValidateDropdownSelection()
{
    // If any combo is not initialized yet → disable and exit
    if (MonthComboBox == null ||
        YearComboBox == null ||
        OptionsComboBox == null)
    {
        GoToPage2Button.IsEnabled = false;
        return;
    }

    // If any selection is null → disable and exit (prevents all crashes)
    if (MonthComboBox.SelectedItem == null ||
        YearComboBox.SelectedItem == null ||
        OptionsComboBox.SelectedItem == null)
    {
        GoToPage2Button.IsEnabled = false;
        return;
    }

    // Safe casting
    var monthItem = MonthComboBox.SelectedItem as ComboBoxItem;
    var yearItem = YearComboBox.SelectedItem as ComboBoxItem;
    var optionItem = OptionsComboBox.SelectedItem as ComboBoxItem;

    if (monthItem == null || yearItem == null || optionItem == null)
    {
        GoToPage2Button.IsEnabled = false;
        return;
    }

    string month = monthItem.Content?.ToString() ?? "";
    string year = yearItem.Content?.ToString() ?? "";
    string category = optionItem.Content?.ToString() ?? "";

    // Enable Continue only if valid
    bool valid =
        !month.Contains("Select") &&
        !year.Contains("Select") &&
        !category.Contains("Select");

    GoToPage2Button.IsEnabled = valid;
}


        // Navigate to Page3
        private void GoToPage3_Click(object sender, RoutedEventArgs e)
        {
            ComboBoxItem item = YearComboBox.SelectedItem as ComboBoxItem;
string selectedYear = item?.Content?.ToString() ?? "";

ComboBoxItem item1 = MonthComboBox.SelectedItem as ComboBoxItem;
string selectedMonth = item1?.Content?.ToString() ?? "";

            // Navigate to Page3
            this.NavigationService.Navigate(new Page3(selectedMonth,selectedYear, selectedValue: cellA2Value));
        }





    }
}
