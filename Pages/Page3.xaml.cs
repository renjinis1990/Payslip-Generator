using System;
using System.IO;
using System.Windows.Controls;
using System.Windows;
using System.Diagnostics;
using System.Collections.Generic;

namespace MyDesktopApp.Pages
{
    public partial class Page3 : Page
    {
        private string selectedValue;
        private string selectedMonth;
        private string selectedYear;

        public Page3(string selectedMonth, string selectedYear, string selectedValue)
        {
            InitializeComponent();

            this.selectedMonth = selectedMonth;
            this.selectedYear = selectedYear;
            this.selectedValue = selectedValue;

            LoadFolderTree( selectedValue);    
             LoadPdfFiles(folderPath: Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Payslip",
                selectedValue,
                selectedYear,
                selectedMonth
            )   );        // Load PDFs for default selected folder
        }
 private void LoadFolderTree(string folderValue)
        {
            string basePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Payslip",
                selectedValue
            );

            FolderTree.Items.Clear();

            if (!Directory.Exists(basePath))
            {
                MessageBox.Show("No folder found for selected category.");
                return;
            }

            TreeViewItem categoryNode = new TreeViewItem
            {
                Header = selectedValue,
                Tag = basePath
            };

            string[] yearFolders = Directory.GetDirectories(basePath);

            foreach (string yearPath in yearFolders)
            {
                string year = Path.GetFileName(yearPath);

                TreeViewItem yearNode = new TreeViewItem
                {
                    Header = year,
                    Tag = yearPath
                };

                string[] monthFolders = Directory.GetDirectories(yearPath);

                foreach (string monthPath in monthFolders)
                {
                    string month = Path.GetFileName(monthPath);

                    TreeViewItem monthNode = new TreeViewItem
                    {
                        Header = month,
                        Tag = monthPath
                    };

                    yearNode.Items.Add(monthNode);
                }

                categoryNode.Items.Add(yearNode);
            }

            FolderTree.Items.Add(categoryNode);
            categoryNode.IsExpanded = true;
        }

        private void BackToPage1_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Navigate(new Page1());
        }

        // ✔ Default Load – loads PDFs for selected Category → Year → Month
        // private void LoadPdfFiles()
        // {
        //     string baseFolder = Path.Combine(
        //         Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        //         "Payslip",
        //         selectedValue,
        //         selectedYear,
        //         selectedMonth
        //     );

        //     if (Directory.Exists(baseFolder))
        //         LoadPdfFiles(baseFolder);
        //     else
        //         PdfListPanel.Children.Clear();
        // }

        // ✔ Loads PDFs from a specific folder
        private void LoadPdfFiles(string folderPath, string? filter = null)
{
    PdfListPanel.Children.Clear();

    if (!Directory.Exists(folderPath))
    {
        PdfListPanel.Children.Add(new TextBlock
        {
            Text = "No PDF files found.",
            Foreground = System.Windows.Media.Brushes.Gray,
            FontSize = 14,
            Margin = new Thickness(5)
        });
        return;
    }

    string[] pdfFiles = Directory.GetFiles(folderPath, "*.pdf");

    if (pdfFiles.Length == 0)
    {
        PdfListPanel.Children.Add(new TextBlock
        {
            Text = "No PDF files found.",
            Foreground = System.Windows.Media.Brushes.Gray,
            FontSize = 14,
            Margin = new Thickness(5)
        });
        return;
    }

    foreach (string pdf in pdfFiles)
    {
        string fileName = Path.GetFileName(pdf);

        StackPanel row = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(5)
        };

        row.Children.Add(new TextBlock
        {
            Text = fileName,
            Width = 280,
            FontSize = 14,
            VerticalAlignment = VerticalAlignment.Center
        });

        Button openBtn = new Button
        {
            Content = "Open",
            Tag = pdf,
            Width = 60,
            Height = 24,
            Margin = new Thickness(10, 0, 0, 0)
        };
        openBtn.Click += OpenPdf_Click;

        row.Children.Add(openBtn);

        PdfListPanel.Children.Add(row);
    }
}

        // ✔ Build folder structure for TreeView
        // ✔ Open PDF
        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string filePath && File.Exists(filePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open file:\n{ex.Message}");
                }
            }
        }

        // ✔ Refresh PDF List
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
             LoadPdfFiles(folderPath: Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Payslip",
                selectedValue,
                selectedYear,
                selectedMonth
            )   );
            
        }

        // ✔ User selects a folder in TreeView
  private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
{
    if (e.NewValue is TreeViewItem selectedNode)
    {
        string folderPath = selectedNode.Tag.ToString();

        // Load PDFs in the right panel
        if (Directory.Exists(folderPath))
        {
            LoadPdfFiles(folderPath);
        }

        // 👉 If Year selected — Load months dynamically
        if (IsYearFolder(folderPath))
        {
            LoadMonthNodes(selectedNode, folderPath);
        }
    }
}
private bool IsYearFolder(string path)
{
    string folderName = Path.GetFileName(path);

    return folderName.Length == 4 && folderName.All(char.IsDigit);
}
private void LoadMonthNodes(TreeViewItem yearNode, string yearPath)
{
    yearNode.Items.Clear(); // prevent duplicates

    string[] monthFolders = Directory.GetDirectories(yearPath);

    foreach (string monthPath in monthFolders)
    {
        string month = Path.GetFileName(monthPath);

        TreeViewItem monthNode = new TreeViewItem
        {
            Header = month,
            Tag = monthPath
        };

        yearNode.Items.Add(monthNode);
    }

    yearNode.IsExpanded = true;
}


        // ✔ Search filter
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string baseFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Payslip",
                selectedValue,
                selectedYear,
                selectedMonth
            );

            if (!Directory.Exists(baseFolder))
                return;

            PdfListPanel.Children.Clear();

            string search = SearchTextBox.Text.Trim().ToLower();
            string[] pdfFiles = Directory.GetFiles(baseFolder, "*.pdf");

            foreach (string pdf in pdfFiles)
            {
                if (Path.GetFileName(pdf).ToLower().Contains(search))
                {
                    StackPanel row = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Margin = new Thickness(5)
                    };

                    row.Children.Add(new TextBlock
                    {
                        Text = Path.GetFileName(pdf),
                        Width = 280,
                        FontSize = 14
                    });

                    Button openBtn = new Button
                    {
                        Content = "Open",
                        Tag = pdf,
                        Width = 60,
                        Height = 24,
                        Margin = new Thickness(10, 0, 0, 0)
                    };

                    openBtn.Click += OpenPdf_Click;

                    row.Children.Add(openBtn);
                    PdfListPanel.Children.Add(row);
                }
            }
        }

        private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Text = "";
            LoadPdfFiles(folderPath: Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Payslip",
                selectedValue,
                selectedYear,
                selectedMonth
            )   );
        }
    };
}

