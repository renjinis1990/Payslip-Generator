using System.Configuration;
using System.Data;
using System.Windows;
using ClosedXML.Excel; //to read the excel file

namespace MyDesktopApp;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    public static class AppMemory
{
    public static string WorkbookPath { get; set; }
        public static XLWorkbook UploadedWorkbook { get; set; }

        public static string SelectedMonth { get; set; }
        public static string SelectedYear { get; set; }
        public static string SelectedCategory { get; set; }
    
}

}

