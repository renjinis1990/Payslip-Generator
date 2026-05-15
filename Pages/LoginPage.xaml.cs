using System.Windows;
using System.Windows.Controls;
using System.IO;
using Microsoft.Win32;
using MyDesktopApp.Pages;
namespace MyApp.Pages
{
    public partial class LoginPage : Page
    {
        public LoginPage()
        {
            InitializeComponent(); // required to wire up named controls
        }

        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string username = UsernameBox.Text;
            string password = PasswordBox.Password;
Console.WriteLine(username);
Console.WriteLine(password);
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                ErrorText.Text = "Enter username and password";
                return;
            }

            if (username == "admin" && password == "1234")
            {
              
                this.NavigationService.Navigate(new Page1());}
            else
            {
                ErrorText.Text = "Invalid username or password";
            }
        }
    
    
    
   
    
    }


       
}
