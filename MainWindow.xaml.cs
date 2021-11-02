using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Ionic.Zip;

namespace Барак
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        Microsoft.Office.Interop.Word.Application app = null;
        Microsoft.Office.Interop.Excel.Application app2 = null;
        Document doc = null;
        Workbook doc2 = null;

        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            button.Content = "Виконую...";
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Файли для підбору паролю (*.docx, *.xlsx, *.zip)|*.docx;*.xlsx;*.zip";
            file.InitialDirectory = @"C:\Users\kukha\source\repos\Барак\bin\Debug";
            file.ShowDialog();
            if (file.FileName.EndsWith(".docx"))
            {
                app = new Microsoft.Office.Interop.Word.Application();
                app.Visible = false;
                for (int i = 0; i < 1000; i++)
                {
                    button.Content = i.ToString();
                    try
                    {
                        doc = app.Documents.Open(file.FileName, PasswordDocument: i.ToString());
                        MessageBox.Show("Password is: " + i.ToString());
                        if (doc != null)
                        {
                            doc.Close();
                        }
                        app.Quit();
                        break;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                
            }
            if (file.FileName.EndsWith(".xlsx"))
            {
                app2 = new Microsoft.Office.Interop.Excel.Application();
                app2.Visible = false;
                for (int i = 0; i < 1000; i++)
                {
                    button.Content = i.ToString();
                    try
                    {
                        doc2 = app2.Workbooks.Open(file.FileName, Password: i.ToString());
                        MessageBox.Show("Password is: " + i.ToString());
                        if (doc2 != null)
                        {
                            doc2.Close();
                        }
                        app2.Quit();
                        break;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                    }
                    catch(Exception)
                    {
                        throw;
                    }
                }
            }
            if (file.FileName.EndsWith(".zip"))
            {
                ZipFile zip = ZipFile.Read(file.FileName, new ReadOptions { Encoding = Encoding.GetEncoding("cp866") });
                for (int i = 0; i < 1000; i++)
                {
                    button.Content = i.ToString();
                    try
                    {
                        zip[0].ExtractWithPassword(file.FileName, i.ToString());
                    }
                    catch (BadPasswordException)
                    {
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("Password is: " + i);
                        break;
                    }
                }
            }
            if (app != null)
            {
                try
                {
                    app.Quit();
                }
                catch (Exception)
                {
                }
            }
            if (app2 != null)
            {
                try
                {
                    app2.Quit();
                }
                catch (Exception)
                {
                }
            }
            button.Content = "Вибрати файл...";
        }
    }
}
