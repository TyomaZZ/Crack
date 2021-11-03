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
using System.Threading;
using System.Diagnostics;

namespace Барак
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string path = @"C:\Файли\Crack";
        Microsoft.Office.Interop.Word.Application wordApp = null;
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        OpenFileDialog fileDialog;
        Document documentWord = null;
        Workbook workbookExcel = null;
        Thread thread1;
        string password;

        public MainWindow()
        {
            InitializeComponent();
        }

        public delegate void DelegateOnFinish();
        public void OnFinish()
        {
            button.IsEnabled = true;
            button.Content = "Вибрати файл...";
            Status.Content = "\"" + fileDialog.FileName.Split('\\')[fileDialog.FileName.Split('\\').Length - 1] + "\"" + " = " + password;
            Status.MouseLeftButtonDown += new MouseButtonEventHandler(OpenCrackedFile);
            Status.Cursor = Cursors.Hand;
            prgrsBar.IsIndeterminate = false;
            WindowStyle = WindowStyle.SingleBorderWindow;
            Height += 40;
        }

        public delegate void DelegatePasswordDisplay(string a);
        public void CounterButton(string a)
        {
            Status.Content = a;
        }

        public delegate void DelegateOnStart();
        public void OnStart()
        {
            WindowStyle = WindowStyle.None;
            Height -= 40;
            button.Content = "Виконую...";
            Status.Cursor = Cursors.Arrow;
            Status.MouseLeftButtonDown -= new MouseButtonEventHandler(OpenCrackedFile);
            prgrsBar.IsIndeterminate = true;
        }

        public void OpenCrackedFile(object sender, MouseButtonEventArgs e)
        {
            Status.Cursor = Cursors.Arrow;
            Status.MouseLeftButtonDown -= new MouseButtonEventHandler(OpenCrackedFile);
            if (fileDialog.FileName.EndsWith(".docx"))
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                documentWord = wordApp.Documents.Open(fileDialog.FileName, PasswordDocument: password);
                return;
            }
            if (fileDialog.FileName.EndsWith(".xlsx"))
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;
                workbookExcel = excelApp.Workbooks.Open(fileDialog.FileName, Password: password);
                return;
            }
            if (fileDialog.FileName.EndsWith(".zip"))
            {
                Process.Start("explorer.exe", path);
                return;
            }
            MessageBox.Show("Unknown error");
        }
        
        private void DocxFlash()
        {
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Dispatcher.BeginInvoke(new DelegateOnStart(OnStart));
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++ )
                {
                    for (int k = 0; k < 10; k++)
                    {
                        Dispatcher.BeginInvoke(new DelegatePasswordDisplay(CounterButton), i.ToString() + j.ToString() + k.ToString());
                        try
                        {
                            documentWord = wordApp.Documents.Open(fileDialog.FileName, PasswordDocument: i.ToString() + j.ToString() + k.ToString());
                            password = i.ToString() + j.ToString() + k.ToString();
                            if (documentWord != null)
                            {
                                documentWord.Close();
                            }
                            wordApp.Quit();
                            Dispatcher.BeginInvoke(new DelegateOnFinish(OnFinish));
                            thread1.Abort();
                            return;
                        }
                        catch (System.Runtime.InteropServices.COMException) { }
                    }
                }
            }
        }

        private void xlxFlash()
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            Dispatcher.BeginInvoke(new DelegateOnStart(OnStart));
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    for (int k = 0; k < 10; k++)
                    {
                        Dispatcher.BeginInvoke(new DelegatePasswordDisplay(CounterButton), i.ToString() + j.ToString() + k.ToString());
                        try
                        {
                            workbookExcel = excelApp.Workbooks.Open(fileDialog.FileName, Password: i.ToString() + j.ToString() + k.ToString());
                            password = i.ToString() + j.ToString() + k.ToString();
                            if (workbookExcel != null)
                            {
                                workbookExcel.Close();
                            }
                            excelApp.Quit();
                            Dispatcher.BeginInvoke(new DelegateOnFinish(OnFinish));
                            thread1.Abort();
                            return;
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        { }
                    }
                }
            }
        }

        public void ZipCrack()
        {
            Dispatcher.BeginInvoke(new DelegateOnStart(OnStart));
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    for (int k = 0; k < 10; k++)
                    {
                        Dispatcher.BeginInvoke(new DelegatePasswordDisplay(CounterButton), i.ToString() + j.ToString() + k.ToString());
                        if (ZipFile.CheckZipPassword(fileDialog.FileName, i.ToString() + j.ToString() + k.ToString()))
                        {
                            password = i.ToString() + j.ToString() + k.ToString();
                            Dispatcher.BeginInvoke(new DelegateOnFinish(OnFinish));
                            thread1.Abort();
                            return;
                        }
                    }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            button.IsEnabled = false;

            fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Файли для підбору паролю |*.docx;*.xlsx;*.zip"; //(*.docx, *.xlsx, *.zip)
            fileDialog.InitialDirectory = path;
            fileDialog.ShowDialog();

            if (fileDialog.FileName.EndsWith(".docx"))
            {
                thread1 = new Thread(new ThreadStart(DocxFlash));
                thread1.Start();
            }

            if (fileDialog.FileName.EndsWith(".xlsx"))
            {
                thread1 = new Thread(new ThreadStart(xlxFlash));
                thread1.Start();
            }

            if (fileDialog.FileName.EndsWith(".zip"))
            {
                thread1 = new Thread(new ThreadStart(ZipCrack));
                thread1.Start();
            }

            if (fileDialog.FileName.Equals(""))
            {
                button.IsEnabled = true;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (thread1 != null)
            {
                thread1.Abort();
            }
            try
            {
                if (documentWord != null)
                {
                    documentWord.Close();
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                }
                if (workbookExcel != null)
                {
                    workbookExcel.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
            catch (Exception) { }
        }
    }
}
