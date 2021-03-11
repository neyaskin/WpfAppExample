using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Printing;
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

namespace OpenPrintColorFolderFileDialog
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if ((bool)openFileDialog.ShowDialog())
            {
                var streamData = "";

                using (StreamReader streamReader = new StreamReader(openFileDialog.FileName))
                {
                    streamData = streamReader.ReadToEnd();
                }

                MessageBox.Show(streamData, $"File name: {openFileDialog.FileName}");
            }
        }

        private void PrintFileBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] fileBytes = null;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if ((bool)openFileDialog.ShowDialog())
            {
                fileBytes = File.ReadAllBytes(openFileDialog.FileName);
            }

            PrintDialog printDialog = new PrintDialog();
            if ((bool)printDialog.ShowDialog())
            {
                DocumentViewer documentViewer = new DocumentViewer();

                printDialog.PrintDocument(((IDocumentPaginatorSource)documentViewer.Document).DocumentPaginator,
                    System.Text.Encoding.UTF8.GetString(fileBytes));
            }
        }
    }
}