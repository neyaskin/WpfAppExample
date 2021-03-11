using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace CreatePdfDocXlsx
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

        private void SaveToPdfBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogPdf(new SaveFileDialog(), "gg");

        }

        private void SaveToDocBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogDoc(new SaveFileDialog(), "gg");

        }

        private void SaveToXlsxBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogExcel(new SaveFileDialog(), "gg");

        }

        private void SaveToBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogChangeFormat(new SaveFileDialog(), "gg");

        }

        public void SaveFileDialogPdf(SaveFileDialog saveFileDialog, string data)
        {
            saveFileDialog.FileName = "report Pdf";
            saveFileDialog.Filter = "pdf documents |*.pdf";

            if ((bool)saveFileDialog.ShowDialog())
            {
                Word.Application application = new Word.Application();
                Word.Document document = application.Documents.Add();
                Word.Paragraph paragraph = document.Paragraphs.Add();

                paragraph.Range.Text = data;

                document.SaveAs(saveFileDialog.FileName, Word.WdSaveFormat.wdFormatPDF);
                application.Quit();
            }
        }

        public void SaveFileDialogDoc(SaveFileDialog saveFileDialog, string data)
        {
            saveFileDialog.FileName = "report Doc";
            saveFileDialog.Filter = "doc documents |*.doc";
            if ((bool)saveFileDialog.ShowDialog())
            {
                Word.Application application = new Word.Application();
                Word.Document document = application.Documents.Add();
                Word.Paragraph paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = data;
                document.SaveAs(saveFileDialog.FileName);
                application.Quit();
            }
        }
        public void SaveFileDialogExcel(SaveFileDialog saveFileDialog, string data)
        {
            saveFileDialog.FileName = "report Excel";
            saveFileDialog.Filter = "excel documents |*.xlsx";

            if ((bool)saveFileDialog.ShowDialog())
            {
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add();
                Excel.Worksheet worksheet = application.Worksheets.Add();

                worksheet.Name = data;
                worksheet.UsedRange.Cells[1, 1] = data;

                workbook.SaveAs(saveFileDialog.FileName);
                application.Quit();
            }
        }

        public void SaveFileDialogChangeFormat(SaveFileDialog saveFileDialog, string data)
        {
            saveFileDialog.Filter = "doc documents |*.doc|pdf documents |*.pdf|excel documents |*.xlsx";
            
            if ((bool)saveFileDialog.ShowDialog())
            {
                if (saveFileDialog.FilterIndex == 1)
                {
                    Word.Application application = new Word.Application();
                    Word.Document document = application.Documents.Add();
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    paragraph.Range.Text = data;
                    document.SaveAs(saveFileDialog.FileName);
                    application.Quit();
                }
                else if (saveFileDialog.FilterIndex == 2)
                {
                    Word.Application application = new Word.Application();
                    Word.Document document = application.Documents.Add();
                    Word.Paragraph paragraph = document.Paragraphs.Add();

                    paragraph.Range.Text = data;

                    document.SaveAs(saveFileDialog.FileName, Word.WdSaveFormat.wdFormatPDF);
                    application.Quit();
                }
                else if (saveFileDialog.FilterIndex == 3)
                {
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Add();
                    Excel.Worksheet worksheet = application.Worksheets.Add();

                    worksheet.Name = data;
                    worksheet.UsedRange.Cells[1, 1] = data;

                    workbook.SaveAs(saveFileDialog.FileName);
                    application.Quit();
                }
            }
        }
    }
}
