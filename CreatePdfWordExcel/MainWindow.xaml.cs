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

namespace CreatePdfWordExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SaveToPdfBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogMulti(textForSaveInPdfTBox.Text, 1);
        }

        private void SaveToWordBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogMulti(textForSaveInWordTBox.Text, 2);
        }

        private void SaveToExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialogMulti(textForSaveInExcelTBox.Text, 3);
        }

        public void SaveFileDialogMulti(string dataString, int indexSaveFileDialog)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            if (indexSaveFileDialog == 1)
            {
                saveFileDialog.FileName = "reportToPDF";
                saveFileDialog.Filter = "pdf documents |*.pdf";

                if ((bool)saveFileDialog.ShowDialog())
                {
                    Word.Application application = new Word.Application();
                    Word.Document document = application.Documents.Add();
                    Word.Paragraph paragraph = document.Paragraphs.Add();

                    paragraph.Range.Text = dataString;
                    // Add image
                    //application.Selection.InlineShapes.(@"Z:\WpfAppPractice\CreatePdfWordExcel\Resources\babyYoda.jpeg");

                    document.SaveAs(saveFileDialog.FileName, Word.WdSaveFormat.wdFormatPDF);

                }
            }
            else if (indexSaveFileDialog == 2)
            {
                saveFileDialog.FileName = "reportToWord";
                saveFileDialog.Filter = "word documents |*.doc";

                if ((bool)saveFileDialog.ShowDialog())
                {
                    Word.Application application = new Word.Application();
                    Word.Document document = application.Documents.Add();
                    Word.Paragraph paragraph = document.Paragraphs.Add();

                    paragraph.Range.Text = dataString;

                    document.SaveAs(saveFileDialog.FileName);
                    application.Quit();

                }
            }
            else if (indexSaveFileDialog == 3)
            {
                saveFileDialog.FileName = "reportToWord";
                saveFileDialog.Filter = "excel documents |*.xlsx";

                if ((bool)saveFileDialog.ShowDialog())
                {
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Add();

                    Excel.Worksheet worksheet = (Excel.Worksheet)application.Worksheets.Add();
                    worksheet.Name = "Firts page";
                    worksheet.UsedRange.Cells[1, 1] = dataString;

                    workbook.SaveAs(saveFileDialog.FileName);
                    application.Quit();
                }
            }
        }
    }
}
