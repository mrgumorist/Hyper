using Microsoft.Win32;
using Aspose.Cells;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Hyper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string @path;
        string @path2;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                @path = (openFileDialog.FileName);
                patth.Text = path;
            }
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<string> list = new List<string>();
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbooks workBooks = appExcel.Workbooks;
            Excel.Workbook excelSheet = workBooks.Open(@path, false, ReadOnly: true);

            foreach (Excel.Worksheet worksheet1 in excelSheet.Worksheets)
            {
                Excel.Hyperlinks hyperLinks = worksheet1.Hyperlinks;
                foreach (Excel.Hyperlink lin in hyperLinks)
                {
                    list.Add(lin.Address);
                }
            }
            @path2 += @"\res.xlsx";
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            for (int i = 1; i < list.Count()+1; i++)
            {
                string el = "A" + i.ToString();
                Aspose.Cells.Cell cell = worksheet.Cells[el];
                int index = worksheet.Hyperlinks.Add(el, 1, 1, list[i-1]);
                worksheet.Hyperlinks[index].TextToDisplay = list[i - 1];
                worksheet.Hyperlinks[index].ScreenTip = list[i - 1];
              
            }
            workbook.Save(@path2);
            MessageBox.Show("Succesfull");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
           @path2= dialog.SelectedPath;
            patth_Copy.Text = @path2;
        }
    }
}
