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
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DoubleList = System.Collections.Generic.List<System.Collections.Generic.List<System.String>>;
using TripleList = System.Collections.Generic.List<System.Collections.Generic.List<System.Collections.Generic.List<System.String>>>;

namespace Excel_Transformation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        const int YEAR = 2017;
        List<String> paths = new List<String>();

        public MainWindow()
        {
            InitializeComponent();
        }
        static void DoMagic(IList<String> paths)
        {
            if (paths.Count == 0)
                return;
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Add();
            Excel._Worksheet output = book.ActiveSheet;
            output.Name = "Sheet1";
            int cur_row = 1;

            foreach (var sheet in GetData(paths))
            {
                var count = sheet.Count;
                var (species, maxY) = GetSpecies(sheet, 1, 5);

                int x = 2;
                string header = sheet[x][0];
                while(!String.IsNullOrEmpty(header))
                {
                    var info = GetInfo(sheet[x]);
                    for (int y = 5; y < maxY; y++)
                    {
                        string current = sheet[x][y];
                        if (!String.IsNullOrEmpty(current))
                        {
                            Write(ref output, cur_row++, species[y], current, info);
                        }
                    }
                    if (++x >= count )
                        break;
                    else
                        header = sheet[x][0];
                }
                break; ////////////////////////////// only the first one has useful data?
            }

            Marshal.ReleaseComObject(output);
            object missing = System.Reflection.Missing.Value;
            book.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/NELESY.xlsx", missing,
                missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);
            book.Close();
            Marshal.ReleaseComObject(output);
            app.Quit();
            Marshal.ReleaseComObject(app);

            paths.Clear();
        }
        static void Write(ref Excel._Worksheet sheet,int row, String bird, string count, Info info)
        {
            sheet.Cells[row, 1] = bird;
            sheet.Cells[row, 2] = info.Point;
            sheet.Cells[row, 3] = info.Name ;
            sheet.Cells[row, 4] = info.Name;
            sheet.Cells[row, 5] = info.Date;
            sheet.Cells[row, 6] = count;
            sheet.Cells[row, 7] = "jedinci";
        }
        static TripleList GetData(IList<String> Paths)
        {
            Excel.Application app1 = new Excel.Application();
            var data = new TripleList();
            foreach (var path in Paths)
            {
                foreach (Excel._Worksheet workSheet in app1.Workbooks.Open(path).Sheets)
                {
                    Excel.Range range = workSheet.UsedRange;
                    int rowCount = range.Rows.Count;
                    int colCount = range.Columns.Count;

                    var sheet = new DoubleList();
                    for (int x = 1; x <= colCount; x++)
                    {
                        var column = new List<String>();
                        for (int y = 1; y <= rowCount; y++)
                        {
                            string cell = range.Cells[y, x]?.Value2?.ToString();
                            column.Add(cell is null ? cell : (cell + "\t"));
                        }
                        sheet.Add(column);
                    }
                    data.Add(sheet);

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(workSheet);
                }
            }
            app1.Quit();
            Marshal.ReleaseComObject(app1);
            return data;
        }
        
        static (Dictionary<int,String>, int) GetSpecies(DoubleList sheet, int startX, int startY)
        {
            var dict = new Dictionary<int, String>();
            string content = sheet[startX][startY];
            int lenght = sheet[startX].Count;

            while (!String.IsNullOrEmpty(content))
            {
                dict.Add(startY, content);
                if (++startY == lenght - 1)
                    break;
                content = sheet[startX][startY];
            }
            return (dict, startY);
        }
        static Info GetInfo(List<String> column)
        {
            if (string.IsNullOrEmpty(column[0]))
                return null;

            var date = column[2].Split('.');
            return new Info(column[0], column[1], new DateTime(YEAR, int.Parse(date[1]), int.Parse(date[0])));
        }
        class Info
        {
            public String Point;
            public String Name;
            private DateTime date;
            public String Date => String.Format("{0:yyyyMMdd}", date);
            
            public Info(string point, string name, DateTime _date)
            {
                date = _date;
                Name = name;
                Point = point;
            }
        }
        private void Process_Click(object sender, RoutedEventArgs e)
        {
            DoMagic(paths);
        }

        private void Choose_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                paths.Add(openFileDialog.FileName);
            }
        }
    }
    
}
