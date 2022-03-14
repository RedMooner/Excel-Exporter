using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OneWorks.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace OneWorks
{
    public class ExcelExporter
    {
        public class Header
        {
            public string Title { get; set; }
            public Rows Row { get; set; }
            public enum Rows { A, B, C, D, E, F, G, H, J, K, L, M, N, P, Q, R, S, T, V, W, X, Y, Z }
            public Header(string title, Rows row)
            {
                Title = title;
                Row = row;
            }
        }
        public class Cell
        {
            public enum Rows { A, B, C, D, E, F, G, H, J, K, L, M, N, P, Q, R, S, T, V, W, X, Y, Z }
            public Rows Row { get; set; }
            public string Data { get; set; }
            public Cell(Rows row, string data)
            {
                //if(string.IsNullOrWhiteSpace(data))
                //    throw new Exception("NullReference Exception data does't enter");
                Row = row;
                Data = data;
            }
        }
        private string FileName { get; set; }
        public List<Header> Headers = new List<Header>();
        public List<Cell> Cells = new List<Cell>();
        public ExcelExporter(string filename)
        {
            if (string.IsNullOrWhiteSpace(filename))
                throw new Exception("NullReference Exception filename does't enter");
            FileName = filename;
        }
        public void Export()
        {
            //TODO: передавать данные отдельным классом
            try
            {
                Excel.Application excelApp = new Excel.Application();
                // Сделать приложение Excel видимым
                excelApp.Visible = true;
                excelApp.Workbooks.Add();
                Excel._Worksheet workSheet = excelApp.ActiveSheet;
                // Установить заголовки столбцов в ячейках
                for (int i = 0; i < Headers.Count; i++)
                    workSheet.Cells[1, Headers[i].Row.ToString()] = Headers[i].Title;
                int row = 2;
                string current_row = Cells[0].Row.ToString();
                for (int i = 0; i < Cells.Count; i++)
                {
                    if (current_row != Cells[i].Row.ToString())
                    {
                        current_row = Cells[i].Row.ToString();
                        row = 2;
                    }
                    workSheet.Cells[row, Cells[i].Row.ToString()] = Cells[i].Data;
                    row++;
                }
                excelApp.DisplayAlerts = false;
                workSheet.SaveAs(string.Format(@"{0}\" + $"{FileName}.xlsx", Environment.CurrentDirectory));
                excelApp.Quit();
            }
            catch (Exception exc)
            {
                MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
            }
        }

    }

}