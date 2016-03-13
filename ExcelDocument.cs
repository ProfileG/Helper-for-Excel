using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;

namespace FreshMap
{
    class ExcelDocument
    {
        public static bool Stats;
        private _Excel.Application _application = null;
        private _Excel.Workbook _workBook = null;
        private _Excel.Worksheet _workSheet = null;
        private object _missingObj = System.Reflection.Missing.Value;
        

   


        //КОНСТРУКТОР
        public ExcelDocument()
        {
            _application = new _Excel.Application();
            _workBook = _application.Workbooks.Add(_missingObj);
            _workSheet = (_Excel.Worksheet)_workBook.Worksheets.get_Item(1);

            //Статус и обработка данных.
            Stats = true;
            Parametrs.ExcelRowsNum = _workSheet.UsedRange.Rows.Count;
            Parametrs.ExcelColumnsNum = _workSheet.UsedRange.Columns.Count;
        }

        public ExcelDocument(string pathToTemplate, int sheets)
        {
            object pathToTemplateObj = pathToTemplate;

            _application = new _Excel.Application();
            //Открываем книгу.                                                                                                                                                        
             _workBook = _application.Workbooks.Open((string)pathToTemplateObj, 0, false, 5, "", "", false, _Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
             _workSheet = (_Excel.Worksheet)_workBook.Sheets[sheets];
            //Статус и обработка данных.
            Stats = true;
            Parametrs.ExcelRowsNum = _workSheet.UsedRange.Rows.Count;
            Parametrs.ExcelColumnsNum = _workSheet.UsedRange.Columns.Count;

        }

        // ВИДИМОСТЬ ДОКУМЕНТА
        public bool Visible
        {
            get
            {
                return _application.Visible;
            }
            set
            {
                _application.Visible = value;
            }
        }

        // ВСТАВКА ЗНАЧЕНИЯ В ЯЧЕЙКУ
        public void SetCellValue(string cellValue, int rowIndex, int columnIndex)
        {
            _workSheet.Cells[rowIndex, columnIndex] = cellValue;
            
        }

        // ЧТЕНИЕ ЗНАЧЕНИЯ
        public string GetCellValue(int rowIndex, int columnIndex)
        {
            

            _Excel.Range cellRange = (_Excel.Range)_workSheet.Cells[rowIndex, columnIndex];
            return cellRange.Text;
            
        }

        // CОХРАНЕНИЕ ДОКУМЕНТА
        public void Save()
        {
          

                  _workBook.Save();
             //System.GC.Collect();
        }


        // ЗАКРЫТИЕ ДОКУМЕНТА
        public void Close()
        {
            Stats = false;

            _workBook.Close(false, _missingObj, _missingObj);

            _application.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);

            _application = null;
            _workBook = null;
            _workSheet = null;

            System.GC.Collect();
        }

      
    }
}
