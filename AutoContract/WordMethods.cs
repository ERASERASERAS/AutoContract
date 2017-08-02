using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace AutoContract
{
    public class WordMethods
    {



        private Object _missingObj = System.Reflection.Missing.Value;
        private Object _trueObj = true;
        private Object _falseObj = false;

        private Word._Application _application;
        private Word._Document _document;

        private Object _templatePathObj;

        private Word.Range _currentRange = null;

        private Word.Table _table = null;

        private WordSelection _selection;

        


        public WordSelection Selection
        {
            get { return _selection; }
            set { throw new Exception("Ошибка! Доступно только для чтения."); }
        }

        public Word.Range CurrentRange
        {
            get { return _currentRange; }
        }

        public static char NewLineChar { get { return (char)11; } }

        public bool Closed
        {
            get
            {
                if (_application == null || _document == null) { return true; }
                else { return false; }
            }
        }

        public bool Visible
        {
            get
            {
                if (Closed) { throw new Exception("Ошибка при попытке изменить видимость Microsoft Word. Программа или документ уже закрыты."); }
                return _application.Visible;
            }
            set
            {
                if (Closed) { throw new Exception("Ошибка при попытке изменить видимость Microsoft Word. Программа или документ уже закрыты."); }
                _application.Visible = value;
            }
        }

        public void goToBegin()
        {
            object start = 0;
            object end = 0;
            this._currentRange = this._document.Range(ref start, ref end);
            //_selection = new WordSelection(_currentRange);
        }


        public void goToBookmark(string bookmarkName)
        {
           int i = _document.Bookmarks.Count; 
            if (Closed) throw new Exception("Документ был закрыт.");

            Object bookmarkNameObj = bookmarkName;
            Word.Range bookmarkRange = null;

            try
            {
                bookmarkRange = _document.Bookmarks.get_Item(ref bookmarkNameObj).Range;
               // _currentRange = _document.Bookmarks.get_Item(ref bookmarkNameObj).Range;
            }
            catch (Exception e)
            {
                throw new Exception("Возникла ошибка при поиске закладки" + bookmarkName + "\nОписание ошибки:" + e.Message);
                Console.WriteLine("Возникла ошибка при поиске закладки" + bookmarkName + "\nОписание ошибки:" + e.Message);

            }
            
            _currentRange = bookmarkRange;
            _currentRange.Font.Name = "Times New Roman";
            _currentRange.Font.Size = 12;
            
            _selection = new WordSelection(_currentRange);
        }



       

        public WordMethods(string templatePath): this(templatePath, false) { }

        public WordMethods(string templatePath, bool startVisible)
        {
            //создаем обьект приложения word
            _application = new Word.Application();

            // создаем путь к файлу используя имя файла
            _templatePathObj = templatePath;

            // если вылетим не этом этапе, приложение останется открытым
            try
            {
                _document = _application.Documents.Add(ref  _templatePathObj, ref _missingObj, ref _missingObj, ref _missingObj);
            }
            catch (Exception error)
            {
                this.close();
                throw error;
            }
            Visible = startVisible;

            // устанавливаем текущую позицию в начало документа
            goToBegin();
        }

        public void close()
        {
            if (_document != null)
            {
                _document.Close(ref _falseObj, ref  _missingObj, ref _missingObj);
            }
            _application.Quit(ref _missingObj, ref  _missingObj, ref _missingObj);
            _document = null;
            _application = null;
        }

        public void replace(string toFind, string toReplace)
        {
            if (Closed) { throw new Exception("Документ уже закрыт"); }

            object toFindObj = toFind;
            object toReplaceObj = toReplace;
            Word.Range wordRange;
            object typeOfFindAndRep = Word.WdReplace.wdReplaceAll;

            try
            {
                for (int i = 1; i <= _document.Sections.Count; i++)
                {
                    wordRange = _document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange.Find;
                    object[] wordFindParams = new object[15] { toFindObj,_missingObj,_missingObj,_missingObj,_missingObj,_missingObj,_missingObj,_missingObj,_missingObj,toReplaceObj,
                                                               typeOfFindAndRep,_missingObj,_missingObj,_missingObj,_missingObj};
                    wordFindObj.GetType().InvokeMember("Execute", System.Reflection.BindingFlags.InvokeMethod, null, wordFindObj, wordFindParams);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Ошибка при замене строк\n"+e.Message);
            }

        }

        public static void FillShowTemplate(Action<WordMethods> method,WordMethods wm)
        {
            try
            {
                method(wm);
            }
            catch (Exception e)
            {
                if (wm == null) throw new Exception("Ошибка! Документ был закрыт.");
                else throw new Exception("Возникла некоторая ошибка при выполнении действия");
            }
        }

        public void insertTable(int numRows, int numColumns,bool flag)  //flag - определяет,какую надо вставить таблицу. false - календарный план или сметная стоимость типа "человеко-часы", true - сметная стоимость.
        {
            insertTable(numRows, numColumns, BorderType.Single , flag);
        }

        public void insertTable(int numRows, int numColumns, BorderType border, bool flag)
        {

            _table = _document.Tables.Add(_currentRange, numRows, numColumns, ref _missingObj, ref _missingObj);
            
            switch (border)
            {
                case BorderType.None:
                    _table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                    _table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                    break;
                case BorderType.Single:
                    _table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    _table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    break;
                case BorderType.Double:
                    _table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
                    _table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
                    break;
                default:
                    _table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                    _table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                    break;
            }

            if (flag)
            {
                _table.Rows[numRows].Cells.Merge();
            }
            
            _currentRange = _table.Range;        
            _selection = new WordSelection(_currentRange, false);
            
        }

        public void setHeightForRow(int index, int height)
        {
            _table.Rows[index].SetHeight(height, Word.WdRowHeightRule.wdRowHeightExactly);
        }

        public void setWidthForCell(int i, int j, int value)
        {
            _table.Cell(i, j).Width = value;
        }

        public void setTableAlignment(string tableName, Word.WdCellVerticalAlignment alignment)
        {
            //_currentRange.HorizontalInVertical = Word.WdHorizontalInVerticalType.
        }

        public void setColumnWidth(int columnIndex, int widthPixels)
        {
            if (_table == null) { throw new Exception("Ошибка при установке ширины колонки в таблице Word - текущая таблица не выбрана (SetColumnWidth(int columnIndex, int widthPixels))"); }
            _table.Columns[columnIndex].SetWidth(widthPixels, Word.WdRulerStyle.wdAdjustFirstColumn);
            // _table.Columns[columnIndex].SetWidth(widthPixels, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter);
        }



        public void setSelectionToCell(int rowIndex, int columnIndex)
        {
            if (_table == null) { throw new Exception("Ошибка при выборе ячейки в таблице Word, не выбрана текущая таблица."); }

            _currentRange = _table.Cell(rowIndex, columnIndex).Range;
            _selection = new WordSelection(_currentRange, false);
            
            
        }

        public void setSelectionToRow(int index)
        {
            if (_table == null) { throw new Exception("Ошибка при выборе строки в таблице. Возможно, не выбрана сама таблица"); }
            _currentRange = _table.Rows[index].Range;
            _selection = new WordSelection(_currentRange, false);
        }


        public void save(string path)
        {
            Object pathToSaveObj = path;
            _document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj, ref _missingObj);
        }

        public void insertFile(string path)
        {
            if (_currentRange == null) { throw new Exception("Ничего не выбрано"); }
            _currentRange.InsertFile(path);
        }


           

    }
}
