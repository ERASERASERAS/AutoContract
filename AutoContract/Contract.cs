using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace AutoContract
{
    class Contract
    {
        private static string _contractType;
        private int _numberContract;
        private string _date;
        private static double _cost;
        private string _fullFirmName;
        private string _shortFirmName;
        private string _customerName;
        private string _payerName;
        private string _directorName;
        private string _complDate;
        private string _inn;
        private string _kpp;
        private string _ogrn;
        private string _checkingAccount;
        private string _correspondentAccount;
        private string _phoneNumber;
        private string _fax;
        private string _email;
        private int _numberAccount;
        private static bool hasAdvance;
        private static string advance;
        
        private static string pathToFolderWithFiles;

        private static LinkedList<string> headersOfEstimatedCostForm = new LinkedList<string>();

        public static void addToHeadersOfEstimatedCostForm(string header)
        {
            headersOfEstimatedCostForm.AddLast(header);
        }

        public static LinkedList<string> getHeadersOfEstimatedCostForm()
        {
            return headersOfEstimatedCostForm;
        }

        public static string PATHTOFOLDERWITHFILES
        {
            get { return pathToFolderWithFiles; }
            set { pathToFolderWithFiles = value; }
        }
        

        public static string getElementOfHeadersOfEstimatedCostForm(int index)
        {
            string value;
            if (index == 1)
            {
                value = headersOfEstimatedCostForm.First.Value;
            }
            else
            {
                LinkedListNode<string> node = headersOfEstimatedCostForm.First.Next;
                for (int i = 2; i < index; i++)
                {
                    node = node.Next;
                }
                value = node.Value;
            }
            return value;
        }

        public static string CONTRACTTYPE
        {
            get { return _contractType; }
            set { _contractType = value; }
        }

        public int NUMBERCONTRACT
        {
            get { return _numberContract; }
            set { _numberContract = value; }
        }

        public static string ADVANCE
        {
            get { return ADVANCE; }
            set { ADVANCE = value; }
        }

        public string DATE
        {
            get { return _date; }
            set { _date = value; }
        }

        public static double COST
        {
            get { return _cost; }
            set { _cost = value; }
        }

        public string FULLFIRMNAME
        {
            get { return _fullFirmName; }
            set { _fullFirmName = value; }
        }

        public string SHORTFIRMNAME
        {
            get { return _shortFirmName; }
            set { _shortFirmName = value; }
        }

        public string CUSTOMERNAME
        {
            get { return _customerName; }
            set { _customerName = value; }
        }

        public string PAYERNAME
        {
            get { return _payerName; }
            set { _payerName = value; }
        }

        public string DIRECTORNAME
        {
            get { return _directorName; }
            set { _directorName = value; }
        }

        public string COMPLDATE
        {
            get { return _complDate; }
            set { _complDate = value;}
        }

        public string INN
        {
            get { return _inn; }
            set { _inn = value; }
        }

        public string KPP
        {
            get { return _kpp; }
            set { _kpp = value; }
        }

        public string OGRN
        {
            get { return _ogrn; }
            set { _ogrn = value; }
        }

        public string CHECKINGACCOUNT
        {
            get { return _checkingAccount; }
            set { _checkingAccount = value; }
        }

        public string CORRESPONDENTACCOUNT
        {
            get { return _correspondentAccount; }
            set { _correspondentAccount = value; }
        }

        public string PHONENUMBER
        {
            get { return _phoneNumber; }
            set { _phoneNumber = value; }
        }

        public string FAX
        {
            get { return _fax; }
            set { _fax = value; }
        }

        public string EMAIL
        {
            get { return _email; }
            set { _email = value; }
        }

        public int  NUMBERACCOUNT
        {
            get { return _numberAccount; }
            set { _numberAccount = value; }
        }

        public static bool HASADVANCE
        {
            get { return hasAdvance; }
            set { hasAdvance = value; }
        }

        public Contract(string contractType, int numberContract, string date, double cost, string fullFirmName, string shortFirmName, string customerName, string payerName, string directorName, string complDate, string inn, string kpp, string ogrn, string checkingAccount, string correspondentAccount, string phoneNumber, string fax, string email, int numberAccount)
        {
            _contractType = contractType;
            _numberContract = numberContract;
            _date = date;
            _cost = cost;
            _fullFirmName = fullFirmName;
            _shortFirmName = shortFirmName;
            _customerName = customerName;
            _payerName = payerName;
            _directorName = directorName;
            _complDate = complDate;
            _inn = inn;
            _kpp = kpp;
            _ogrn = ogrn;
            _checkingAccount = checkingAccount;
            _correspondentAccount = correspondentAccount;
            _phoneNumber = phoneNumber;
            _fax = fax;
            _email = email;
            _numberAccount = numberAccount;
        }

        public void changeDirectorsName(string directorsName)
        {

        }


        public static string[] getPartsOfDate(string date)      //[day,monthsword,year]
        {
            string day = date.Substring(0, 2);
            string monthWord = getMonthsWord(date.Substring(3, 2));
            string year = date.Substring(6, 4);
            string[] res = { day, monthWord, year };
            return res;
        }

        public static string getMonthsWord(string monthNumber)
        {
            switch (monthNumber)
            {
                case "1": return "января"; break;
                case "2": return "февраля"; break;
                case "3": return "марта"; break;
                case "4": return "апреля"; break;
                case "5": return "мая"; break;
                case "6": return "июня"; break;
                case "7": return "июля"; break;
                case "8": return "августа"; break;
                case "9": return "сентября"; break;
                case "10": return "октября"; break;
                case "11": return "ноября"; break;
                case "12": return "декабря"; break;
                default: return "null"; break;
            }
        }

        public static string getShortDirName(string name)
        {
            string shortName = name.Substring(0, name.IndexOf(' ')) + " " + name.Substring(name.IndexOf(' ') + 1 , 1) + ". " + name.Substring(name.LastIndexOf(' ') + 1 , 1) + ".";
            return shortName;
        }

        //различные методы меняющие что-то.

        public static void insertTable(WordMethods wm, string type, int rowsCount, int colsCount)
        {
            switch (type)
            {
                case "worksch":
                    insertWorkSchedulesTable(wm, rowsCount, colsCount);
                    break;
                case "estcost":
                    insertEstimatedCostTable(wm, rowsCount, colsCount);
                    break;
                default:
                    break;
            }
        }

        public static void insertToCell(WordMethods wm, string text, int i, int j)
        {
            wm.setSelectionToCell(i, j);
            wm.Selection.Text = text;
        }


        public static void goToBookmark(WordMethods wm, string bookmarksName)
        {
            wm.goToBookmark(bookmarksName);
        }

        public static void replaceString(WordMethods wm, string label, string newString)
        {
            wm.replace(label, newString);
        }

        private static void insertWorkSchedulesTable(WordMethods wm, int rowsCount, int colsCount)
        {
            wm.goToBookmark("w1");
            wm.insertTable(rowsCount, colsCount, false);
        }

        public static void setAttributesForWorkSchedulesTable(WordMethods wm, int rowsCount, int colsCount) // Counts from dataGridView
        {
            wm.setHeightForRow(2, 15);
            wm.setColumnWidth(1, 40);
            wm.setHeightForRow(1, 45);

            for (int i = 0; i < colsCount; i++)
            {
                wm.setSelectionToCell(1, i + 1);
                wm.Selection.FontSize = 12;
                wm.Selection.FontName = "Times New Roman";
                wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wm.Selection.Aligment = TextAligment.Center;
                wm.setSelectionToCell(2, i + 1);
                wm.Selection.FontSize = 12;
                wm.Selection.FontName = "Times New Roman";
                wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wm.Selection.Aligment = TextAligment.Center;
            }


            for (int i = 0; i < rowsCount - 1; i++)
            {
                for (int j = 0; j < colsCount; j++)
                {
                    wm.setSelectionToCell(i + 3, j + 1);  //wordDoc.SetColumnWidth(1,120);//wordDoc.SetColumnWidth(2, 60)
                    wm.Selection.FontSize = 12;
                    wm.Selection.FontName = "Times New Roman";
                    wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                   wm.Selection.Aligment = TextAligment.Center;

                    // _wm.Selection.setLineSpacing(1f);
                    //_wm.setColumnWidth(1, 300);
                    //_wm.Selection.Aligment = TextAligment.Center; //где-то нужно будет поменять выравнивание                                                                                                                     
                }
            }

            for (int i = 0; i < rowsCount - 1; i++)
            {
                wm.setSelectionToCell(i + 3, 2);
                wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wm.Selection.Aligment = TextAligment.Left;
                wm.setSelectionToCell(i + 3, 5);
                wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wm.Selection.Aligment = TextAligment.Left;
            }
        }

        public static void setAttributesForEstimatedCostTable(WordMethods wm, int rowsCount, int colsCount)
        {
            for (int i = 0; i < colsCount; i++)
            {
                wm.setSelectionToCell(1, i + 1);
                wm.Selection.Text = Contract.getElementOfHeadersOfEstimatedCostForm(i + 1);
                wm.Selection.FontSize = 12;
                wm.Selection.FontName = "Times New Roman";
                wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                wm.Selection.Aligment = TextAligment.Center;
            }

            for (int i = 0; i < rowsCount - 1; i++)
            {
                for (int j = 0; j < colsCount; j++)
                {
                    wm.setSelectionToCell(i + 2, j + 1);  //wordDoc.SetColumnWidth(1,120);//wordDoc.SetColumnWidth(2, 60)
                    wm.Selection.FontSize = 12;
                    wm.Selection.FontName = "Times New Roman";
                    wm.Selection.VertAl = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    wm.Selection.Aligment = TextAligment.Center;

                }
            }

            //for (int i = 0; i < dataGridView1.RowCount; i++)
            //{
            //    wm.setWidthForCell(i + 1, 1, 40);
            //    wm.setWidthForCell(i + 1, 3, 200);
            //}


        }

        private static void insertEstimatedCostTable(WordMethods wm, int rowsCount, int colsCount)
        {
            wm.goToBookmark("w2");
            wm.insertTable(rowsCount, colsCount, true);
        }

        public static void insertFileCalculateEstCost(WordMethods wm)
        {
            wm.goToBookmark("estcostfile");
            wm.insertFile(pathToFolderWithFiles + "est_cost.doc");
        }

        public static void insertFileWithAdvance(WordMethods wm)
        {
            
        }


        public static void setDocumentsVisible(WordMethods wm, bool value)
        {
            wm.Visible = value;
        }


    }
}
