using System;
using System.Data;
using System.Numerics;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Lub1_excel_
{
    public partial class MyExcel : Form
    {
        private int _maxCols = 155;
        private int _maxRows = 35;
        
        // // // Initializing part // /// //
        private void InitializeAllCells()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach(DataGridViewCell cell in row.Cells)
                {
                    InitializeSinglCell(row, cell);
                }
            }
        }
        private void InitializeSinglCell(DataGridViewRow row, DataGridViewCell cell)
        {
            string cellName = char.ToString((char)(65 + cell.ColumnIndex)) + (row.Index + 1);
            Cell cell1 = new Cell(cellName);
            cell.Tag = cell1;
            cell.Value = "";
        }
        private string IntToAlph(int a)
        {
            const int num = 26;
            int _a = a;
            string colHeadText = "";

            if (_a >= num)
            {
                colHeadText += IntToAlph(_a / num - 1);
            }
            colHeadText += char.ToString((char)((int)'A' + _a % num));

            return colHeadText;
        }
        private void InitializeDataGridView()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ColumnCount = _maxCols;
            dataGridView1.RowCount = _maxRows;
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.HeaderText = IntToAlph(col.Index);
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = "" + (row.Index + 1);
            }
        }

        // // // Main method // // //
        public MyExcel()
        {
            InitializeComponent();
            InitializeDataGridView();
            InitializeAllCells();
            CellManager.Instance.DataGridView = dataGridView1;
        }

        // // // Main part // // //
        private void minusProblem(Cell cell)
        {
            string[] wordAr = cell.strValue.Split('-');
            string lit = null;

            if (wordAr.Length > 1)
            {
                cell.Formula = wordAr[0];
                for (int ii = 1; ii < wordAr.Length; ii++)
                {
                    foreach (char c in wordAr[ii])
                    {
                        if (c == '(' || c == '+' || c == '*' || c == '/' || c == ':' || c == '%')
                            break;
                        lit += c;
                    }
                    if (lit == null)
                        cell.Formula += string.Format("-{0}", wordAr[ii]);
                    else
                    {
                        cell.Formula += string.Format("-({0})", lit);
                        char[] chLit = lit.ToCharArray();
                        cell.Formula += wordAr[ii].TrimStart(chLit);
                        lit = null;
                    }
                }
                cell.Formula = cell.Formula.TrimStart('=');
            }
            else
                cell.Formula = cell.strValue.TrimStart('=');
        }
        private void CulcExpr(Cell cell, int col, int row)
        {
            minusProblem(cell);
            cell.doubValue = Calculator.Evaluate(cell.Formula);
            if (cell.strValue[0] == '=') // якщо нема знака дорівнює, то вважаємо, що користувач хоче бачити формулу
                dataGridView1[col, row].Value = cell.doubValue;
            else
                dataGridView1[col, row].Value = cell.strValue;

            int i = 0;
            while (CellManager.Instance[i] != null)
            {
                Cell cellParent = CellManager.Instance[i++];
                cellParent[cellParent.IndexParent++] = cell;
                CellManager.Instance[cellParent.IndexParent - 1] = null;
            }
            CellManager.Instance.index = 0;
        }
        private void ParentCheck(Cell cell)
        {
            for (int i = 0; cell[i] != null; i++)
            {
                cell[i].doubValue = Calculator.Evaluate(cell[i].Formula);
                int _row = Int32.Parse(cell[i].Name.TrimStart(cell[i].Name[0])) - 1;
                int _col = (char)(cell[i].Name[0] - (int)'A');
                if (cell[i].strValue[0] == '=') // якщо нема знака дорівнює, то вважаємо, що користувач хоче бачити формулу
                    dataGridView1[_col, _row].Value = cell[i].doubValue;
                else
                    dataGridView1[_col, _row].Value = cell[i].strValue;
                ParentCheck(cell[i]);

            }
            for (int i = 0; CellManager.Instance[i] != null; i++)
            {
                CellManager.Instance[i] = null;
            }
            CellManager.Instance.index = 0;
        }
        private void afterEdit(Cell cell, int col, int row)
        {
            // Видаляю усі вайтспейси перед записом
            cell.strValue = (cell.strValue != null && cell.strValue != "") ? cell.strValue.TrimStart(' ') : null;

            if (cell.strValue != null && cell.strValue != "")
            {
                foreach (char c in cell.strValue) // перевірка на службовий символ
                {
                    if (c == '|')
                        throw new ReservedSymbolException();
                }

                CulcExpr(cell, col, row);
            }
            else
            {
                cell.Formula = null;
                cell.doubValue = 0;
            }

            ParentCheck(cell);
        }
        private void ErrorCheck(Cell cell, int col, int row)
        {
            try
            {
                afterEdit(cell, col, row);
            }
            catch (OutOfGridException /*ex*/)
            {
                dataGridView1[col, row].Value = "#ERROR: Out of grid";
            }
            catch (DivideByZeroException /*ex*/)
            {
                dataGridView1[col, row].Value = "#ERROR: Division by zero";
            }
            catch (IndexOutOfRangeException /*ex*/)
            {
                dataGridView1[col, row].Value = "#ERROR: Don't do a cycle";
            }
            catch (ReservedSymbolException /*ex*/)
            {
                dataGridView1[col, row].Value = "#ERROR: Don't use '|' symbol, it is reserved";
            }
            catch (NotFormulaException /*ex*/)
            {
                dataGridView1[col, row].Value = "#ERROR: Check the correctness of entering formulas";
            }
        }
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Cell cell = (Cell)dataGridView1[e.ColumnIndex, e.RowIndex].Tag;
            if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null)
            {
                cell.strValue = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                ErrorCheck(cell, e.ColumnIndex, e.RowIndex);
            }
            else
                cell.strValue = null;
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            Cell cell = (Cell)dataGridView1[e.ColumnIndex, e.RowIndex].Tag;
            label2.Text = cell.Name;
            textBox1.Text = cell.strValue;
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int _row = Int32.Parse(label2.Text.TrimStart(label2.Text[0]))-1;
                int _col = (char)(label2.Text[0] - 65);

                Cell cell = (Cell)dataGridView1[_col, _row].Tag;
                cell.strValue = textBox1.Text;
                dataGridView1[_col,_row].Value = cell.strValue;
                ErrorCheck(cell, _col, _row);
            }
        }
        private void deleteCell(int col, int row)
        {

            Cell cell = (Cell)dataGridView1[col, row].Tag;
            cell.strValue = null;
            cell.Formula = null;
            cell.doubValue = 0;
            ErrorCheck(cell, col, row);
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                int row = Int32.Parse(label2.Text.TrimStart(label2.Text[0])) - 1;
                int col = (char)(label2.Text[0] - (int)'A');

                dataGridView1[col, row].Value = null;
                textBox1.Text = null;
                deleteCell(col, row);
            }
        }


        // // // Grid editing part // // //
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Cell cell = (Cell)dataGridView1[e.ColumnIndex, e.RowIndex].Tag;
            dataGridView1[e.ColumnIndex, e.RowIndex].Value = cell.strValue;
        }
        private void AddRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(new DataGridViewRow());

            DataGridViewRow newRow = dataGridView1.Rows[dataGridView1.RowCount - 1];
            newRow.HeaderCell.Value = "" + (newRow.Index + 1);
            foreach (DataGridViewCell cell in newRow.Cells)
            {
                InitializeSinglCell(newRow, cell);
            }
            ++_maxRows;
        }
        private void AddColumToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Add(new DataGridViewColumn(dataGridView1.Rows[0].Cells[0]));

            DataGridViewColumn newCol = dataGridView1.Columns[dataGridView1.ColumnCount - 1];
            newCol.HeaderText = char.ToString((char)(65 + newCol.Index));

            foreach (DataGridViewRow _Row in dataGridView1.Rows)
            {
                InitializeSinglCell(_Row, _Row.Cells[dataGridView1.ColumnCount - 1]);
            }
            ++_maxCols;
        }
        private void DeleteRowToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 1)
            {
                MessageBox.Show("Не можна видалити останній рядок");
                return;
            }
            foreach (DataGridViewCell cell in dataGridView1.Rows[dataGridView1.RowCount - 1].Cells)
            {
                Cell cell1 = (Cell)dataGridView1[cell.ColumnIndex, cell.RowIndex].Tag;
                for (int i = 0; cell1[i] != null; i++)
                {
                    int _row = Int32.Parse(cell1[i].Name.TrimStart(cell1[i].Name[0])) - 1;
                    int _col = (char)(cell1[i].Name[0] - (int)'A');
                    dataGridView1[_col, _row].Value = "ERROR: Out of grid";
                    ParentCheck(cell1[i]);
                }

            }
            dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 1);
            --_maxRows;
        }
        private void DeleteColToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Columns.Count == 1)
            {
                MessageBox.Show("Не можна видалити останній стовпчик");
                return;
            }
            foreach (DataGridViewRow _Row in dataGridView1.Rows)
            {
                int _col = (char)(_Row.Cells[dataGridView1.ColumnCount - 1].ColumnIndex);
                int _row = (char)(_Row.Cells[dataGridView1.ColumnCount - 1].RowIndex);
                Cell cell1 = (Cell)dataGridView1[_col, _row].Tag;
                for (int i = 0; cell1[i] != null; i++)
                {
                    int row = Int32.Parse(cell1[i].Name.TrimStart(cell1[i].Name[0])) - 1;
                    int col = (char)(cell1[i].Name[0] - (int)'A');
                    dataGridView1[col, row].Value = "ERROR: Out of grid";
                    ParentCheck(cell1[i]);
                }
            }
            dataGridView1.Columns.RemoveAt(dataGridView1.Columns.Count - 1);
            --_maxCols;
        }
        private void clearTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InitializeAllCells();
        }


        // // //   Save/Open part // // //
        private string currentPath = null;
        private void SaveDataGridView(string filePath)
        {
            String s, st;
            StreamWriter srw = new StreamWriter(filePath);
            
            string[] _list = filePath.Split("\\");
            string _fileName = _list[_list.Length-1];
            this.Text = _fileName;
            st = "";
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Cell cell = (Cell)dataGridView1.Rows[i].Cells[j].Tag;
                    s = cell.strValue == null ? "": cell.strValue;
                    st += s + "|";
                }
                srw.WriteLine(st);
                st = "";
            }
            srw.Close();
        }
        private void SaveFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (currentPath != null)
            {
                SaveDataGridView(currentPath);               
            }
            else
            {
                SaveFileAsToolStripMenuItem_Click(sender, e);
            }
        }
        private void SaveFileAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SaveDataGridView((string)saveFileDialog1.FileName);
                currentPath = (string)saveFileDialog1.FileName;
            }
        }
        private void PreparingForTheOpening(string currentPath)
        {
            String s = null, st;
            StreamReader strRead = new StreamReader(currentPath);

            _maxCols = 0;
            _maxRows = 0;
            while ((st = strRead.ReadLine()) != null)
            {
                string[] arr = st.Split("|");
                _maxCols = arr.Length-1;
                ++_maxRows;
            }
            InitializeDataGridView();
            InitializeAllCells();
            strRead.Close();

        }
        private void OpenFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileToolStripMenuItem_Click(sender, e); // зберігаємо попередній файл

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String st;
                StreamReader strRead = new StreamReader((string)openFileDialog1.FileName);
                currentPath = (string)openFileDialog1.FileName;
                
                string[] _list = currentPath.Split("\\");
                string _fileName = _list[_list.Length - 1];
                this.Text = _fileName;

                PreparingForTheOpening(currentPath); // Знаходження кількості col\row та ініціалізація grid
                
                string[] arr;
                int row = 0, col = 0;              
                while ((st = strRead.ReadLine()) != null)
                {
                    arr = st.Split("|");
                    foreach(string c in arr)
                    {
                        if (c != null && c != "") 
                        {
                            dataGridView1[col, row].Value = c;
                            Cell cell = (Cell)dataGridView1[col, row].Tag;
                            cell.strValue = c;
                            ErrorCheck(cell, col, row);
                        }
                        col++;
                    }
                    col = 0;
                    row++;
                }
                strRead.Close();
            }
        }

        // // // Info/Close form button // // //
        private void informationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("* Якщо ви бажаєте ввести формулу, то поставте перед виразом знак '=':\n" +
                            " ( =2+3/4 )\n" +
                            "* У даній версії програми доступні такі дії: \n" +
                            "\t1) утворення складних формул (поставити '=' перед записом) з \n" +
                            "\t\t 1. додаванням ( '+' ),\n" +
                            "\t\t 2. відніманням ( '-' ), \n" +
                            "\t\t 3. множенням ( '*' ),\n" +
                            "\t\t 4. діленням (div - '/', mod - '%', ділення з остачею- ':' )\n" +
                            "\t   простих чисел та посилань на клітинки (=(А1+А2)*B3)\n " +
                            "\t2) запис довільних даних у таблицю (не ставити '='):\n" +
                            "\t   ( 2+3), ( довільний текст )\n" +
                            "* При посиланні ня клітинку, результат буде успішним, якщо виконані наступні умови:\n" +
                            "\t 1) У клітинці, на яку посилаються, вписане дійсне число\n" +
                            "\t 2) У клітинці, на яку посилаються, вписана формула (не важливо,\n" +
                            "\t стоїть '=' чи ні, головне, щоб ця формула була записана\n" +
                            "\t правильно)\n" +
                            "\t 3) У клітинці, на яку посилаються, вписана назва іншої клітинки,\n" +
                            "\t чи формула, що містить цю назву (не важливо, стоїть '=' чи ні,\n" +
                            "\t головне, щоб ця формула була записана правильно)");

        }
        private void MyExcel_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveFileToolStripMenuItem_Click(sender,e);
        }
        private void newFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileToolStripMenuItem_Click(sender, e);
            this.Text = "MyExcel";
            currentPath = null;
            _maxCols = 15;
            _maxRows = 35;
            InitializeDataGridView();
            InitializeAllCells();
            //CellManager.Instance.DataGridView = dataGridView1;
        }
    }

    // // // Some useful additional classes // // //
    class Cell
    {
        private Cell[] _child = new Cell[50];
        private double _doubValue;
        private string _strValue;
        private string _formula;// 5-4
        private string _name;
        private int indexParent;
        // // //
        public double doubValue
        {
            get { return _doubValue; }
            set { _doubValue = value; }
        }
        public string strValue
        {
            get { return _strValue; }
            set { _strValue = value; }
        }
        public string Formula
        {
            get { return _formula; }
            set { _formula = value; }
        }
        public string Name
        {
            get { return _name; }
        }
        public int IndexParent
        {
            get { return indexParent; }
            set { indexParent = value; }
        }
        public Cell this[int i]
        {
            get { return _child[i]; }
            set { _child[i] = value; }
        }

        // // //
        public Cell()
        {
            _name = null;
            _doubValue = 0;
            _strValue = null;
            _formula = null;
            indexParent = 0;
        }
        public Cell( string name)
        {
            _formula = null;
            _strValue = null;
            _name = name;
            _doubValue = 0;
            indexParent = 0;
        }
    }
    class CellManager
    {
        private static DataGridView _dataGridView;
        private static CellManager _instance;
        private static int _indexParent = 0;
        private Cell[] _parent = new Cell[50];

        // // //
        public Cell this[int i]
        {
            get { return _parent[i]; }
            set { _parent[i] = value; }
        }
        public int index
        {
            get { return (int)_indexParent; }
            set { _indexParent = value; }
        }
        public DataGridView DataGridView 
        { 
            set { _dataGridView = value; } 
            get { return _dataGridView; } 
        }
        public static CellManager Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new CellManager();
                return _instance;
            }
        }
        // // //
        public bool outOfArrCheck()
        {
            if (_indexParent == 49)
                return true;
            return false;
        }
        public Cell GetCell(string cellName)
        {
            int _row = 0;
            int _col = 0;
            string _strCol = cellName;
            int ii = 1;
            for (int i = cellName.Length - 1; int.TryParse("" + cellName[i], out int _a); i--)
            {
                _row += _a*ii;
                ii *= 10;
                _strCol = _strCol.TrimEnd((char)(_a + 48));
            }

            ii = 0;
            for (int i = _strCol.Length - 1; i > -1; i--)
            {
                ii = (int)Math.Pow(26, (_strCol.Length - (i + 1)) );
                _col += (char)(_strCol[i] - 64) * ii; // рахуємо за принципом 26^0 + 26^1 + 26^2 ...
            }

            if ((_dataGridView.ColumnCount < _col) || (_dataGridView.RowCount < _row))
                    throw new OutOfGridException();

            Cell cell = (Cell)_dataGridView[--_col, --_row].Tag;
            return cell;

        }

    }

    // // // Exceptions // // //
    class OutOfGridException : System.Exception
    {
        public OutOfGridException() { }
    }
    class ReservedSymbolException : Exception
    {
        public ReservedSymbolException() { }
    }
    class NotFormulaException : Exception
    {
        public NotFormulaException() { }
    }

    /*  class cacheItem
      {
          public string name;
          private Cell[] _child = new Cell[50];
          public int index;
          public cacheItem()
          {
              //_child = new Cell[0];
              name = null;
              index = 0;
          }
          public Cell Child
          {
              set{_child[index++] = value;}
          }
      }
      class Cache
      {
          private static Cache _instanceCache;
          private static cacheItem[] _item = new cacheItem[50];
          private static int _index;
          private Cache()
          {
              _index = 0;
          }
          public static Cache InstanceCache
          {
              get
              {
                  if (_instanceCache == null)
                      _instanceCache = new Cache();
                  return _instanceCache;
              }
          }
          public int Index
          {
              get { return _index; }
          }
          public void addCacheItem(string name)
          {
              _item[_index++].name = name;
          }
          public void addCacheItem(Cell cell)
          {
              //_item[0]._child[0] = cell;
              _item[_index].Child = cell;
              //_item[_index]._child[_item[_index].index++] = cell;
          }
          public string getCacheItemName(int ii)
          {
              return "";
              //return _item[ii].name;
          }
          public Cell getCacheItem()
          {
              return null;
             // return _item[_index]._child[_item[_index--].index--] != null ? _item[_index]._child[_item[_index--].index--]:null ;
          }




      }*/
}