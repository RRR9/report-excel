using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportExcel
{
    public partial class MainForm : Form
    {
        private Excel.Application _excelApp;
        private Excel.Workbook _excelWB;
        private Excel.Worksheet _excelSheet;
        private Excel.Range _excelRange;

        private HashSet<int> _existKeys;

        private Dictionary<int, List<Pharmacy>> _result;
        private SortedDictionary<int, Dictionary<string, List<Pharmacy>>> _mainResult;

        private Dictionary<string, ExcelFile> _excelFiles;
        private Dictionary<string, List<Pharmacy>> _excelData;
        private Dictionary<string, double> _priceDiscount;
        private HashSet<string> _searchWords;
        private int _currentDistance;

        private void CheckObjectsAfterCloseForm(object sender, FormClosedEventArgs e)
        {
            _excelWB?.Close();
            _excelApp.Quit();
        }

        public MainForm()
        {
            _excelApp = new Excel.Application();

            FormClosed += CheckObjectsAfterCloseForm;

            MaximizeBox = false;

            _priceDiscount = new Dictionary<string, double>();
            _existKeys = new HashSet<int>();
            _result = new Dictionary<int, List<Pharmacy>>();
            _excelFiles = new Dictionary<string, ExcelFile>();
            _excelData = new Dictionary<string, List<Pharmacy>>();
            _searchWords = new HashSet<string>();
            _mainResult = new SortedDictionary<int, Dictionary<string, List<Pharmacy>>>();

            _currentDistance = 0;
            InitializeComponent();
        }

        private string FindDistance(string[] words, string search, out int distance)
        {
            int minDistance = int.MaxValue;
            string findWord = null;
            search = search.ToLower();
            foreach(string w in words)
            {
                string word = w.ToLower();
                if(word == "упак" || word == "амп" || word == "амп.")
                {
                    continue;
                }
                int n = word.Length + 1;
                int m = search.Length + 1;
                int[,] d = new int[n, m];
                d[0, 0] = 0;
                for(int j = 1; j < m; ++j)
                {
                    d[0, j] = d[0, j - 1] + 1;
                }
                for (int i = 1; i < n; ++i)
                {
                    d[i, 0] = d[i - 1, 0] + 1;
                    for (int j = 1; j < m; ++j)
                    {
                        if(word[i - 1] != search[j - 1])
                        {
                            d[i, j] = Math.Min(d[i - 1, j] + 1, Math.Min(d[i, j - 1] + 1, d[i - 1, j - 1] + 1));
                        }
                        else
                        {
                            d[i, j] = d[i - 1, j - 1];
                        }
                    }
                }
                if(d[n - 1, m - 1] < minDistance)
                {
                    minDistance = d[n - 1, m - 1];
                    findWord = word;
                }
            }

            distance = minDistance;
            return findWord;
        }

        private void AddSearchingWord_Button_Click(object sender, EventArgs e)
        {
            if(searchingWord_TextBox.Text.Length > 0)
            {
                string s = searchingWord_TextBox.Text;
                searchingWords_ListBox.Items.Add(s);
                _searchWords.Add(s);
                searchingWord_TextBox.Text = "";
            }
        }

        private void DeleteSearchingWord_Button_Click(object sender, EventArgs e)
        {
            string s = (string)searchingWords_ListBox.SelectedItem;
            _searchWords.Remove(s);
            searchingWords_ListBox.Items.Remove(s);
        }

        private void AddData_Button_Click(object sender, EventArgs e)
        {
            _excelData.Clear();

            foreach (var item in excelFiles_ListBox.Items)
            {
                string fullFileName = (string)item;
                _excelWB = _excelApp.Workbooks.Open(fullFileName);
                _excelSheet = (Excel.Worksheet)_excelWB.ActiveSheet;
                string fileName = Path.GetFileName(fullFileName);

                if(!_excelFiles.TryGetValue(fileName, out ExcelFile excelFile))
                {
                    MessageBox.Show($"Excel файл НЕ ЗАГРУЖЕН в программу: {fileName}");
                    continue;
                }

                dynamic name;
                dynamic price;
                if(excelFile.RowName == -1 || excelFile.ColumnName == -1 || excelFile.RowPrice == -1 || excelFile.ColumnPrice == -1)
                {
                    MessageBox.Show($"НЕ ЗАДАНЫ начальные строки и столбцы для: {item as string}");
                    continue;
                }
                int rowName = excelFile.RowName, columnName = excelFile.ColumnName;
                int rowPrice = excelFile.RowPrice, columnPrice = excelFile.ColumnPrice;
                _excelData.Add(fileName, new List<Pharmacy>());

                while ((name = (_excelSheet.Cells[rowName, columnName] as Excel.Range).Value) != null)
                {
                    price = (_excelSheet.Cells[rowPrice, columnPrice] as Excel.Range).Value;
                    _excelData[fileName].Add(new Pharmacy(Convert.ToString(name), Convert.ToString(price)));

                    ++rowName;
                    ++rowPrice;
                }
                _excelWB.Close();
            }

        }

        private void AddExcelFile_Button_Click(object sender, EventArgs e)
        {
            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            if (openFileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            excelFiles_ListBox.Items.Add(openFileDialog.FileName);
            string fileName = Path.GetFileName(openFileDialog.FileName);
            ExcelFile excelFile = new ExcelFile();
            excelFile.FullfileName = openFileDialog.FileName;
            _excelFiles.Add(fileName, excelFile);
        }

        private void SavePosition_Button_Click(object sender, EventArgs e)
        {
            string fileName = (string)excelFiles_ListBox.SelectedItem;
            fileName = Path.GetFileName(fileName);
            if(fileName is null || fileName.Length == 0)
            {
                MessageBox.Show("Имя файла пустое!");
                return;
            }

            Regex regex = new Regex(@"[^0-9]");
            string value = rowName_TextBox.Text;

            if(regex.IsMatch(value))
            {
                MessageBox.Show("Неправильное значение строки для 'Наименования продукта'");
                return;
            }

            value = columnName_TextBox.Text;
            if (regex.IsMatch(value))
            {
                MessageBox.Show("Неправильное значение столбца для 'Наименования продукта'");
                return;
            }

            value = rowPrice_TextBox.Text;
            if (regex.IsMatch(value))
            {
                MessageBox.Show("Неправильное значение строки для 'Цена продукта'");
                return;
            }

            value = columnPrice_TextBox.Text;
            if (regex.IsMatch(value))
            {
                MessageBox.Show("Неправильное значение столбца для 'Цена продукта'");
                return;
            }

            if(rowName_TextBox.Text != rowPrice_TextBox.Text)
            {
                MessageBox.Show("Наименования продукта и его цена должны находиться на одной строке!");
                return;
            }

            if(_excelFiles.TryGetValue(fileName, out ExcelFile excelFile))
            {
                excelFile.RowName = Convert.ToInt32(rowName_TextBox.Text);
                excelFile.ColumnName = Convert.ToInt32(columnName_TextBox.Text);
                excelFile.RowPrice = Convert.ToInt32(rowPrice_TextBox.Text);
                excelFile.ColumnPrice = Convert.ToInt32(columnPrice_TextBox.Text);
            }
            else
            {
                MessageBox.Show("Файл в программе не существует!\n Добавьте excel файл!");
                return;
            }

            value = discount_TextBox.Text;
            if(regex.IsMatch(value))
            {
                MessageBox.Show("Скидка от поставщика указанна неверно!");
            }
            else
            {
                if(!_priceDiscount.ContainsKey(fileName))
                {
                    _priceDiscount.Add(fileName, Convert.ToDouble(value));
                } 
                else
                {
                    _priceDiscount[fileName] = Convert.ToDouble(value);
                }
            }
        }

        private void ExcelFiles_ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string fileName = (string)excelFiles_ListBox.SelectedItem;
            fileName = Path.GetFileName(fileName);
            if(fileName is null)
            {
                return;
            }
            _excelFiles.TryGetValue(fileName, out ExcelFile excelFile);
            rowName_TextBox.Text = Convert.ToString(excelFile.RowName);
            columnName_TextBox.Text = Convert.ToString(excelFile.ColumnName);
            rowPrice_TextBox.Text = Convert.ToString(excelFile.RowPrice);
            columnPrice_TextBox.Text = Convert.ToString(excelFile.ColumnPrice);

            _priceDiscount.TryGetValue(fileName, out double priceDiscount);
            discount_TextBox.Text = Convert.ToString(priceDiscount);
        }

        private void ExportData_Button_Click(object sender, EventArgs e)
        {
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.FileName = "Отчёт_сравнения_цен_" + DateTime.Now.ToString("yyyy-MM-dd");
            saveFileDialog.CheckPathExists = true;
            if(saveFileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            string pathFile = saveFileDialog.FileName;
            string fileName = Path.GetFileName(pathFile);

            int indexRow = 2, indexColumn = 1;

            _excelApp.Visible = false;
            _excelWB = _excelApp.Workbooks.Add(Type.Missing);
            _excelSheet = _excelWB.ActiveSheet;

            _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn], _excelSheet.Cells[indexRow, indexColumn + 1]];
            _excelRange.Columns.ColumnWidth = 2.7;

            indexColumn += 2;

            int saveIndexRowPos;
            int saveIndexColumnPos;
            int maxBottomRowIndex = 0;

            foreach (var searchWord in _searchWords)
            {
                _mainResult.Clear();

                _excelSheet.Cells[indexRow, indexColumn] = "Заданное в поиске";

                _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn], _excelSheet.Cells[indexRow, indexColumn]];

                _excelRange.Font.Bold = true;
                _excelRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _excelRange.EntireColumn.AutoFit();

                indexRow += 1;

                _excelSheet.Cells[indexRow, indexColumn] = searchWord;
                _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn], _excelSheet.Cells[indexRow, indexColumn]];
                _excelRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                _excelRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                indexRow += 3;
                saveIndexRowPos = indexRow;
                saveIndexColumnPos = indexColumn;

                foreach (var currentExcel in _excelData)
                {
                    foreach (var listItem in currentExcel.Value)
                    {
                        string[] words = listItem.Name.Split(new char[] { ' ', '-', ',' });
                        string keyWord = FindDistance(words, searchWord, out int distance);
                        string excelFileName = currentExcel.Key;

                        if(!_mainResult.ContainsKey(distance))
                        {
                            _mainResult.Add(distance, new Dictionary<string, List<Pharmacy>>());
                        }

                        if(!_mainResult[distance].ContainsKey(excelFileName))
                        {
                            _mainResult[distance].Add(excelFileName, new List<Pharmacy>());
                        }

                        _mainResult[distance][excelFileName].Add(new Pharmacy(listItem.Name, listItem.Price));

                    }
                }

                HashSet<string> usedExcelFile = new HashSet<string>();
                foreach (var value in _mainResult)
                {
                    foreach (var value2 in value.Value)
                    {
                        if (!usedExcelFile.Contains(value2.Key))
                        {
                            _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn], _excelSheet.Cells[indexRow, indexColumn + 1]];
                            _excelRange.Merge();
                            _excelRange.Value2 = value2.Key;

                            _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn], _excelSheet.Cells[indexRow, indexColumn]];
                            _excelRange.EntireColumn.AutoFit();

                            _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn - 1], _excelSheet.Cells[indexRow, indexColumn + 1]];
                            _excelRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            _excelRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            _excelRange.Font.Bold = true;

                            indexRow += 1;

                            _excelSheet.Cells[indexRow, indexColumn] = "Номенклатура";
                            _excelSheet.Cells[indexRow, indexColumn + 1] = "Цена";
                            _excelSheet.Cells[indexRow, indexColumn + 2] = "Цена со скидкой";
                            _excelSheet.Cells[indexRow, indexColumn - 1] = "№";

                            _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn - 1], _excelSheet.Cells[indexRow, indexColumn + 2]];
                            _excelRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            _excelRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            _excelRange.Font.Bold = true;

                            _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn + 2], _excelSheet.Cells[indexRow, indexColumn + 2]];
                            _excelRange.Columns.ColumnWidth = 16;

                            indexRow += 1;
                            maxBottomRowIndex = Math.Max(indexRow, maxBottomRowIndex);

                            int countItem = 1;
                            double priceDiscount;
                            foreach (var value3 in value2.Value)
                            {
                                _excelSheet.Cells[indexRow, indexColumn] = value3.Name;
                                _excelSheet.Cells[indexRow, indexColumn + 1] = value3.Price;
                                _excelSheet.Cells[indexRow, indexColumn - 1] = countItem;

                                priceDiscount = Convert.ToDouble(value3.Price) * (100 - _priceDiscount[value2.Key]) / 100;
                                _excelSheet.Cells[indexRow, indexColumn + 2] = priceDiscount;

                                _excelRange = _excelSheet.Range[_excelSheet.Cells[indexRow, indexColumn - 1], _excelSheet.Cells[indexRow, indexColumn + 2]];
                                _excelRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                indexRow += 1;
                                countItem += 1;
                                maxBottomRowIndex = Math.Max(indexRow, maxBottomRowIndex);
                            }
                            usedExcelFile.Add(value2.Key);
                            indexRow = saveIndexRowPos;
                            indexColumn += 5;
                        }
                    }
                }

                indexRow = maxBottomRowIndex + 5;
                indexColumn = saveIndexColumnPos;
            }
            _excelWB.SaveAs(pathFile);
            _excelWB.Close();
        }

    }
}
