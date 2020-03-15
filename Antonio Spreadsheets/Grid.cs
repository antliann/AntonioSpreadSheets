using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using Antlr4.Runtime.Misc;

namespace Antonio_Spreadsheets
{
    public class Grid
    {
        private const int initColCount = 12;
        private const int initRowCount = 21;
        public int current_cols;
        public int current_rows;

        private Alphabet ltr_sys = new Alphabet();
        public Dictionary<string, string> dictionary = new Dictionary<string, string>();
        public List<List<Cell>> grid = new List<List<Cell>>();

        public Grid()
        {
            current_cols = initColCount;
            current_rows = initRowCount;

            for (int i = 0; i < initRowCount; ++i)
            {
                List<Cell> row = new List<Cell>();
                for (int j = 0; j < initColCount; ++j)
                {
                    string name = ltr_sys.ToLTR(j) + i.ToString();
                    row.Add(new Cell(name, i, j));
                    dictionary.Add(name, "");
                }
                grid.Add(row);
            }
        }

        public void SetGrid(int row, int col)
        {
            Clear();

            current_cols = col;
            current_rows = row;

            for (int i = 0; i < current_rows; ++i)
            {
                List<Cell> newRow = new List<Cell>();
                for (int j = 0; j < current_cols; ++j)
                {
                    string name = ltr_sys.ToLTR(j) + i.ToString();
                    newRow.Add(new Cell(name, i, j));
                    dictionary.Add(name, "");
                }
                grid.Add(newRow);
            }
        }

        public string FullName(int col, int row)
        {
            return ltr_sys.ToLTR(col) + row.ToString();
        }

        public void RefreshReferences()
        {
            foreach (List<Cell> list in grid)
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                        cell.referencesFromThis.Clear();
                    if (cell.new_referencesFromThis != null)
                        cell.new_referencesFromThis.Clear();

                    if (cell.Expression == "")
                        continue;
                    string new_expression = cell.Expression;
                    if (cell.Expression[0] == '=')
                    {
                        new_expression = ConvertReferences(cell.RowIndex, cell.ColIndex, cell.Expression);
                        cell.referencesFromThis.AddRange(cell.new_referencesFromThis);
                    }
                }
        }

        public void ChangeCellWithAllPointers(int row, int col, string expression, DataGridView table)
        {
            grid[row][col].DelPointersAndReferences();
            grid[row][col].Expression = expression;

            //set new references
            grid[row][col].new_referencesFromThis.Clear();
            string value = expression;

            if (expression != "")
            {
                if (expression[0] != '=')
                {
                    grid[row][col].Value = expression;
                    dictionary[FullName(col, row)] = expression;
                    foreach (Cell cell in grid[row][col].pointersToThis)
                    {
                        RefreshCellAndPointers(cell, table);
                    }
                    return;
                }
            }

            string new_expression = ConvertReferences(row, col, expression);
            if (new_expression != "")
                new_expression = new_expression.Remove(0, 1);

            // чекаем на циклы
            if (!grid[row][col].CheckForCycle(grid[row][col].new_referencesFromThis))
            {
                grid[row][col].Value = "ЦИКЛ";
                //grid[row][col].Expression = "ЦИКЛ";
                MessageBox.Show("Клітини циклічно посилаються одна на одну. Змініть вираз!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dictionary[FullName(col, row)] = expression;                //dobavil
                foreach (Cell cell in grid[row][col].pointersToThis)
                {
                    RefreshCellAndPointers(cell, table);
                }

                return;
            }
            // добавляет ссылаемое в клетку
            grid[row][col].AddPointersAndReferences();

            // меняет поля класса
            if (expression != "")
            {
                value = Calculate(new_expression);
            }

            if (value == "ПОМИЛКА")                     
            {
                MessageBox.Show("Помилка в клітині " + FullName(col, row) + '!', "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                grid[row][col].Value = value;
                dictionary[FullName(col, row)] = value;
               // foreach (Cell cell in grid[row][col].pointersToThis)
               // {
                //    RefreshCellAndPointers(cell, table);                              // DOBAVIL
               // }
                return;
            }

            grid[row][col].Value = value;
            dictionary[FullName(col, row)] = value;

            foreach (Cell cell in grid[row][col].pointersToThis)
            {
                RefreshCellAndPointers(cell, table);
            }
        }

        public bool RefreshCellAndPointers(Cell cell, DataGridView table)
        {
            cell.new_referencesFromThis.Clear();

            string new_expression = ConvertReferences(cell.RowIndex, cell.ColIndex, cell.Expression);
            new_expression = new_expression.Remove(0, 1);
            string value = Calculate(new_expression);

            if (value == "ПОМИЛКА")                   
            {
                MessageBox.Show("Помилка в клітині " + cell.Index + '!', "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                grid[cell.RowIndex][cell.ColIndex].Value = value;
                dictionary[FullName(cell.ColIndex, cell.RowIndex)] = value;
                table[cell.ColIndex, cell.RowIndex].Value = value;
                foreach (Cell point in cell.pointersToThis)
                RefreshCellAndPointers(point, table);                            //DOBAVIL
                        
                return false;
            }

            grid[cell.RowIndex][cell.ColIndex].Value = value;
            dictionary[FullName(cell.ColIndex, cell.RowIndex)] = value;
            table[cell.ColIndex, cell.RowIndex].Value = value;

            foreach (Cell point in cell.pointersToThis)
                if (!RefreshCellAndPointers(point, table))
                    return false;
            return true;
        }

        public string Calculate(string expression)
        {
            string res = null;
            try
            {
                res = Convert.ToString(Calculator.Evaluate(expression));
                if (res == "∞")
                {
                    res = "НЕСКІНЧЕННІСТЬ";
                    DialogResult mes = MessageBox.Show("Помилка обчислення: ви намагаєтеся поділити на нуль або порахувати дуже велике значення!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return res;
            }
            catch
            {
                return "ПОМИЛКА";
            }
        }

        public string ConvertReferences(int row, int col, string expression)
        {
            string cellPattern = @"[A-Z]+[0-9]+"; //+
            Regex regex = new Regex(cellPattern, RegexOptions.IgnoreCase);
            int[] nums;

            foreach (Match match in regex.Matches(expression))
                if (dictionary.ContainsKey(match.Value))
                {
                    nums = ltr_sys.FromIndex(match.Value);
                    grid[row][col].new_referencesFromThis.Add(grid[nums[1]][nums[0]]); //swap 1 0
                }

            MatchEvaluator myEvaluator = new MatchEvaluator(RefToValue);
            string new_expression = regex.Replace(expression, myEvaluator);
            return new_expression;
        }

        public string RefToValue(Match m)
        {
            if (dictionary.ContainsKey(m.Value))
                if (dictionary[m.Value] == "")
                    return "0";
                else
                    return dictionary[m.Value];
            return m.Value;
        }

        public void Clear()
        {
            foreach (List<Cell> list in grid)
                list.Clear();
            grid.Clear();

            dictionary.Clear();

            current_rows = 0;
            current_cols = 0;
        }

        public void AddRow(DataGridView table)
        {
            current_rows++;
            List<Cell> new_row = new List<Cell>();

            for (int i = 0; i < current_cols; i++)
            {
                string name = FullName(i, current_rows - 1);
                new_row.Add(new Cell(name, current_rows - 1, i));
                dictionary.Add(name, "");
            }
            grid.Add(new_row);

            RefreshReferences();

            foreach (List<Cell> list in grid)
                foreach (Cell cell in list)
                    if (cell.referencesFromThis != null)
                        foreach (Cell cell_in_ref in cell.referencesFromThis)
                            if (cell_in_ref.RowIndex == current_rows - 1)
                                if (!cell_in_ref.pointersToThis.Contains(cell))
                                    cell_in_ref.pointersToThis.Add(cell);

            for (int i = 0; i < current_cols; i++)
               ChangeCellWithAllPointers(current_rows - 1, i, "", table);

        }

        public void AddCol(DataGridView table)
        {
            current_cols++;
            for (int i = 0; i < current_rows; i++)
            {
                string name = FullName(current_cols - 1, i);
                grid[i].Add(new Cell(name, i, current_cols - 1));
                dictionary.Add(name, "");
            }

            RefreshReferences();

            foreach (List<Cell> list in grid)
                foreach (Cell cell in list)
                    if (cell.referencesFromThis != null)
                        foreach (Cell cell_in_ref in cell.referencesFromThis)
                            if (cell_in_ref.ColIndex == current_cols - 1)
                                if (!cell_in_ref.pointersToThis.Contains(cell))
                                    cell_in_ref.pointersToThis.Add(cell);

            for (int i = 0; i < current_rows; i++)
                ChangeCellWithAllPointers(i, current_cols - 1, "", table);
        }

        public bool DeleteRow(DataGridView table)
        {
            // клетки которые ссылаются на клетки из последнего ряда
            List<Cell> sufferedCells = new List<Cell>();
            // не пустые клетки
            List<string> notEmptyCells = new List<string>();
            if (current_rows == 0)
                return false;

            int curRow = current_rows - 1;
            if (current_rows == 1)
            {
                DialogResult result = MessageBox.Show("Неможливо видалити цей рядок, оскільки він єдиний!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            for (int i = 0; i < current_cols; i++)
            {
                string name = FullName(i, curRow);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(name);
                if (grid[curRow][i].pointersToThis.Count() != 0)
                    sufferedCells.AddRange(grid[curRow][i].pointersToThis);
            }

            if ((sufferedCells.Count() != 0) || (notEmptyCells.Count() != 0))
            {
                string errorMessage = "";
                if (notEmptyCells.Count() != 0)
                {
                    errorMessage = "Ви збираєтесь видалили рядок, що містить заповнені клітини: ";
                    errorMessage += string.Join(", ", notEmptyCells.ToArray());
                    errorMessage += "." + Environment.NewLine + Environment.NewLine;
                }

                if (sufferedCells.Count() != 0)
                {
                    errorMessage += "Цей рядок містить клітини, на які є посилання в таблиці: ";
                    foreach (Cell cell in sufferedCells)
                        errorMessage += string.Join(", ", cell.Index);
                    errorMessage += "." + Environment.NewLine + Environment.NewLine;
                }

                errorMessage += "Продовжити?";
                DialogResult result = MessageBox.Show(errorMessage, "Попередження", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == System.Windows.Forms.DialogResult.No)
                    return false;
            }

            // удаляем клетки из словаря
            for (int i = 0; i < current_cols; i++)
            {
                string name = FullName(i, curRow);
                dictionary.Remove(name);
            }

            // обновляем поврежденные клетки
            foreach (Cell cell in sufferedCells)
                RefreshCellAndPointers(cell, table);

            // удаляем последнюю строку таблицы
            grid.RemoveAt(curRow);
            current_rows--;
            return true;
        }

        public bool DeleteCol(DataGridView dataGridView)
        {
            //клетки которые ссылаются на клетки из последнего столбца
            List<Cell> sufferedCells = new List<Cell>();
            //не пустые клетки
            List<string> notEmptyCells = new List<string>();
            //if (current_cols == 0)
            //    return false;
            if (current_cols == 1)
            {
                DialogResult result = MessageBox.Show("Неможливо видалити цей стовпчик, оскільки він єдиний!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            int curCol = current_cols - 1;

            for (int i = 0; i < current_rows; i++)
            {
                string name = FullName(current_cols - 1, i);
                if (dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] != " ")
                    notEmptyCells.Add(name);
                if (grid[i][curCol].pointersToThis.Count() != 0)
                    sufferedCells.AddRange(grid[i][curCol].pointersToThis);
            }

            if ((sufferedCells.Count() != 0) || (notEmptyCells.Count() != 0))
            {
                string message_error = "";
                if (notEmptyCells.Count() != 0)
                {
                    message_error = "Ви збираєтесь видалили стовпчик, що містить заповнені клітини: ";
                    message_error += string.Join(", ", notEmptyCells.ToArray());
                    message_error += "." + Environment.NewLine + Environment.NewLine;
                }
                if (sufferedCells.Count() != 0)
                {
                    message_error += "Цей стовпчик містить клітини, на які є посилання в таблиці: ";
                    foreach (Cell cell in sufferedCells)
                        message_error += string.Join(", ", cell.Index);
                    message_error += "." + Environment.NewLine + Environment.NewLine;
                }
                message_error += "Продовжити?";
                DialogResult result = MessageBox.Show(message_error, "Попередження", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == System.Windows.Forms.DialogResult.No)
                    return false;
            }

            //удаляем клетки из словаря
            for (int i = 0; i < current_rows; i++)
            { 
                string name = ltr_sys.ToLTR(curCol) + i.ToString();
                dictionary.Remove(name);
            }

            //обновляем поврежденные клетки
            foreach (Cell cell in sufferedCells)
                RefreshCellAndPointers(cell, dataGridView);

            //удаляем последний столбец таблицы
            for (int i = 0; i < current_rows; i++)
                grid[i].RemoveAt(curCol);

            current_cols--;
            return true;
        }

        public void Save(StreamWriter sw)
        {
            sw.WriteLine(current_rows);
            sw.WriteLine(current_cols);
            foreach (List<Cell> list in grid)
                foreach (Cell cell in list)
                {
                    sw.WriteLine(cell.Index);
                    sw.WriteLine(cell.Expression);
                    sw.WriteLine(cell.Value);

                    if (cell.referencesFromThis == null)
                        sw.WriteLine(0);
                    else
                    {
                        sw.WriteLine(cell.referencesFromThis.Count);
                        foreach (Cell refCell in cell.referencesFromThis)
                            sw.WriteLine(refCell.Index);
                    }

                    if (cell.pointersToThis == null)
                        sw.WriteLine(0);
                    else
                    {
                        sw.WriteLine(cell.pointersToThis.Count);
                        foreach (Cell ptrCell in cell.pointersToThis)
                            sw.WriteLine(ptrCell.Index);
                    }
                }
        }

        public void Open(int row, int col, StreamReader sr, DataGridView dataGridView)
        {
            for (int r = 0; r < row; r++)
            {
                for (int c = 0; c < col; c++)
                {

                    string index = sr.ReadLine();
                    string expression = sr.ReadLine();
                    string value = sr.ReadLine();

                    if (expression != "")
                        dictionary[index] = value;
                    else
                        dictionary[index] = "";

                    int refCount = Convert.ToInt32(sr.ReadLine());
                    List<Cell> newRef = new List<Cell>();
                    string refer;
                    for (int i = 0; i < refCount; i++)
                    {
                        refer = sr.ReadLine();
                        newRef.Add(grid[ltr_sys.FromIndex(refer)[1]][ltr_sys.FromIndex(refer)[0]]);
                    }


                    int ptrCount = Convert.ToInt32(sr.ReadLine());
                    List<Cell> newPtr = new List<Cell>();
                    string point;
                    for (int i = 0; i < ptrCount; i++)
                    {
                        point = sr.ReadLine();
                        newPtr.Add(grid[ltr_sys.FromIndex(point)[1]][ltr_sys.FromIndex(point)[0]]);
                    }


                    grid[r][c].SetCell(value, expression, newRef, newPtr);

                    int icol = grid[r][c].ColIndex;
                    int irow = grid[r][c].RowIndex;
                    dataGridView[icol, irow].Value = dictionary[index];
                }
            }

        }
    }
}