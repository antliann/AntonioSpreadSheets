using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Antonio_Spreadsheets
{
    public partial class Form1 : Form
    {
        public Grid GR = new Grid();
        private static Alphabet ltr_sys = new Alphabet();

        public Form1()
        {
            InitializeComponent();
            InitTable(GR.current_rows, GR.current_cols);
        }
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.відкритиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.зберегтиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.інфоToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripLabel6 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripLabel3 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripLabel7 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripLabel4 = new System.Windows.Forms.ToolStripLabel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ExprTextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.інфоToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(5, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1325, 28);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.відкритиToolStripMenuItem,
            this.зберегтиToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.ShortcutKeyDisplayString = "";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(59, 24);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // відкритиToolStripMenuItem
            // 
            this.відкритиToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("відкритиToolStripMenuItem.Image")));
            this.відкритиToolStripMenuItem.Name = "відкритиToolStripMenuItem";
            this.відкритиToolStripMenuItem.ShortcutKeyDisplayString = "Ctrl+O";
            this.відкритиToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.відкритиToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.відкритиToolStripMenuItem.Text = "Відкрити";
            this.відкритиToolStripMenuItem.Click += new System.EventHandler(this.відкритиToolStripMenuItem_Click);
            // 
            // зберегтиToolStripMenuItem
            // 
            this.зберегтиToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("зберегтиToolStripMenuItem.Image")));
            this.зберегтиToolStripMenuItem.Name = "зберегтиToolStripMenuItem";
            this.зберегтиToolStripMenuItem.ShortcutKeyDisplayString = "Ctrl+S";
            this.зберегтиToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.зберегтиToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.зберегтиToolStripMenuItem.Text = "Зберегти";
            this.зберегтиToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal;
            this.зберегтиToolStripMenuItem.Click += new System.EventHandler(this.зберегтиToolStripMenuItem_Click);
            // 
            // інфоToolStripMenuItem
            // 
            this.інфоToolStripMenuItem.Name = "інфоToolStripMenuItem";
            this.інфоToolStripMenuItem.Size = new System.Drawing.Size(55, 24);
            this.інфоToolStripMenuItem.Text = "Інфо";
            this.інфоToolStripMenuItem.Click += new System.EventHandler(this.інфоToolStripMenuItem_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "sprt";
            this.openFileDialog1.FileName = "Відкрити файл";
            this.openFileDialog1.Filter = "Spreadsheets файли (*sprt)|*.sprt";
            this.openFileDialog1.FilterIndex = 0;
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "sprt";
            this.saveFileDialog1.Filter = "Spreadsheets файли (*sprt)|*.sprt";
            this.saveFileDialog1.Title = "Таблиця";
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel2,
            this.toolStripLabel6,
            this.toolStripLabel3,
            this.toolStripLabel7,
            this.toolStripLabel1,
            this.toolStripLabel4});
            this.toolStrip1.Location = new System.Drawing.Point(0, 28);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 1, 60);
            this.toolStrip1.Size = new System.Drawing.Size(1325, 113);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripLabel2.Image")));
            this.toolStripLabel2.Margin = new System.Windows.Forms.Padding(15, 1, 15, 2);
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Padding = new System.Windows.Forms.Padding(0, 30, 0, 0);
            this.toolStripLabel2.Size = new System.Drawing.Size(124, 50);
            this.toolStripLabel2.Text = "Додати рядок";
            this.toolStripLabel2.Click += new System.EventHandler(this.toolStripLabel2_Click);
            // 
            // toolStripLabel6
            // 
            this.toolStripLabel6.Image = ((System.Drawing.Image)(resources.GetObject("toolStripLabel6.Image")));
            this.toolStripLabel6.Margin = new System.Windows.Forms.Padding(0, 1, 15, 2);
            this.toolStripLabel6.Name = "toolStripLabel6";
            this.toolStripLabel6.Size = new System.Drawing.Size(140, 50);
            this.toolStripLabel6.Text = "Видалити рядок";
            this.toolStripLabel6.Click += new System.EventHandler(this.toolStripLabel6_Click);
            // 
            // toolStripLabel3
            // 
            this.toolStripLabel3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripLabel3.Image")));
            this.toolStripLabel3.Margin = new System.Windows.Forms.Padding(15, 1, 15, 2);
            this.toolStripLabel3.Name = "toolStripLabel3";
            this.toolStripLabel3.Size = new System.Drawing.Size(146, 50);
            this.toolStripLabel3.Text = "Додати стовпчик";
            this.toolStripLabel3.Click += new System.EventHandler(this.toolStripLabel3_Click);
            // 
            // toolStripLabel7
            // 
            this.toolStripLabel7.Image = ((System.Drawing.Image)(resources.GetObject("toolStripLabel7.Image")));
            this.toolStripLabel7.Margin = new System.Windows.Forms.Padding(0, 1, 15, 2);
            this.toolStripLabel7.Name = "toolStripLabel7";
            this.toolStripLabel7.Size = new System.Drawing.Size(162, 50);
            this.toolStripLabel7.Text = "Видалити стовпчик";
            this.toolStripLabel7.Click += new System.EventHandler(this.toolStripLabel7_Click);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(0, 50);
            // 
            // toolStripLabel4
            // 
            this.toolStripLabel4.Image = ((System.Drawing.Image)(resources.GetObject("toolStripLabel4.Image")));
            this.toolStripLabel4.Margin = new System.Windows.Forms.Padding(30, 1, 0, 2);
            this.toolStripLabel4.Name = "toolStripLabel4";
            this.toolStripLabel4.Size = new System.Drawing.Size(138, 50);
            this.toolStripLabel4.Text = "Список функцій";
            this.toolStripLabel4.Click += new System.EventHandler(this.toolStripLabel4_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridView1.Location = new System.Drawing.Point(0, 141);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1325, 453);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // ExprTextBox
            // 
            this.ExprTextBox.Location = new System.Drawing.Point(32, 101);
            this.ExprTextBox.Name = "ExprTextBox";
            this.ExprTextBox.Size = new System.Drawing.Size(415, 22);
            this.ExprTextBox.TabIndex = 1;
            this.ExprTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ExprTextBox_KeyDown);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(468, 94);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(114, 38);
            this.button1.TabIndex = 4;
            this.button1.Text = "Підтвердити";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1325, 594);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.ExprTextBox);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Antonio SpreadSheets";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void InitTable(int row, int col)
        {
            for (int i = 0; i < col; i++)
            {
                string colname = ltr_sys.ToLTR(i);
                dataGridView1.Columns.Add(colname, colname);
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            dataGridView1.RowCount = row;
            for (int i = 0; i < row; ++i)
            {
                dataGridView1.Rows[i].HeaderCell.Value = i.ToString();
            }
            GR.SetGrid(row, col);
        }

        private void button1_Click(object sender, EventArgs e)                       // ПОДТВЕРЖДЕНИЕ ВВОДА
        {
            int col = dataGridView1.SelectedCells[0].ColumnIndex;
            int row = dataGridView1.SelectedCells[0].RowIndex;
            string expr = ExprTextBox.Text;

            GR.ChangeCellWithAllPointers(row, col, expr, dataGridView1);
            dataGridView1[col, row].Value = GR.grid[row][col].Value;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)    // ПОКАЗЫВАЕТ ВЫРАЖЕНИЕ
        {
            int col = dataGridView1.SelectedCells[0].ColumnIndex;
            int row = dataGridView1.SelectedCells[0].RowIndex;

            string expr = GR.grid[row][col].Expression;
            string value = GR.grid[row][col].Value;

            ExprTextBox.Text = expr;
            ExprTextBox.Focus();
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)              //  ДОБАВИТЬ СТРОКУ
        {
            DataGridViewRow row = new DataGridViewRow();

            dataGridView1.Rows.Add(row);

            RefreshRowNumbers();
            GR.AddRow(dataGridView1);
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)                // ДОБАВИТЬ СТОЛБИК
        {
            string colname = ltr_sys.ToLTR(GR.current_cols);
            dataGridView1.Columns.Add(colname, colname);

            GR.AddCol(dataGridView1);
            RefreshColNumbers();
        }

        private void ExprTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            PressEnter(sender, e);
        }

        private void toolStripLabel6_Click(object sender, EventArgs e)              // УДАЛИТЬ СТРОКУ
        {
            int curRow = GR.current_rows - 1;
            if (!GR.DeleteRow(dataGridView1))
                return;
            dataGridView1.Rows.RemoveAt(curRow);
            RefreshRowNumbers();
        }

        private void toolStripLabel7_Click(object sender, EventArgs e)              // УДАЛИТЬ СТОЛБИК
        {
            int curCol = GR.current_cols - 1;
            if (!GR.DeleteCol(dataGridView1))
                return;

            dataGridView1.Columns.RemoveAt(curCol);
            RefreshColNumbers();
        }

        private void RefreshRowNumbers()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
                row.HeaderCell.Value = String.Format("{0}", row.Index);
        }

        private void RefreshColNumbers()
        {
            int col = dataGridView1.ColumnCount;
            for (int i = 0; i < col; i++)
            {
                string colname = ltr_sys.ToLTR(i);
                dataGridView1.Columns[i].HeaderCell.Value = colname;
            }
        }

        private void зберегтиToolStripMenuItem_Click(object sender, EventArgs e)               // ЗБЕРЕГТИ ФАЙЛ
        {
            SaveDoc();
        }

        private void відкритиToolStripMenuItem_Click(object sender, EventArgs e)                // ВІДКРИТИ ФАЙЛ
        {
            OpenDoc();
        }

        private void інфоToolStripMenuItem_Click(object sender, EventArgs e)               // КНОПКА ОПЕРАЦІЙ
        {
            ShowInfo();
        }

        private void toolStripLabel4_Click(object sender, EventArgs e)                    // КНОПКА ФУНКЦІЙ
        {
            ShowFunctions();
        }


                  // Действия при нажатиях:

        private void SaveDoc()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Antonio SpreadSheets (*sprt)|*.sprt";
            saveFileDialog.Title = "Зберегти файл";
            saveFileDialog.FileName = "Таблиця Antonio SpreadSheets";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                FileStream fs = (FileStream)saveFileDialog.OpenFile();

                StreamWriter sw = new StreamWriter(fs);

                GR.Save(sw);

                sw.Close();
                fs.Close();
            }
        }

        private void OpenDoc()
        {
            DialogResult res = MessageBox.Show("Поточну таблицю буде знищено після відкриття нового файлу! Бажаєте її зберегти?", "Увага", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (res == System.Windows.Forms.DialogResult.Yes)
                SaveDoc();
            if (res == System.Windows.Forms.DialogResult.Cancel)
                return;
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Antonio SpreadSheets (*sprt)|*.sprt";
                openFileDialog.Title = "Відктрити файл";
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                    return;
                StreamReader sr = new StreamReader(openFileDialog.FileName);

                GR.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                Int32.TryParse(sr.ReadLine(), out int row);
                Int32.TryParse(sr.ReadLine(), out int col);

                InitTable(row, col);

                GR.Open(row, col, sr, dataGridView1);

                sr.Close();
            }
            catch
            {
                string text = "Неможливо відкрити файл, оскільки він є пошкодженим!";
                DialogResult result = MessageBox.Show(text, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PressEnter(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int col = dataGridView1.SelectedCells[0].ColumnIndex;
                int row = dataGridView1.SelectedCells[0].RowIndex;
                string expr = ExprTextBox.Text;

                GR.ChangeCellWithAllPointers(row, col, expr, dataGridView1);
                dataGridView1[col, row].Value = GR.grid[row][col].Value;
            }
        }

        private void ShowFunctions()
        {
            string spaces = "                                                                                         ";
            string func = "= M20            Посилання на клітину M20"
                        + Environment.NewLine + Environment.NewLine
                        + "= A + B          Сума А та В"
                        + Environment.NewLine + Environment.NewLine
                        + "= А - В           Різниця А та В"
                        + Environment.NewLine + Environment.NewLine
                        + "= A * B           Добуток А та В"
                        + Environment.NewLine + Environment.NewLine
                        + "= А / В           Частка А та В"
                        + Environment.NewLine + Environment.NewLine
                        + "= А ^ В          А в степені В"
                        + Environment.NewLine + Environment.NewLine
                        + "= inc(A)          Збільшення А на 1"
                        + Environment.NewLine + Environment.NewLine
                        + "= dec(B)         Зменшення B на 1"
                        + Environment.NewLine + Environment.NewLine
                        + "= max(А, В)    Найбільше з А та В"
                        + Environment.NewLine + Environment.NewLine
                        + "= min(А, В)    Найменше з А та В"
                        + Environment.NewLine + Environment.NewLine
                        + "= + А             Унарний плюс (Не змінює число)"
                        + Environment.NewLine + Environment.NewLine
                        + "= - B              Унарний мінус (Змінює знак числа на" + spaces + " протилежний)";

            DialogResult result = MessageBox.Show(func, "Список функцій", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowInfo()
        {
            string info = "Програма Antonio SpreadSheets призначена для роботи з таблицями даних."

                + Environment.NewLine + Environment.NewLine

                + "Функції програми дозволяють додавати та видаляти рядки і стовпчики за допомогою відповідних кнопок."

                + Environment.NewLine + Environment.NewLine

                + "Щоб встановити значення для певної клітини, виділіть клітину, введіть вираз у текствове поле та "
                + "натисніть кнопку «Підтвердити» або клавішу Enter. "
                + "У клітині ви побачите значення введеного виразу, а сам вираз ви побачите "
                + "у текстовому полі при натисканні на клітину."

                + Environment.NewLine + Environment.NewLine

                + "Формат запису виразів ви можете побачити у розділі «Список функцій»."

                + Environment.NewLine + Environment.NewLine

                + "Збереження та відкриття файлів доступно у розділі «Файл»."

                + Environment.NewLine + Environment.NewLine

                + "© Лянной Антон, 2019." + Environment.NewLine
                + "Все права защищены (но это не точно)";
            DialogResult result = MessageBox.Show(info, "Інфо", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult res = MessageBox.Show("Після закриття програми поточну таблицю буде знищено! Бажаєте її зберегти?", "Увага", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (res == System.Windows.Forms.DialogResult.Yes)
                SaveDoc();
            if (res == System.Windows.Forms.DialogResult.Cancel)
                e.Cancel = true;
        }
    }
}

/*foreach (Cell cell in grid[row][col].pointersToThis)
            {
                RefreshCellAndPointers(cell, table);
            }*/
