using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;

namespace Antonio_Spreadsheets
{
    public class Cell
    {
        public string Index { get; private set; }
        public int ColIndex { get; private set; }
        public int RowIndex { get; private set; }
        public string Value;
        public string Expression;

        public List<Cell> pointersToThis = new List<Cell>();
        public List<Cell> referencesFromThis = new List<Cell>();
        public List<Cell> new_referencesFromThis = new List<Cell>();

        public Cell(string index, int row, int col)
        {
            Index = index;
            RowIndex = row;
            ColIndex = col;
            Value = "0";
            Expression = "";
        }

        public void SetCell(string value, string expression, List<Cell> references, List<Cell> pointers)
        {
            this.Value = value;
            this.Expression = expression;

            this.referencesFromThis.Clear();
            this.referencesFromThis.AddRange(references);

            this.pointersToThis.Clear();
            this.pointersToThis.AddRange(pointers);
        }
        public bool CheckForCycle(List<Cell> check_list)
        {
            foreach (Cell susp in check_list)
            {
                if (susp.Index == Index) return false;
            }
            foreach (Cell point in pointersToThis)
            {
                foreach (Cell susp in check_list)
                {
                    if (susp.Index == point.Index) return false;
                }
                if (!point.CheckForCycle(check_list)) return false;
            }
            return true;
        }
        public void AddPointersAndReferences()
        {
            foreach (Cell point in new_referencesFromThis)
            {
                point.pointersToThis.Add(this);
            }
            referencesFromThis = new_referencesFromThis;
        }
        public void DelPointersAndReferences()
        {
            if (referencesFromThis != null)
            foreach (Cell point in referencesFromThis)
            {
                point.pointersToThis.Remove(this);
            }
            referencesFromThis = null;
        }
    }
}