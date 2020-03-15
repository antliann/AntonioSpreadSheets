using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Antlr4.Runtime.Misc;
using Antlr4.Runtime;

namespace Antonio_Spreadsheets
{
    public class Alphabet
    {
        public string ToLTR(int i)
        {
            int k = 0;
            int[] ltr_num = new int[200];
            while (i > 25)
            {
                ltr_num[k] = i / 26 - 1;
                i = i % 26;
                k++;
            }
            ltr_num[k] = i;
            string letters = "";
            for (int j = 0; j <= k; j++)
            {
                letters += ((char)('A' + ltr_num[k])).ToString();
            }
            return letters;
        }
        public int[] FromIndex(string index)
        {
            int[] digit_num = new int[2];
            StringBuilder first_part = new StringBuilder();
            int letter_num = 0;
            foreach (char c in index)
            {
                if (Char.IsLetter(c))
                {
                    first_part.Append(c);
                    letter_num++;
                    continue;
                }
                string first = first_part.ToString();

                char[] char_index = first.ToCharArray();
                int len = char_index.Length;
                int num1 = 0;
                for (int i = len - 2; i >= 0; --i)
                {
                    num1 += (((int)char_index[i] - (int)'A') + 1) * Convert.ToInt32(Math.Pow(26, len - i - 1));
                }
                num1 += ((int)char_index[len - 1] - (int)'A');
                digit_num[0] = num1;
                break;
            }
            digit_num[1] = Convert.ToInt32(index.Substring(letter_num));
            return digit_num;
        }
    }
}