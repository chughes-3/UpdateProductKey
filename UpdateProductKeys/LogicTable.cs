using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UpdateProductKeys
{
    public class LogicTable
    {
        public static List<bool> Start()
        {
            return new List<bool>();
        }
    }
    public static class LogicTableExtensions
    {
        #region Condition

        public static List<bool> Condition(this List<bool> table, bool test, string tableValues)
        {
            if (table.Count == 0)
                table.AddRange(tableValues.ToCharArray().Select(i => true));

            if (table.Count != tableValues.Length)
                throw new Exception();

            bool result = test;

            List<bool> res = new List<bool>();

            for (int i = 0; i < tableValues.Length; i++)
            {
                switch (tableValues[i])
                {
                    // no modifier
                    case '-':
                        res.Add(table[i]);
                        break;

                    // if test is true
                    case 'T':
                        res.Add(table[i] && result);
                        break;
                    
                    // if test is false
                    case 'F':
                        res.Add(table[i] && !result);
                        break;

                    default:
                        res.Add(false);
                        break;
                }
            }

            return res;
        }

        public static List<bool> Condition(this List<bool> table, Func<bool> test, string tableValues)
        {
            return table.Condition(test(), tableValues);
        }

        public static List<bool> Condition(this List<bool> table, string tableValues, Func<bool> test)
        {
            return table.Condition(test(), tableValues);
        }

        public static List<bool> Condition(this List<bool> table, string tableValues, bool test)
        {
            return table.Condition(test, tableValues);
        }

        #endregion

        #region Action

        public static List<bool> Action(this List<bool> table, Action action, string tableValues)
        {
            if (table.Count != tableValues.Length)
                throw new Exception();

            bool doAction = false;

            for (int i = 0; i < tableValues.Length; i++)
            {
                if (tableValues[i] == 'X' && table[i])
                    doAction = true;
            }

            if (doAction)
                action();

            return table;
        }

        public static List<bool> Action(this List<bool> table, string tableValues, Action action)
        {
            return table.Action(action, tableValues);
        }

        #endregion
    }
}
