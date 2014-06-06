/**************************************************************************\
Module Name:   Entry.cs 
Project:       excelParse
Author:        Steven Rau

This file contains the Entry class. An entry is contained in a single row
of info in the .csv input file.
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelParse
{
    /**
     * This class is used to store table entries.
     * There are fields for each of the 10 fields we want to parse
     */
    class Entry : IComparable<Entry>
    {
        //Fields fot the 10 fields we want to store for an entry
        public string colA;
        public string colB;
        public string colC;
        public string colD;
        public string colE;
        public string colF;
        public string colG;
        public int colH;
        public int colI;
        public int colM;

        public Entry(string a, string b, string c, string d, string e, string f, string g, int h, int i, int m)
        {
            colA = a;
            colB = b;
            colC = c;
            colD = d;
            colE = e;
            colF = f;
            colG = g;
            colH = h;
            colI = i;
            colM = m;
        }

        /**
         * Create a custom toString method
         */
        public override string ToString()
        {
            string output = colA + " " + colB + " " + colC + " " + colD + " " + colE + " " +
                            colF + " " + colG + " " + colH + " " + colI + " " + colM;
            return output;
        }

        /*
         * Custom comparison method that simply comares by rat ID for now
         */
        public int CompareTo(Entry compareEntry)
        {
            return this.colH.CompareTo(compareEntry.colH);
        }
    }
}
