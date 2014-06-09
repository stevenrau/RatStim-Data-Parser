/**************************************************************************\
Module Name:   ParseAndPrint.cs 
Project:       excelParse
Author:        Steven Rau

This file is used to parse in .csv files and output their contents to
either .txt or .xlsx files
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace excelParse
{
    /*
     * This class is used to read in CSV or Excel files and parse them
     */
    public class ParseAndPrint
    {
        private string outPath;
        private List<string> inPaths;
        private int inPathCount;
        private List<Entry> entries;
        Dictionary<int, RatById> ratsById;
        List<int> ratIds;

        /**
         * Constructor for the parseAndPrint class.
         * 
         * @param input  List of string representations of the input paths
         *        output String representation of the output path
         */
        public ParseAndPrint(List<string> input, int numInPaths, string output)
        {
            inPaths = new List<string>(input);
            inPathCount = numInPaths;
            outPath = output;
            entries = new List<Entry>();
            ratsById = new Dictionary<int, RatById>();
            ratIds = new List<int>();

            getCsvEntries();
            getRatsById();
        }

        /**
         * Prints the important values from the .csv input file to a specified text file
         * Mainly used for testing and debugging. Shouldn't be used in the final product.
         * 
         * @param output  String representation of the ouotput file. Needs to be a text file.
         */
        public void printCsvToText(string output)
        {
            //Open the output file
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(output);

            //Start reading from the input file
            try
            {
                var reader = new StreamReader(File.OpenRead(inPaths.First()));
                while (!reader.EndOfStream)
                {
                    //Read in an entire line
                    var line = reader.ReadLine();
                    //Then split the values separated by a comma
                    var values = line.Split(',');

                    outFile.Write(values[0]);
                    outFile.Write(" ");
                    outFile.Write(values[1]);
                    outFile.Write(" ");
                    outFile.Write(values[2]);
                    outFile.Write(" ");
                    outFile.Write(values[3]);
                    outFile.Write(" ");
                    outFile.Write(values[4]);
                    outFile.Write(" ");
                    outFile.Write(values[5]);
                    outFile.Write(" ");
                    outFile.Write(values[6]);
                    outFile.Write(" ");
                    outFile.Write(values[7]);
                    outFile.Write(" ");
                    outFile.Write(values[8]);
                    outFile.Write(" ");
                    outFile.Write(values[12]);
                    outFile.Write("\n");
                }
            }
            catch (IOException )
            {
                MessageBox.Show(inPaths.First()+" is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                outFile.Close();
            }

        }

        /** Takes the csv entries and output them to an excel file at the
         * path outPath, one entry per row
         */
        public void printToExcelUnsorted()
        {
            // Create the file using the FileInfo object
            var file = new FileInfo(outPath);
            if (file.Exists)
            {
                file.Delete();  // ensures we create a new workbook
                file = new FileInfo(outPath);
            }

            //Create the Excel package and make a new workbook
            ExcelPackage pck = new ExcelPackage(file);
            ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("Master");

            int i = 1;
            foreach (Entry curEntry in entries)
            {
                worksheet.Cells[i, 1].Value = curEntry.colA;
                worksheet.Cells[i, 2].Value = curEntry.colB;
                worksheet.Cells[i, 3].Value = curEntry.colC;
                worksheet.Cells[i, 4].Value = curEntry.colD;
                worksheet.Cells[i, 5].Value = curEntry.colE;
                worksheet.Cells[i, 6].Value = curEntry.colF;
                worksheet.Cells[i, 7].Value = curEntry.colG;
                worksheet.Cells[i, 8].Value = curEntry.colH;
                worksheet.Cells[i, 9].Value = curEntry.colI;
                worksheet.Cells[i, 10].Value = curEntry.colM;

                i++;
            }

            //Resive the columns so that they fit nicely
            for (i = 1; i <= worksheet.Dimension.End.Column; i++) 
            { 
                worksheet.Column(i).AutoFit(); 
            }

            pck.Save();
        }

        public void printToExcelSorted()
        {
            // Create the file using the FileInfo object
            var file = new FileInfo(outPath);
            if (file.Exists)
            {
                file.Delete();  // ensures we create a new workbook
                file = new FileInfo(outPath);
            }

            //Create the Excel package and make a new workbook
            ExcelPackage pck = new ExcelPackage(file);
            ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("Master");

            int i = 1;
            foreach(int ratId in ratIds)
            {
                foreach(string curStim in Constants.stims)
                {
                    foreach(Entry curEntry in ratsById[ratId].entries)
                    {
                        string entryStim = System.Text.RegularExpressions.Regex.Replace(curEntry.colG, @"\s+", "");
                        if (curStim.CompareTo(entryStim) == 0)
                        {
                            worksheet.Cells[i, 1].Value = curEntry.colA;
                            worksheet.Cells[i, 2].Value = curEntry.colB;
                            worksheet.Cells[i, 3].Value = curEntry.colC;
                            worksheet.Cells[i, 4].Value = curEntry.colD;
                            worksheet.Cells[i, 5].Value = curEntry.colE;
                            worksheet.Cells[i, 6].Value = curEntry.colF;
                            worksheet.Cells[i, 7].Value = curEntry.colG;
                            worksheet.Cells[i, 8].Value = curEntry.colH;
                            worksheet.Cells[i, 9].Value = curEntry.colI;
                            worksheet.Cells[i, 10].Value = curEntry.colM;

                            i++;
                        }
                    }
                }
                i++; //Advance one more row (leave a blank row between rat #s)
            }

            //Resive the columns so that they fit nicely
            for (i = 1; i <= worksheet.Dimension.End.Column; i++) 
            { 
                worksheet.Column(i).AutoFit(); 
            }

            pck.Save(); //And save

            System.Diagnostics.Process.Start(outPath);
        }

        /**
         * Prints the csv entries from the entries list to a txt file
         * */
        public void printEntriesToText(string output)
        {
            //Open the output file
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(output);

            //Start reading from the input file
            try
            {
                foreach (Entry curEntry in entries)
                {
                    outFile.Write(curEntry);
                    outFile.Write("\n");
                }
            }
            catch (IOException)
            {
                MessageBox.Show(inPaths.First() + " is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                outFile.Close();
            }

        }

        /**
         * Stores entries with the same ID in a list within a RatById object which itself is
         * stored in a dictionary along with other RatById objects for all the other IDs
         */
        private void getRatsById()
        {
            //Go throught the entry list and count the different rat IDs (store them in a list)
            foreach (Entry curEntry in entries)
            {
                //If it's a new ID we havent seen yet...
                if (!ratIds.Contains(curEntry.colH))
                {
                    //Add it to the id list
                    ratIds.Add(curEntry.colH);
                    //And make a new RatById entry to keep track of that rat's entries
                    RatById newrat = new RatById(curEntry.colH);
                    newrat.entries.Add(curEntry);
                    ratsById.Add(newrat.id, newrat);
                }
                //Else just add the entry to the ratById entry with the corresponding ID 
                else
                {
                    ratsById[curEntry.colH].entries.Add(curEntry);
                }
            }
        }

        /**
         * Method to get the entries from the csv input file and store them in 
         * the entries list. Should only need to be called once in the constructer
         */
        private void getCsvEntries()
        {
            //Start reading from the input file
            try
            {
                var reader = new StreamReader(File.OpenRead(inPaths.First()));
                while (!reader.EndOfStream)
                {
                    //Read in an entire line
                    string line = reader.ReadLine();
                    line.Replace("  ", "");
                    //Then split the values separated by a comma
                    var values = line.Split(',');

                    //Make a new Entry object for this entry, then add it to the list
                    Entry newEntry = new Entry(values[0], values[1], values[2], values[3], values[4], values[5], values[6],  
                                               Convert.ToInt32(values[7]), Convert.ToInt32(values[8]), Convert.ToInt32(values[12]));
                    entries.Add(newEntry);
                }
            }
            catch (IOException)
            {
                MessageBox.Show(inPaths.First() + " is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            entries.Sort();
        }
    }
}
