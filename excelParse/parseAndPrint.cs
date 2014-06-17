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

namespace RatStim
{
    /*
     * This class is used to read in CSV or Excel files and parse them
     */
    public class ParseAndPrint
    {
        private string outPath;               //String representation of the output path
        private List<string> inPaths;         //List of all the input csv file paths
        private int inPathCount;              //The number of input files
        private List<Entry> entries;          //A list to keep track of all the entries read in from the input csv files
        Dictionary<string, RatById> ratsById; //Each unique rat ID gets an entry in this dictionary with all of its entries
        List<string> ratIds;                  //A list of all he unique rat IDs
        List<double> avgs;                    //A list of the averages for each stimulus group for each rat ID
        List<string> ratStims;                //A list of all the stimulus values for the rats being entered

        /**
         * Constructor for the parseAndPrint class.
         * 
         * @param input     List of string representations of the input paths
         *        numPaths  The number of input files provided
         *        output    String representation of the output path
         */
        public ParseAndPrint(List<string> input, int numInPaths, string output)
        {
            inPaths = new List<string>(input);
            inPathCount = numInPaths;
            outPath = output;
            entries = new List<Entry>();
            ratsById = new Dictionary<string, RatById>();
            ratIds = new List<string>();
            avgs = new List<double>();
            ratStims = new List<string>();

            getCsvEntries();
            getRatsById();
        }

        /**
         * Prints the important values from the .csv input file to a specified text file
         * Mainly used for testing and debugging. Shouldn't be used in the final product.
         * 
         * @param output  String representation of the output file. Needs to be a text file.
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

        /**
         * Prints the entries to an excel file sorted by rat ID, with avgs calculated
         * for each stimulus
         */
        public void printIntermediateData()
        {
            // Create the file using the FileInfo object
            string fileName = Path.GetFileName(outPath);
            string extraDirectoryPath = Path.GetDirectoryName(outPath) + "/" + Path.GetFileNameWithoutExtension(outPath) + "_INTERMEDIATE_DATA";
            Directory.CreateDirectory(extraDirectoryPath);
            extraDirectoryPath += "/intermediate_" + fileName;

            var file = new FileInfo(extraDirectoryPath);
            if (file.Exists)
            {
                file.Delete();  // ensures we create a new workbook
                file = new FileInfo(outPath);
            }

            //Create the Excel package and make a new workbook
            ExcelPackage pck = new ExcelPackage(file);
            ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("Master");

            int i = 1;
            foreach(string ratId in ratIds)
            {
                foreach(string curStim in ratStims)
                {
                    //The variables used to calculates averages for each stimulus
                    double curSum = 0;
                    int entryCnt = 0;
                    double curAvg = 0;
                    int p120Cnt = 0;

                    foreach(Entry curEntry in ratsById[ratId].entries)
                    {
                        //string entryStim = System.Text.RegularExpressions.Regex.Replace(curEntry.colG, @"\s+", "");
                        string entryStim = curEntry.colG;
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
                            
                            //Keep track of the avg of this set of entries with common stimulus
                            curSum += curEntry.colM;
                            entryCnt++;

                            if (curStim.CompareTo("p120") == 0)
                            {
                                //If we are on our 6th p120 in a row
                                if (p120Cnt == 5)
                                {
                                    //Then print the avg 
                                    curAvg = curSum / entryCnt;
                                    worksheet.Cells[i - 1, 11].Value = curAvg;
                                    //And add it to the avg list to be used in std deviation calculation
                                    avgs.Add(curAvg);

                                    //Reset the counting vals
                                    curSum = 0;
                                    entryCnt = 0;
                                    p120Cnt = 0;
                                }
                                else
                                {
                                    p120Cnt++;
                                }
                            }

                            i++;
                        }
                    }

                    //Since the p120's have already been printed, don't re-print them
                    if (curStim.CompareTo("p120") != 0)
                    {
                        //Print the avg 
                        curAvg = curSum / entryCnt;
                        worksheet.Cells[i - 1, 11].Value = curAvg;
                        //And add it to the avg list to be used in std deviation calculation
                        avgs.Add(curAvg);
                    }
                }
                i++; //Advance one more row (leave a blank row between rat #s)
            }

            resizeCols(worksheet);

            pck.Save(); //And save

        }

        /**
         * Prints the pre-calculated rat data (avgs for all of the stim values for each rat) 
         * to a well-formatted Master excel sheet
         */
        public void printMasterData()
        {
            var file = new FileInfo(outPath);
            if (file.Exists)
            {
                file.Delete();  // ensures we create a new workbook
                file = new FileInfo(outPath);
            }

            //Create the Excel package and make a new workbook
            ExcelPackage pck = new ExcelPackage(file);
            ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("Master");

            pck.Save();
        }

        /**
         * Resizes the columns of an excel worksheet so that they are sized appropriately
         * to their contents
         * 
         * @param  worksheet  The Excel worksheet we want to resize columns on
         */
        public void resizeCols(ExcelWorksheet worksheet)
        {
            //Resive the columns so that they fit nicely
            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                worksheet.Column(i).AutoFit();
            }
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
            //colD is the rat ID column, in string form. ex: 'A2L1f1'
            foreach (Entry curEntry in entries)
            {
                //If it's a new ID we havent seen yet...
                if (!ratIds.Contains(curEntry.colD))
                {
                    //Add it to the id list
                    ratIds.Add(curEntry.colD);
                    //And make a new RatById entry to keep track of that rat's entries
                    RatById newrat = new RatById(curEntry.colD);
                    newrat.entries.Add(curEntry);
                    ratsById.Add(newrat.id, newrat);
                }
                //Else just add the entry to the ratById entry with the corresponding ID 
                else
                {
                    ratsById[curEntry.colD].entries.Add(curEntry);
                }

                //If it's a new stimulus we haven't seen yet, add it to the stim list
                if (!ratStims.Contains(curEntry.colG))
                {
                    ratStims.Add(curEntry.colG);
                }
            }

            //Sort each RatById's list of entries by trial # so that they are printed in order
            foreach (string curRatId in ratIds)
            {
               // ratsById[curRatId].entries.Sort();
            }

            //Then sorth the ratStim values so that they are alpahbetical
            ratStims.Sort();
        }

        /**
         * Method to get the entries from the csv input files and store them in 
         * the entries list. Should only need to be called once in the constructer
         */
        private void getCsvEntries()
        {
            //Read through each one of the input files
            foreach (string curInPath in inPaths)
            {
                try
                {
                    var reader = new StreamReader(File.OpenRead(curInPath));
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

        /**
         * Returns the string representation of the output file path
         */
        public string getOutPath()
        {
            return outPath;
        }
    }
}
