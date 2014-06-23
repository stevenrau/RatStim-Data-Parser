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
using System.Drawing;
using System.Windows;
using OfficeOpenXml.Style;

namespace RatStim
{
    /*
     * This class is used to read in CSV or Excel files and parse them
     */
    public class ParseAndPrint
    {
        private string outPath;                //String representation of the output path
        private string extraOutPath;           //String representation of the path where the intermediate files are stored
        private string outFileName;            //String representation of the file name without the directory/path info
        private List<string> inPaths;          //List of all the input csv file paths
        private int inPathCount;               //The number of input files
        private List<Entry> entries;           //A list to keep track of all the entries read in from the input csv files
        Dictionary<string, RatById> ratsById;  //Each unique rat ID gets an entry in this dictionary with all of its entries
        List<string> ratIds;                   //A list of all he unique rat IDs
        List<string> ratStims;                 //A list of all the stimulus values for the rats being entered

        /**
         * Constructor for the parseAndPrint class.
         * 
         * @param input     List of string representations of the input paths
         *        numPaths  The number of input files provided
         *        output    String representation of the output path
         */
        public ParseAndPrint(List<string> input, int numInPaths, string output)
        {
            //Set the nput path info
            inPaths = new List<string>(input);
            inPathCount = numInPaths;

            //Set the output file names and paths
            outPath = output;
            outFileName = Path.GetFileName(outPath);
            extraOutPath = Path.GetDirectoryName(outPath) + "/" + Path.GetFileNameWithoutExtension(outPath) + "_INTERMEDIATE_DATA";
            Directory.CreateDirectory(extraOutPath);
            extraOutPath += "/intermediate_" + outFileName;

            //Create the lists for storing data
            entries = new List<Entry>();
            ratsById = new Dictionary<string, RatById>();
            ratIds = new List<string>();
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
         * @TODO  Refactor this nasty method
         * 
         * Prints the entries to an excel file sorted by rat ID, with avgs calculated
         * for each stimulus
         */
        public void printIntermediateData()
        {
            // Create the file using the FileInfo object
            var file = new FileInfo(extraOutPath);
            if (file.Exists)
            {
                file.Delete();  // ensures we create a new workbook
                file = new FileInfo(extraOutPath);
            }

            //Create the Excel package and make a new workbook
            ExcelPackage pck = new ExcelPackage(file);

            foreach(string ratId in ratIds)
            {
                //Give each rat ID its own worksheet. Makes the file more readable
                ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add(ratId.Replace("\0", string.Empty));
                setupIntermediateHeaders(worksheet);

                //List used to store the 6 values for each stim value. Passed to markIntermediateAbnormal to help calc std deviation
                List<double> curValsList = new List<double>();
                int row = 2; //Start on 2nd row since there are column headers
                int p120Cnt = 0;
                foreach(string curStim in ratStims)
                {
                    //The variables used to calculates averages for each stimulus
                    double curSum = 0;
                    int entryCnt = 0;
                    double curAvg = 0;

                    foreach(Entry curEntry in ratsById[ratId].entries)
                    {
                        //string entryStim = System.Text.RegularExpressions.Regex.Replace(curEntry.colG, @"\s+", "");
                        string entryStim = curEntry.colG;
                        if (curStim.CompareTo(entryStim) == 0)
                        {
                            //Add its value to the curValsList
                            curValsList.Add(curEntry.colM);

                            worksheet.Cells[row, 1].Value = curEntry.colA;
                            worksheet.Cells[row, 2].Value = curEntry.colB;
                            worksheet.Cells[row, 3].Value = curEntry.colC;
                            worksheet.Cells[row, 4].Value = curEntry.colD;
                            worksheet.Cells[row, 5].Value = curEntry.colE;
                            worksheet.Cells[row, 6].Value = curEntry.colF;
                            worksheet.Cells[row, 7].Value = curEntry.colG;
                            worksheet.Cells[row, 8].Value = curEntry.colH;
                            worksheet.Cells[row, 9].Value = curEntry.colI;
                            worksheet.Cells[row, 10].Value = curEntry.colM;
                            
                            //Keep track of the avg of this set of entries with common stimulus
                            curSum += curEntry.colM;
                            entryCnt++;

                            if (curStim.Contains("p120"))
                            {
                                //If we are on our 6th p120 in a row
                                if (p120Cnt == 5 || p120Cnt == 11 || p120Cnt == 17)
                                {
                                    //Then print the avg 
                                    curAvg = curSum / 6;
                                    worksheet.Cells[row, 11].Value = Math.Round(curAvg, 2);
                                    worksheet.Cells[row, 11].Style.Font.Color.SetColor(Color.Red);
                                    //And add it to the avg list to be used in std deviation calculation and add it to the avg list for this stimulus value
                                    switch(p120Cnt)
                                    {
                                        case 5:
                                            ratsById[ratId].addAvg(Constants.P120_BEFORE_STR, curAvg);
                                            break;
                                        case 11:
                                            ratsById[ratId].addAvg(Constants.P120_DURING_STR, curAvg);
                                            break;
                                        case 17:
                                            ratsById[ratId].addAvg(Constants.P120_AFTER_STR, curAvg);
                                            break;
                                        default:
                                            break;
                                    }

                                    //Find and highlight any abnormal data
                                    markIntermediateAbnormal(worksheet, row, curValsList, curAvg);

                                    //Then reset the counting vals
                                    curSum = 0;
                                    entryCnt = 0;
                                    p120Cnt++;
                                    row++;
                                    curValsList.Clear();
                                }
                                else
                                {
                                    p120Cnt++;
                                }
                            }

                            row++;
                        }
                    }

                    //Since the p120's have already been printed, don't re-print them
                    if (!curStim.Contains("p120"))
                    {
                        //Print the avg 
                        curAvg = curSum / entryCnt;
                        worksheet.Cells[row - 1, 11].Value = Math.Round(curAvg, 2);
                        worksheet.Cells[row - 1, 11].Style.Font.Color.SetColor(Color.Red);
                        //And add it to the avg list to be used in std deviation calculation
                        ratsById[ratId].addAvg(curStim, curAvg);
                        //Find and highlight any abnormal data
                        markIntermediateAbnormal(worksheet, row - 1, curValsList, curAvg);

                    }

                    //Reset the current values list now that we are done with this stim val
                    curValsList.Clear();
                    row++;
                }

                resizeCols(worksheet);
            }

            pck.Save(); //And save
        }

        /**
         * Sets up the column headers for the intermediate file
         * 
         * @param  worksheet  The master Excel worksheet we are working on
         */
        public void setupIntermediateHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Rat ID #";
            worksheet.Cells[1, 2].Value = "Time";
            worksheet.Cells[1, 4].Value = "Rat ID";
            worksheet.Cells[1, 7].Value = "Stimulus";
            worksheet.Cells[1, 9].Value = "Trial Number";
            worksheet.Cells[1, 10].Value = "Value";
            worksheet.Cells[1, 11].Value = "Average";
            worksheet.Cells[1, 12].Value = "Std deviation";

            worksheet.Cells[1, 11].Style.Font.Color.SetColor(Color.Red);
            worksheet.Cells[1, 12].Style.Font.Color.SetColor(Color.Green);

            worksheet.Cells[3, 14].Value = "View different rats by selecting the worksheet with the desired rat ID";
            worksheet.Cells[3, 14].Style.Font.Bold = true;
            worksheet.Cells[3, 14].Style.Font.Italic = true;

            worksheet.Cells[4, 14].Value = "Red highlighted entries have a value >= 2 standard deviations from the mean";
            worksheet.Cells[4, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[4, 14].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);

            worksheet.Cells[5, 14].Value = "Yellow highlighted entries have a value >= 1 standard deviation from the mean";
            worksheet.Cells[5, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[5, 14].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);

            worksheet.Cells["A1:L1"].Style.Font.Bold = true;

            resizeCols(worksheet);
            increaseColSize(worksheet, 10);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            worksheet.View.FreezePanes(2, 1);
        }

        /**
         * Highlights any abnormal entries in the intermediate data worksheet. 
         * Entries are considered abnormal if they are > 1 SD away from the mean
         *         -> Yellow highlight is 1SD <= val < 2SD
         *         ->    Red highlight is >= 2SD
         *         
         * @param  worksheet  The master Excel worksheet we are working on
         * @param  row        The last row index of the values we are working with
         * @param  vals       The list of values we are working with
         * @param  avg        The avg of the list of vals provided
         */
        public void markIntermediateAbnormal(ExcelWorksheet worksheet, int row, List<double> vals, double avg)
        {
            double stdDev = calcStdDev(vals);

            //Print the std devaition
            worksheet.Cells[row, 12].Value = Math.Round(stdDev, 2);
            worksheet.Cells[row, 12].Style.Font.Color.SetColor(Color.Green);
            
            for (int i =0; i < 6; i++)
            {
                double diff = double.Parse(worksheet.Cells[row, 10].Value.ToString()) - avg;
                if (Math.Abs(diff) >= (2*stdDev))
                {
                    string rowString = "A" + row + ":J" + row;
                    worksheet.Cells[rowString].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[rowString].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                }
                else if (Math.Abs(diff) >= (stdDev))
                {
                    string rowString = "A" + row + ":J" + row;
                    worksheet.Cells[rowString].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[rowString].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                }

                row--;
            }
        }

        /**
         * Prints the pre-calculated rat data (avgs for all of the stim values for each rat) 
         * to a well-formatted Master excel sheet. This file will be located at the ouputPath provided
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

            setupMasterHeaders(worksheet);

            printRatInfoToMaster(worksheet);

            pck.Save();
        }

        /**
         * Fills in the cells in the master fie with the previously calculated data
         * 
         * @param  worksheet  The master Excel worksheet fill in the data for
         */
        public void printRatInfoToMaster(ExcelWorksheet worksheet)
        {
            int row = 2;
            foreach (string ratId in ratIds)
            {
                worksheet.Cells[row, Constants.RAT_ID].Value = ratId;
                worksheet.Cells[row, Constants.STRAIN].Value = "";    //Don't have a value for this
                worksheet.Cells[row, Constants.WEIGHT].Value = "";    //Don't have a value for this
                worksheet.Cells[row, Constants.P120_BEFORE].Value = Math.Round(ratsById[ratId].getAvg(Constants.P120_BEFORE_STR), 2);
                worksheet.Cells[row, Constants.P120_DURING].Value = Math.Round(ratsById[ratId].getAvg(Constants.P120_DURING_STR), 2);
                worksheet.Cells[row, Constants.P120_AFTER].Value = Math.Round(ratsById[ratId].getAvg(Constants.P120_AFTER_STR), 2);
                worksheet.Cells[row, Constants.NO_STIM].Value = Math.Round(ratsById[ratId].getAvg(Constants.NO_STIM_STR), 2);
                worksheet.Cells[row, Constants.PP3_ALONE].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP3_ALONE_STR), 2);
                worksheet.Cells[row, Constants.PP6_ALONE].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP6_ALONE_STR), 2);
                worksheet.Cells[row, Constants.PP12_ALONE].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP12_ALONE_STR), 2);
                worksheet.Cells[row, Constants.PP3_30].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP3_30_STR), 2);
                worksheet.Cells[row, Constants.PP6_30].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP6_30_STR), 2);
                worksheet.Cells[row, Constants.PP12_30].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP12_30_STR), 2);
                worksheet.Cells[row, Constants.PP3_50].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP3_50_STR), 2);
                worksheet.Cells[row, Constants.PP6_50].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP6_50_STR), 2);
                worksheet.Cells[row, Constants.PP12_50].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP12_50_STR), 2);
                worksheet.Cells[row, Constants.PP3_80].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP3_80_STR), 2);
                worksheet.Cells[row, Constants.PP6_80].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP6_80_STR), 2);
                worksheet.Cells[row, Constants.PP12_80].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP12_80_STR), 2);
                worksheet.Cells[row, Constants.PP3_140].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP3_140_STR), 2);
                worksheet.Cells[row, Constants.PP6_140].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP6_140_STR), 2);
                worksheet.Cells[row, Constants.PP12_140].Value = Math.Round(ratsById[ratId].getAvg(Constants.PP12_140_STR), 2);

                row++;
            }

            string avgRow = "A" + row + ":V" + row;
            worksheet.Cells[avgRow].Style.Font.Color.SetColor(Color.Red);
            worksheet.Cells[row, 2].Value = "Average";
            row++;

            string devRow = "A" + row + ":V" + row;
            worksheet.Cells[devRow].Style.Font.Color.SetColor(Color.Green);
            worksheet.Cells[row, 2].Value = "Std deviation";
            row++;

            string cntRow = "A" + row + ":V" + row;
            worksheet.Cells[cntRow].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[row, 2].Value = "Count";          

            printMasterTotals(worksheet);
        }

        /**
         * Prints the average, std deveiation, and counts for each of the stimulus columns
         * on the Master ouptut file
         * 
         * @param  worksheet  The Excel worksheet we want to print the data on
         */
        public void printMasterTotals(ExcelWorksheet worksheet)
        {
            int curRow = 2;  //The first row with relevant data will be row 2 (The first rat)
            int curCol = 4;  //The first columns with relevant data will be column 4 (P120 before)
            int finalCol = 22; //The final column we want to calculate date for (Col V, pp12[140ms])

            List<double> colVals = new List<double>();
            //Walk down the columns and gather the data
            while(curCol <= finalCol)
            {
                int rowsProcessed = 0;
                while (rowsProcessed < ratIds.Count)
                {
                    colVals.Add((double)worksheet.Cells[curRow, curCol].Value);
                    rowsProcessed++;
                    curRow++;
                }

                //Print the values for this column
                worksheet.Cells[curRow, curCol].Value = Math.Round(colVals.Average(), 2);
                curRow++;
                worksheet.Cells[curRow, curCol].Value = Math.Round(calcStdDev(colVals), 2);
                curRow++;
                worksheet.Cells[curRow, curCol].Value = colVals.Count;

                curRow = 2;      //Reset to the first row.
                colVals.Clear(); //Clear the column list
                curCol++;        //Increment the column
            }
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
         * Increases the size of the coumns in the given worksheet by the provided
         * integer size.
         * 
         * @param  worksheet  The Excel worksheet we want to resize columns on
         * @param  increase   The size amount we want to increase the columns by
         */
        public void increaseColSize(ExcelWorksheet worksheet, int increase)
        {
            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                worksheet.Column(i).Width += increase;
            }
        }

        /**
         * Sets up the given worksheet as the master worksheet with the proper column headers
         * and format for easy reading
         * 
         * @param  worksheet  The Excel worksheet we want setup
         */
        public void setupMasterHeaders(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Rat ID #";
            worksheet.Cells[1, 2].Value = "Strain";
            worksheet.Cells[1, 3].Value = "Weights";
            worksheet.Cells[1, 4].Value = "P120 before";
            worksheet.Cells[1, 5].Value = "P120 during";
            worksheet.Cells[1, 6].Value = "P120 after";
            worksheet.Cells[1, 7].Value = "No stimulus";
            worksheet.Cells[1, 8].Value = "pp3 alone";
            worksheet.Cells[1, 9].Value = "pp6 alone";
            worksheet.Cells[1, 10].Value = "pp12 alone";
            worksheet.Cells[1, 11].Value = "pp3 (30 ms)";
            worksheet.Cells[1, 12].Value = "pp6 (30 ms)";
            worksheet.Cells[1, 13].Value = "pp12 (30 ms)";
            worksheet.Cells[1, 14].Value = "pp3 (50 ms)";
            worksheet.Cells[1, 15].Value = "pp6 (50 ms)";
            worksheet.Cells[1, 16].Value = "pp12 (50 ms)";
            worksheet.Cells[1, 17].Value = "pp3 (80 ms)";
            worksheet.Cells[1, 18].Value = "pp6 (80 ms)";
            worksheet.Cells[1, 19].Value = "pp12 (80 ms)";
            worksheet.Cells[1, 20].Value = "pp3 (140 ms)";
            worksheet.Cells[1, 21].Value = "pp6 (140 ms)";
            worksheet.Cells[1, 22].Value = "pp12 (140 ms)";

            worksheet.Cells["A1:V1"].Style.Font.Bold = true;

            resizeCols(worksheet);
            increaseColSize(worksheet, 6);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            
            worksheet.View.FreezePanes(2, 3);
        }

        /**
         * Calculates the standard deviation of a list of numbers
         * 
         * @param  nums  A list of numbers that you want the std deviaton of
         * 
         * @return  The standard devaition of the numbers in the parameter list
         *          0 if the list is null
         */
        public double calcStdDev(List<double> nums)
        {
            if (null == nums)
            {
                return 0;
            }
            double mean = nums.Average();
            double sumOfSquaresOfDifferences = nums.Select(val => (val - mean) * (val - mean)).Sum();
            double sd = Math.Sqrt(sumOfSquaresOfDifferences / (nums.Count-1));

            return sd;
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
         * stored in a dictionary along with other RatById objects for all the other IDs.
         * Also creates a list of all stimulus values at the same time.
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
                ratsById[curRatId].entries.Sort();
            }

            //Then sorth the ratStim values so that they are alpahbetical
            ratStims.Sort();

            //And sort the rat IDs so that they are printed in logical order
            ratIds.Sort();
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
                    entries.Clear(); //Clear the list
                    getCsvEntries(); //And try again.
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

        /**
         * Returns the string representation of the path to the extra file
         */
        public string getExtraOutPath()
        {
            return extraOutPath;
        }

        /**
         * Returns the string representation of the output file name
         */
        public string getOutFileName()
        {
            return outFileName;
        }
    }
}
