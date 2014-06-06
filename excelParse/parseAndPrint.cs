using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        //We will be parsing .csv files with 17 columns
        const int NUM_CSV_COLS = 17;

        private string outPath;
        private string inPath;
        List<Entry> entries;

        /**
         * Constructor for the parseAndPrint class.
         * 
         * @param input  String representation of the input path
         *        output String representation of the output path
         */
        public ParseAndPrint(string input, string output)
        {
            inPath = input;
            outPath = output;
            entries = new List<Entry>();
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
                var reader = new StreamReader(File.OpenRead(inPath));
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
                MessageBox.Show(inPath+" is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                outFile.Close();
            }

        }

        public void printToExcel()
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

            pck.Save();
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
                MessageBox.Show(inPath + " is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                outFile.Close();
            }

        }

        /**
         * Method to get the entries from the csv input file and store them in 
         * the entries list
         */
        public void getCsvEntries()
        {
            //Start reading from the input file
            try
            {
                var reader = new StreamReader(File.OpenRead(inPath));
                while (!reader.EndOfStream)
                {
                    //Read in an entire line
                    var line = reader.ReadLine();
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
                MessageBox.Show(inPath + " is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            entries.Sort();
        }

        public string getInputPath()
        {
            return inPath;
        }

        public string getOutputPath()
        {
            return outPath;
        }

    }
}
