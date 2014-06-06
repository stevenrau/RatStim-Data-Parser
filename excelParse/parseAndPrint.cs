using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

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
        }

        /**
         * Prints the important values from the .csv input file to a specified text file
         * Mainly used for testing and debugging. Shouldn't be used in the final product.
         * 
         * @param output  String representation of the ouotput file. Needs to be a text file.
         */
        public void printToText(string output)
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
            catch (IOException e)
            {
                MessageBox.Show(inPath+" is currently in use by another process. Close it to continue.", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                outFile.Close();
            }

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
