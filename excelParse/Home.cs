/**************************************************************************\
Module Name:   Home.cs 
Project:       excelParse
Author:        Steven Rau

This file conatins the action listeners for the buttons/textfields, etc.
on the Home window.
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace RatStim
{
    public partial class Home : Form
    {
        //A list to store all the input csv paths to pass to ParseAndPrint when the Sort button is clicked
        public List<string> inPaths;
        //A count of input files. Used for ouput list numbering and to pass to ParseAndPrint
        public int inPathCount;

        public Home()
        {
            inPaths = new List<string>();
            inPathCount = 0;
            InitializeComponent();
        }

        private void Home_Load(object sender, EventArgs e)
        {

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {
 
        }

        /**
         * Simple "Quit" button in the File drop down menu
         */
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /**
         * The "Browse" button. Opens a file dialog box and allows the user to select a 
         * file to open and parse. If the path does not exist, a warning informs the uer of the error and
         * forces them to pick another valid path.
         * This should be a .csv file
         */
        private void Browse_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK) 
            {
                //Make sure they selected a .csv file
                if (!openFileDialog.FileName.GetLast(3).Equals("csv"))
                {
                    MessageBox.Show("You must select a .csv file", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Browse_Click(sender, e);
                }
                //Display the inputs paths in the text box
                foreach(string curFileName in openFileDialog.FileNames)
                {
                    inPathCount++;
                    inPathDisplay.Text += inPathCount + ") " + curFileName + Environment.NewLine + Environment.NewLine;
                    inPaths.Add(curFileName);
                }
            }
        }

        /** 
         * Button click to clear all of the input files that were previously selected
         */
        private void clearInputsFilesButton_Click(object sender, EventArgs e)
        {
            inPathCount = 0;
            inPaths.Clear();
            inPathDisplay.Text = "";
        }

        /**
         * The "Save as" button. Opens a file dialog box and allows the user to
         * choose a pathname to save the output file to. If the file already exists,
         * display a message ensuring they want to overwrite
         * This should be a .xlsx file
         */
        private void saveAs_Click(object sender, EventArgs e)
        {
            // Get the save dialog and get the path
            DialogResult result = saveFileDialog.ShowDialog();
            if (result == DialogResult.OK) 
            {
                //Make sure they selected a .xlsx or .xls file to save to
                if (!saveFileDialog.FileName.GetLast(4).Equals("xlsx") && !saveFileDialog.FileName.GetLast(3).Equals("xls"))
                {
                    MessageBox.Show("You must save as a .xlsx or .xls file", "Error",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    saveAs_Click(sender, e);
                }
                //Display the path in the text box
                this.outPathDisplay.Text = saveFileDialog.FileName;
            }
        }

        /**
         * The 'Open' option in the File dropdown menu. Simply does the
         * functionality of the "Browse" button
         */
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Browse_Click(sender, e);
        }

        private void sortButton_Click(object sender, EventArgs e)
        {
            if (0 == inPaths.Count)
            {
                DialogResult result = MessageBox.Show("No input file(s) selected.",
                                                      "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if(this.outPathDisplay.Text.CompareTo("") == 0)
            {
                DialogResult result = MessageBox.Show("No output file selected.",
                                                      "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ParseAndPrint myParser = new ParseAndPrint(inPaths, inPathCount, this.outPathDisplay.Text);

            try
            {
                myParser.printIntermediateData();
                myParser.printMasterData();

                SuccessWindow successWindow = new SuccessWindow(this.outPathDisplay.Text);
                successWindow.Show();
            }
            catch(IOException)
            {
                DialogResult result = MessageBox.Show("A file with the same name as the output path specified is currently in use by another process. " +
                                                           "Close it to continue.", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (DialogResult.OK == result)
                {
                    sortButton_Click(sender, e);
                }
                else
                {
                    return;
                }
            }      
        }

        private void readmeToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
