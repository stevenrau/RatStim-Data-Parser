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
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace excelParse
{
    public partial class Home : Form
    {
        public Home()
        {
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
                //Display the path in the text box
                this.inPathDisplay.Text = openFileDialog.FileName;
            }
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
            ParseAndPrint myParser = new ParseAndPrint(this.inPathDisplay.Text, this.outPathDisplay.Text);

            myParser.printToExcelSorted();
            MessageBox.Show("Success! The ouput file was saved to the specified path.", "Success",
                             MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

    }
}
