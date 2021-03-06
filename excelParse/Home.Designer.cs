﻿namespace RatStim
{
    partial class Home
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Home));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.browse = new System.Windows.Forms.Button();
            this.saveAs = new System.Windows.Forms.Button();
            this.browseLabel = new System.Windows.Forms.Label();
            this.saveAsLabel = new System.Windows.Forms.Label();
            this.outPathDisplay = new System.Windows.Forms.TextBox();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.sortButton = new System.Windows.Forms.Button();
            this.inPathDisplay = new System.Windows.Forms.TextBox();
            this.clearInputsFilesButton = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(645, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F)));
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            this.fileToolStripMenuItem.Click += new System.EventHandler(this.fileToolStripMenuItem_Click);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.openToolStripMenuItem.Text = "&Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Q)));
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.exitToolStripMenuItem.Text = "&Quit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click_1);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "CSV Files|*.csv";
            this.openFileDialog.InitialDirectory = "%HOMEDRIVE%%HOMEPATH%";
            this.openFileDialog.Multiselect = true;
            this.openFileDialog.Title = "Choose input file";
            // 
            // browse
            // 
            this.browse.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.browse.Location = new System.Drawing.Point(143, 61);
            this.browse.Name = "browse";
            this.browse.Size = new System.Drawing.Size(137, 35);
            this.browse.TabIndex = 1;
            this.browse.Text = "Browse";
            this.browse.UseVisualStyleBackColor = true;
            this.browse.Click += new System.EventHandler(this.Browse_Click);
            // 
            // saveAs
            // 
            this.saveAs.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveAs.Location = new System.Drawing.Point(143, 210);
            this.saveAs.Name = "saveAs";
            this.saveAs.Size = new System.Drawing.Size(137, 35);
            this.saveAs.TabIndex = 2;
            this.saveAs.Text = "Save As...";
            this.saveAs.UseVisualStyleBackColor = true;
            this.saveAs.Click += new System.EventHandler(this.saveAs_Click);
            // 
            // browseLabel
            // 
            this.browseLabel.AutoSize = true;
            this.browseLabel.Font = new System.Drawing.Font("Calibri", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.browseLabel.Location = new System.Drawing.Point(12, 67);
            this.browseLabel.Name = "browseLabel";
            this.browseLabel.Size = new System.Drawing.Size(124, 21);
            this.browseLabel.TabIndex = 3;
            this.browseLabel.Text = "Input files (.csv):";
            // 
            // saveAsLabel
            // 
            this.saveAsLabel.AutoSize = true;
            this.saveAsLabel.Font = new System.Drawing.Font("Calibri", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveAsLabel.Location = new System.Drawing.Point(3, 216);
            this.saveAsLabel.Name = "saveAsLabel";
            this.saveAsLabel.Size = new System.Drawing.Size(133, 21);
            this.saveAsLabel.TabIndex = 4;
            this.saveAsLabel.Text = "Output file (.xlsx):";
            // 
            // outPathDisplay
            // 
            this.outPathDisplay.BackColor = System.Drawing.SystemColors.Info;
            this.outPathDisplay.Location = new System.Drawing.Point(300, 218);
            this.outPathDisplay.Multiline = true;
            this.outPathDisplay.Name = "outPathDisplay";
            this.outPathDisplay.ReadOnly = true;
            this.outPathDisplay.Size = new System.Drawing.Size(321, 70);
            this.outPathDisplay.TabIndex = 6;
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.FileName = "outputFile.xlsx";
            this.saveFileDialog.Filter = "Excel .xlsx|*.xlsx";
            this.saveFileDialog.Title = "Choose save file location";
            // 
            // sortButton
            // 
            this.sortButton.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sortButton.Location = new System.Drawing.Point(456, 310);
            this.sortButton.Name = "sortButton";
            this.sortButton.Size = new System.Drawing.Size(165, 35);
            this.sortButton.TabIndex = 8;
            this.sortButton.Text = "Sort";
            this.sortButton.UseVisualStyleBackColor = true;
            this.sortButton.Click += new System.EventHandler(this.sortButton_Click);
            // 
            // inPathDisplay
            // 
            this.inPathDisplay.BackColor = System.Drawing.SystemColors.Info;
            this.inPathDisplay.Location = new System.Drawing.Point(300, 61);
            this.inPathDisplay.Multiline = true;
            this.inPathDisplay.Name = "inPathDisplay";
            this.inPathDisplay.ReadOnly = true;
            this.inPathDisplay.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.inPathDisplay.Size = new System.Drawing.Size(321, 133);
            this.inPathDisplay.TabIndex = 7;
            // 
            // clearInputsFilesButton
            // 
            this.clearInputsFilesButton.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearInputsFilesButton.Location = new System.Drawing.Point(143, 114);
            this.clearInputsFilesButton.Name = "clearInputsFilesButton";
            this.clearInputsFilesButton.Size = new System.Drawing.Size(137, 35);
            this.clearInputsFilesButton.TabIndex = 9;
            this.clearInputsFilesButton.Text = "Clear input files";
            this.clearInputsFilesButton.UseVisualStyleBackColor = true;
            this.clearInputsFilesButton.Click += new System.EventHandler(this.clearInputsFilesButton_Click);
            // 
            // Home
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(645, 367);
            this.Controls.Add(this.clearInputsFilesButton);
            this.Controls.Add(this.sortButton);
            this.Controls.Add(this.inPathDisplay);
            this.Controls.Add(this.outPathDisplay);
            this.Controls.Add(this.saveAsLabel);
            this.Controls.Add(this.browseLabel);
            this.Controls.Add(this.saveAs);
            this.Controls.Add(this.browse);
            this.Controls.Add(this.menuStrip1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(1920, 1080);
            this.MinimumSize = new System.Drawing.Size(661, 272);
            this.Name = "Home";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "RatStim Data Parser";
            this.Load += new System.EventHandler(this.Home_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button browse;
        private System.Windows.Forms.Button saveAs;
        private System.Windows.Forms.Label browseLabel;
        private System.Windows.Forms.Label saveAsLabel;
        private System.Windows.Forms.TextBox outPathDisplay;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Button sortButton;
        private System.Windows.Forms.TextBox inPathDisplay;
        private System.Windows.Forms.Button clearInputsFilesButton;
    }
}

