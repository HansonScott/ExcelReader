/*
-----------------------------------------------------------------------------
Copyright (c) 2003 ClaimRemedi, Inc. All Rights Reserved.

PROPRIETARY NOTICE: This software has been provided pursuant to a
License Agreement that contains restrictions on its use. This software
contains valuable trade secrets and proprietary information of
ClaimRemedi, Inc and is protected by Federal copyright law.  It may not
be copied or distributed in any form or medium, disclosed to third parties,
or used in any manner that is not provided for in said License Agreement,
except with the prior written authorization of ClaimRemedi, Inc.

-----------------------------------------------------------------------------
$Log$
Revision 53  2012/02/23 14:16:22  shanson
  Added comment keyword

$Log: /Code/ClaimRemedi.ExcelReader.Client/ExcelReaderClientForm.cs $
 
 1     3/22/11 6:20p Nmayer
 VS 2010 Initial Checkin
 
 2     12/01/10 11:53a Shanson
 Added infrastructure for password protected file (doesn't work though)
 Changed results listView to DataGrid to simplify and expand
 functionality.
 Added handling for bad data files.
 
 1     10/04/10 11:06a Shanson
 Initial check in from transfer from VS40
 
 4     8/31/10 10:18a Shanson
 Added check box functionality of reloading the file on change.
 
 3     8/31/10 10:05a Shanson
 Added check box for header row option, moved main method into form.cs

-----------------------------------------------------------------------------
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ClaimRemedi.ExcelReader.Client
{
    public partial class ExcelReaderClientForm : Form
    {
        #region MAIN
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ExcelReaderClientForm());
        }
        #endregion

        public ExcelReaderClientForm()
        {
            InitializeComponent();
        }

        #region UI Event Methods
        private void button1_Click(object sender, EventArgs e)
        {
            // file open dialog
            FileDialog fd = new OpenFileDialog();
            DialogResult dr = fd.ShowDialog();
            if (dr != DialogResult.OK) { return; }

            // capture filename
            string FileName = fd.FileName;

            // set filename to text box
            this.textBox1.Text = FileName;

            label1.Text = "File name captured.";
            label1.Refresh();
        }
        private void btn_Run_Click(object sender, EventArgs e)
        {
            label1.Text = "Processing...";
            label1.Refresh();
            
            CreateReaderAndLoadFile(this.textBox1.Text);
        }
        #endregion

        #region Private Functions
        private void CreateReaderAndLoadFile(string FileName)
        {
            // create reader object
            Reader r = null;

            if (tb_Password.Text.Length > 0)
            {
                r = new Reader(FileName, tb_Password.Text);
            }
            else
            {
                r = new Reader(FileName);
            }
            label1.Text = "Reader object created, reading data from file...";
            label1.Refresh();

            // determine to use first line as headers
            r.FirstLineIsHeader = this.checkBox1.Checked;

            // read file
            DataSet data = r.ReadAllData();

            label1.Text = "Data read from file, populating list view...";
            label1.Refresh();

            // set data to list view
            // NOTE: assume first sheet for now...
            SetDataToListView(r);

            if (data != null && data.Tables.Count > 0)
            {
                label1.Text = "Data loaded successfully. Total Rows: " + data.Tables[0].Rows.Count;
            }
            else
            {
                label1.Text = "Data was unable to be loaded.";
            }
            label1.Refresh();
        }
        private void SetDataToListView(Reader r)
        {
            if (r == null || r.Data == null) { return; }

            DataTable data = r.Data.Tables[0];

            // catch bad or blank input
            if (data == null ||
                data.Rows.Count == 0 ||
                data.Columns.Count == 0) { return; }

            this.gv_results.DataSource = data;
        }
        #endregion
    } // end class
} // end namespace