using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ClaimRemedi.ExcelReader;

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
            try
            {
                DataSet data = r.ReadAllData();

                label1.Text = "Data read from file, populating list view...";
                label1.Refresh();

                // set data to list view
                // NOTE: assume first sheet for now...
                SetDataToListView(r);

                if (data != null && data.Tables.Count > 0)
                {
                    label1.Text = $"Data loaded successfully. Total Rows: {data.Tables[0].Rows.Count}";
                }
                else
                {
                    label1.Text = "Data was unable to be loaded.";
                }

                lblMetaData.Text = GetMetaDataFromDataSet(data);

            }
            catch(Exception e)
            {
                label1.Text = "Data was unable to be loaded.";
                MessageBox.Show(e.Message);
            }
            label1.Refresh();
        }

        private string GetMetaDataFromDataSet(DataSet data)
        {
            return $"table count: {data.Tables.Count}";
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