/*
-----------------------------------------------------------------------------
Copyright (c) 2012 ClaimRemedi, Inc. All Rights Reserved.

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

$Log: /Code/ClaimRemedi.ExcelReader/Reader.cs $
 
 1     3/22/11 6:20p Nmayer
 VS 2010 Initial Checkin
 
 4     12/23/10 10:27a Ddaine
 Remove unused variable.
 
 3     12/17/10 1:50p Keddy
 When checking for file extension, turned ToLower before comparing
 
 2     12/01/10 11:51a Shanson
 Added infrastructure for password protected files (could not get it to
 work though)
 Added handling for duplicate sheet names.
 
 1     10/04/10 11:06a Shanson
 Initial check in from transfer from VS40
 
 2     8/31/10 10:06a Shanson
 Added handling for header row to set column names and remove itself
 from data set

-----------------------------------------------------------------------------
*/


using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using System.Data.OleDb;

namespace ClaimRemedi.ExcelReader
{
    public class Reader
    {
        public enum ExcelFileType
        {
            XLS_1997_to_2003 = 0,
            XLSX_2007 = 1,
            Unknown = 2,
        }

        #region Member Declarations
        private string m_SourceFilePath;
        private ExcelFileType m_ExcelFileType;
        private string m_sConn;
        private bool m_FirstLineHeaders;
        private DataSet m_data;
        private string m_FilePassword;
        #endregion

        #region Public Properties
        public string FilePath
        {
            get{return m_SourceFilePath;}
        }
        public string FilePassword
        {
            get { return m_FilePassword; }
            set { m_FilePassword = value; }
        }
        public ExcelFileType FileType
        {
            get { return m_ExcelFileType; }
            set
            {
                this.m_ExcelFileType = value;

                // set value to use first line as headers
                string HDR = "No";
                if (m_FirstLineHeaders)
                {
                    HDR = "Yes";
                }

                // based on the file type, set the connection string.
                switch (m_ExcelFileType)
                {
                    case ExcelFileType.Unknown:
                        m_sConn = "";
                        break;
                    case ExcelFileType.XLS_1997_to_2003:
                        // if it is xls - use Jet
                        m_sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                    "Data Source=" + this.FilePath + ";";
                        
                        // NOTE: this does not work for Excel, only for other DB files.
                        if(!String.IsNullOrEmpty(m_FilePassword))
                        {
                            m_sConn +=  "Password=" + m_FilePassword + ";";
                        }

                        m_sConn +=  "Extended Properties=\"Excel 8.0;" + 
                                    "HDR=" + HDR +";" + // treat first row as header
                                    "IMEX=1;\""; // 1 = look at only first 8 rows to determine datatype, 0 = look at all.
                        break;
                    case ExcelFileType.XLSX_2007:
                    default:
                        // if it is xlsx = use Ace
                        m_sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "Data Source=" + this.FilePath + ";";

                        // NOTE: this does not work for Excel, only for other DB files.
                        if(!String.IsNullOrEmpty(m_FilePassword))
                        {
                            m_sConn += "Password=" + m_FilePassword + ";";
                        }

                        m_sConn +=  "Extended Properties=\"Excel 12.0;" +
                                    "HDR=" + HDR + ";" + // treat first row as header
                                    "IMEX=1;\""; // 1 = look at only first 8 rows to determine datatype, 0 = look at all.
                        break;
                }
            }
        }
        public bool FirstLineIsHeader
        {
            set
            {
                m_FirstLineHeaders = value;
            }
            get
            {
                return m_FirstLineHeaders;
            }
        }
        public DataSet Data
        {
            get { return m_data; }
        }
        #endregion

        #region Constructor
        public Reader(string SourcefilePath) : this(SourcefilePath, "") { }
        public Reader(string SourcefilePath, string FilePassword)
        {
            // NOTE: password must be set before type.
            this.FilePassword = FilePassword;

            // set path
            this.m_SourceFilePath = SourcefilePath;
            string ext = Path.GetExtension(FilePath);

            // set type from extention
            if (ext.ToLower() == ".xls")
            {
                this.FileType = ExcelFileType.XLS_1997_to_2003;
            }
            else if (ext.ToLower() == ".xlsx")
            {
                this.FileType = ExcelFileType.XLSX_2007;
            }
            else
            {
                this.FileType = ExcelFileType.Unknown;
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Opens the file and reads all worksheets, building a data table for each sheet, and combining
        /// them into a data set.
        /// </summary>
        /// <returns>a dataset containing dataTables for each sheet in the workbook</returns>
        public DataSet ReadAllData()
        {
            OleDbConnection oConn = new OleDbConnection(m_sConn);
            try
            {
                oConn.Open();
            }
            catch (Exception e)
            {
                throw(e);
            }
            // NOTE: THEORY: the workbook names (sheet names) are the tables we select from
            // so getSchema reads the workbook names as the table names.
            DataTable Workbooks = oConn.GetSchema("Tables");

            // setup output dataSet
            DataSet results = new DataSet(Path.GetFileNameWithoutExtension(this.FilePath));

            // loop through all rows and build a new DataTable for each, adding it to our DataSet
            for (int i = 0; i < Workbooks.Rows.Count; i++)
            {
                DataRow WorkbookRow = Workbooks.Rows[i];
                String WorkbookName = WorkbookRow[2].ToString(); // 2 = TableName (after ID, I assume)

                // setup output dataTable
                DataTable result = new DataTable(WorkbookName);

                // Get the data from this workbookName
                string sSQL = "SELECT * FROM [" + WorkbookName + "]";
                OleDbCommand oCmd = new OleDbCommand(sSQL, oConn);
                OleDbDataAdapter oAdapter = new OleDbDataAdapter(oCmd);
                oAdapter.Fill(result);

                // if first row is header, then capture it.
                if (m_FirstLineHeaders)
                {
                    DataRow HeaderRow = result.Rows[0];
                    for (int c = 0; c < result.Columns.Count; c++)
                    {
                        result.Columns[c].ColumnName = HeaderRow[c].ToString();
                    }

                    // now remove the first line from the data
                    result.Rows.RemoveAt(0);
                }

                // Set dataTable back to dataSet
                //while (results.Tables.Contains(result.TableName))
                //{
                //    result.TableName = result.TableName + "+";
                //}

                results.Tables.Add(result);
            } // end for loop - tables

            // set results to internal data member
            m_data = results;

            // and return it.
            return results;
        }

        /// <summary>
        /// read all lines from worksheet number, essentially turning each value into a CSV line.
        /// </summary>
        /// <param name="filePath">The complete file path to open and read the file</param>
        /// <param name="WorksheetNumber">The sheet number, 0-based index.</param>
        /// <returns>string[] of comma-separated values from each row from the desired table index.</returns>
        public string[] GetCSVLinesFromSheetNumber(int WorksheetNumber)
        {
            // read the file
            DataSet data = ReadAllData();

            // catch an out of bounds request
            if (data.Tables.Count < WorksheetNumber) { return null; }

            // get just this sheet
            DataTable thisTable = data.Tables[WorksheetNumber];

            // setup output
            int rows = thisTable.Rows.Count;
            string[] results = new string[rows];

            // go through each row in this sheet
            for (int i = 0; i < rows; i++)
            {
                // setup output
                string thisRowValue = "";

                // go through each cell in this row
                for (int j = 0; j < thisTable.Columns.Count; j++)
                {
                    // for cell (i,j), capture this value
                    string thisCellValue = thisTable.Rows[i][j].ToString();

                    // add this cell value to the current string.
                    thisRowValue += "," + thisCellValue;
                } // end for loop - j - columns

                // add this row's value to the output (substring takes off the last comma)
                results[i] = thisRowValue.Substring(1);

            } // end for loop - i - rows
            
            // return built results
            return results;
        }
        #endregion
    }
}
