using HP.SDLCToolsBO.Model;
using HP.SDLCToolsBO.Framework;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Newtonsoft.Json;
using System.Web;


namespace HP.SDLCToolsBO.BO
{
    public class ExcelHandlerBO : ExcelHandler
    {
        private static string sheetName = "Lists$"; // Lists is the default sheet name in ListAdministrator

        private ALMConnection almConnection;
        private DataTable excelTable;
        private string domain;
        private string project;

        private ExcelReaderWriter xlsxReaderWriter;

        /// <summary>
        /// Gets the number of list items in an Excel file 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>int Number of list items</returns>
        public override int CountListItems(String fileName)
        {
            excelTable = xlsxReaderWriter.LoadExcelFile(fileName, sheetName);
            return excelTable.Rows.Count - 1;
        }

        //Export to Excel

        #region Variables

        //private Hashtable positionHash = new Hashtable();

        #endregion Variables

        #region Methods

        /// <summary>
        /// A configuration method to set the parameters required by the ExcelWriter
        /// </summary>
        /// <param name="list">A list of ALM lists</param>
        /// <returns>IDictionary Configuration variables encapsulated in a Dictionary</returns>
        private IDictionary<string, string> setWriterConfiguration(List<string> list)
        {
            IDictionary<string, string> config = new Dictionary<string, string>();
            string TempPath = @"C:\excelfiles\";
            Random rnd = new Random();
            string FileName = list[0] + "_" + (DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond.ToString() + rnd.Next(1000).ToString()).Replace('/', '-');
            FileName = FileName.Replace(':', '_');

            config.Add("SavePath", TempPath + FileName + ".xlsx");
            config.Add("ReturnPath", "/web/Excel/" + FileName + ".xlsx");
            config.Add("SheetName", "Lists");

            return config;
        }

        /// <summary>
        /// Metadata configuration to be used by the ExcelWriter
        /// </summary>
        /// <returns>IDictionary Metadata configuration encapsulated in a Dictionary</returns>
        private IDictionary<string, string> SetWriterMetadata()
        {
            IDictionary<string, string> Metadata = new Dictionary<string, string>();

            Metadata.Add("Author", "Omagana");
            Metadata.Add("Keywords", "Office Open XML");
            Metadata.Add("Company", "HPE");

            return Metadata;
        }


        /// <summary>
        /// Generates a Excel file  with the information from the list
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="columnsNames"></param>
        /// <param name="separationChar"></param>
        /// <returns></returns>
        public override string ExportToExcel(List<string> rows, List<string> columnsNames, char separationChar)
        {
            string Message = "";

            try
            {
                //  System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

                if (rows.Count > 0)
                {
                    Message = FillExcelSheet(rows, columnsNames, separationChar);
                }
                return Message;
            }
            catch (Exception e)
            {
                Message = "Error: " + e.Message;
                return Message;
            }
        }

        /// <summary>
        /// Calls the ExcelWriter to fill the Excel file with data and name the file
        /// </summary>
        /// <param name="rows">the information of each row with a separation char for each column</param>
        /// <param name="names">Title of the Columns</param>
        /// <param name="separationChar">Char that separates the information in the rows for each column</param>
        /// <returns>string Excel file fully qualified name</returns>
        public string FillExcelSheet(List<string> rows, List<string> names, char separationChar)
        {
            IDictionary<string, string> config = null;
            try
            {
                //positionHash = new Hashtable();
                //  System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

                config = setWriterConfiguration(names);
                FileInfo newFile = new FileInfo(config["SavePath"]);
                //xlsxReaderWriter = new ExcelReaderWriter(config);
                xlsxReaderWriter.setConfiguration(config);
                xlsxReaderWriter.SetMetadata(SetWriterMetadata());
                var data = FillRows(rows, names, separationChar);
                xlsxReaderWriter.WriteSheet(newFile, data);
                return config["ReturnPath"];
            }
            catch (Exception e)
            {

                return string.Format("Error: {0}, SavePath: {1}, ReturnPath{2}", e.Message, config["SavePath"], config["ReturnPath"]);
            }
        }


        /// <summary>
        /// fills each row of the file, first with the colum names in the "names" paramter and then with the rows
        /// </summary>
        /// <param name="list">information for each row that contains a separation with a char to make the split</param>
        /// <param name="names">name of each column</param>
        /// <param name="separationChar">Character that separates each column in the names list</param>
        /// <returns></returns>
        private List<List<string>> FillRows(List<string> list, List<string> names, char separationChar)
        {
            List<List<string>> data = new List<List<string>>();
            //Fill The Columns
            data.Add(names);

            foreach (string item in list) //Foreach of items that are selected
            {
               
                    string[] e = item.Split(separationChar);

                    List<string> info = new List<string>(e);
                    data.Add(info);

            }

            Console.WriteLine(data);
            return data;
        }

        /// <summary>
        /// Sets the data matrix values. Recursive to enable the generation of sub-lists
        /// </summary>
        /// <param name="precedingCols">To "indent" the data</param>
        /// <param name="customListNode">ALM list data node</param>
        /// <param name="data">The matrix to be filled with data</param>
        private void FillCells(int precedingCols, CustomizationListNode customListNode, List<List<string>> data)
        {
            if (customListNode.ChildrenCount > 0)
            {
                List childList = new List();
                childList = customListNode.Children;

                for (int i = 0; i < childList.Count; i++)
                {
                    CustomizationListNode childListNode = (CustomizationListNode)childList[i + 1];
                    List<string> row = new List<string>();

                    for (int j = 0; j < precedingCols; j++)
                        row.Add("");

                    row.Add(childListNode.Name.ToString());
                    data.Add(row);

                    if (childListNode.Children.Count > 0)
                        FillCells(precedingCols + 1, childListNode, data);
                }
            }
        }



        #endregion

    }
}
