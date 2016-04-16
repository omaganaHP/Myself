using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.Data.OleDb;
using HP.SDLCToolsBO.Framework;

namespace HP.SDLCToolsBO.BO
{
    public class ExcelReaderWriter
    {
        private IDictionary<string, string> configuration;
        private IDictionary<string, string> metadata;
        private static string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";

        public void setConfiguration(IDictionary<string, string> config)
        {
            this.configuration = config;
        }

        public void SetMetadata(IDictionary<string, string> m)
        {
            this.metadata = m;
        }

        /// <summary>
        /// Writes the actual Excel file
        /// </summary>
        /// <param name="newFile">FileInfo file handle</param>
        /// <param name="data">List of lists Data matrix</param>
        public void WriteSheet(FileInfo newFile, List<List<string>> data)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(newFile)) //if i don't encapsulate the excel package the excel file contains garbage
            {
                xlPackage.DebugMode = true;
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add(configuration["SheetName"]);

                xlPackage.Workbook.Properties.Title = metadata["Title"];
                xlPackage.Workbook.Properties.Author = metadata["Author"];
                xlPackage.Workbook.Properties.Subject = metadata["Subject"];
                xlPackage.Workbook.Properties.Keywords = metadata["Keywords"];
                xlPackage.Workbook.Properties.Company = metadata["Company"];

                WriteData(data, worksheet);
                xlPackage.Save();
            }
        }

        /// <summary>
        /// Sets the Excel cell values
        /// </summary>
        /// <param name="data">Data matrix</param>
        /// <param name="worksheet">Worksheet object</param>
        public void WriteData(List<List<string>> data, ExcelWorksheet worksheet)
        {
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    string cellValue = string.Empty;
                    Stream buffer = new MemoryStream(System.Text.Encoding.Default.GetBytes(data[i][j]));
                    XmlSanitizingStream var = new XmlSanitizingStream(buffer);
                    cellValue = var.ReadLine();
                    cellValue = string.IsNullOrEmpty(cellValue) ? string.Empty : cellValue;
                    worksheet.Cell(i + 1, j + 1).Value = cellValue;
                }
            }
        }

    }
}
