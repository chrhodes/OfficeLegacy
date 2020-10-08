using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Collections;
namespace SupportTools_PowerPoint.Data
{

    public partial class ApplicationDataSet
    {
        #region Helpers

        /// <summary>
        /// Search dataTable for a row with the specified RowId and delete it.
        /// </summary>
        /// <param name="rowId"></param>
        /// <param name="dataTable"></param>
        public static void DeleteByRowId(int rowId, DataTable dataTable)
        {
            DataRow[] foundRows = dataTable.Select(string.Format("RowId = {0}", rowId));

            if (foundRows.GetLength(0) != 1)
            {
                throw new ApplicationException("DeleteByRowId() Fatal Error");
            }
            else
            {
                foundRows[0].Delete();
            }
        }

        /// <summary>
        /// Search dataTable for a row with the specified rowId and return it.
        /// </summary>
        /// <param name="rowId"></param>
        /// <param name="dataTable"></param>
        public static DataRow FindByRowId(int rowId, DataTable dataTable)
        {
            DataRow[] foundRows = dataTable.Select(string.Format("RowId = {0}", rowId));

            if (foundRows.GetLength(0) != 1)
            {
                throw new ApplicationException("FindByRowId() Fatal Error");
            }
            else
            {
                return foundRows[0];
            }
        }

        #endregion

        #region Load/Save Datatable to XML file sample code

        partial class DataTable1DataTable
        {
            public string InputFilePath { get; set; }
            public string OutputFilePath { get; set; }

            public void SaveToFile(DataTable1DataTable dataTable)
            {
                using (System.IO.StreamWriter streamWriter = new StreamWriter(OutputFilePath))
                {
                    SaveToStream(streamWriter, dataTable);
                }
            }
            /// <summary>
            /// Save DataTable1DataTable DataTable to stream.
            /// What do we want to return if anything?
            /// </summary>
            /// <param name="dataTable"></param>
            /// <param name="outputStream"></param>
            public void SaveToStream(StreamWriter outputStream, DataTable1DataTable dataTable)
            {
                // TODO: Modify this to reflect the particular elements and columns of interest
                try
                {
                    XElement outputXML =
                        new XElement("DataTable",
                            from DataTable1Row item in dataTable.Rows
                            select new XElement("Row",
                                new XElement("Column1", item.DataColumn1),
                                new XElement("Column2", item.DataColumn2),
                                new XElement("Column3", item.DataColumn3)
                            )
                        );

                    outputStream.Write(outputXML.ToString());
                    outputStream.Flush();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException("SaveToStream() fatal error", ex);
                }
            }

            public void LoadFromFile(DataTable1DataTable dataTable)
            {
                using (StreamReader streamReader = new StreamReader(InputFilePath))
                {
                    LoadFromStream(streamReader, dataTable);
                }
            }
            /// <summary>
            /// Load DataTable1DataTable DataTable from stream.
            /// What do we want to return if anything?
            /// </summary>
            /// <param name="dataTable"></param>
            /// <param name="inputStream"></param>
            public void LoadFromStream(StreamReader inputStream, DataTable1DataTable dataTable)
            {
                try
                {
                    XElement inputXML = XElement.Load(inputStream);

                    foreach (var item in inputXML.Elements("Row"))
                    {
                        DataTable1Row row = NewDataTable1Row();

                        // TODO: Modify this for the particular elements and columns of interest

                        row.DataColumn1 = (string)item.Element("Column1");
                        row.DataColumn2 = (string)item.Element("Column2");
                        row.DataColumn3 = (string)item.Element("Column3");

                        dataTable.AddDataTable1Row(row);
                    }
                }
                catch (Exception ex)
                {
                    throw new ApplicationException("LoadFromStream() fatal error", ex);
                }
            }

            public bool IsValidRow(DataTable1Row row)
            {
                bool isValid = true;

                // TODO: Add validation code

                return isValid;
            }
        }

        #endregion

    }
}
