using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace File.Poc
{
    public class FileService
    {
        public List<T> ListToExcelSax<T>(string path)
        {
            var watch = Stopwatch.StartNew();
            var list = new List<T>();
            var properties = typeof(T).GetProperties();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                //WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                    //row counter
                    int rcnt = 0;
                    while (reader.Read())
                    {
                        //find xml row element type 
                        //to understand the element type you can change your excel file eg : test.xlsx to test.zip
                        //and inside that you may observe the elements in xl/worksheets/sheet.xml
                        //that helps to understand openxml better
                        if (reader.ElementType == typeof(Row))
                        {
                            //create data table row type to be populated by cells of this row
                            //DataRow tempRow = dt.NewRow();
                            T obj = Activator.CreateInstance<T>();
                            //***** HANDLE THE SECOND SENARIO*****
                            //if row has attribute means it is not a empty row
                            if (reader.HasAttributes)
                            {
                                //read the child of row element which is cells
                                //here first element
                                reader.ReadFirstChild();
                                do
                                {
                                    //find xml cell element type 
                                    if (reader.ElementType == typeof(Cell))
                                    {
                                        Cell c = (Cell)reader.LoadCurrentElement();

                                        string cellValue;
                                        int actualCellIndex = GetColumnIndex(c);

                                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                                        {
                                            SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));
                                            cellValue = ssi.Text.Text;
                                        }
                                        else
                                        {
                                            cellValue = c.CellValue.InnerText;
                                        }

                                        //if row index is 0 its header so columns headers are added & also can do some headers check incase
                                        if (rcnt != 0)
                                        {
                                            // instead of tempRow[c.CellReference] = cellValue;
                                            var type = properties[actualCellIndex - 1].PropertyType;
                                            properties[actualCellIndex - 1].SetValue(obj, Convert.ChangeType(cellValue, type));
                                        }
                                    }
                                }
                                while (reader.ReadNextSibling());
                                //if its not the header row so append rowdata to the datatable
                                if (rcnt != 0)
                                {
                                    //dt.Rows.Add(tempRow);
                                    list.Add(obj);
                                }
                                rcnt++;
                                Console.WriteLine($"{rcnt} - {watch.Elapsed}");
                            }
                        }
                    }
                    reader.Close();
                    reader.Dispose();
                }
                document.Close();
                document.Dispose();
            }
            watch.Stop();
            Console.WriteLine($"Tempo levado: {watch.Elapsed}");
            return list;
        }

        public List<T> ListToExcelDOM<T>(string path)
        {
            var watch = Stopwatch.StartNew();
            var list = new List<T>();
            Type typeOfObject = typeof(T);
            var properties = typeOfObject.GetProperties();

            string value;
            DataTable dt = new DataTable();
            using (SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(path, true))
            {
                //Access the main Workbook part, which contains dataWorkbookPart 
                WorkbookPart workbookPart = myWorkbook.WorkbookPart;
                WorksheetPart worksheetPart = null;
                Sheet ss = workbookPart.Workbook.Descendants<Sheet>().SingleOrDefault();
                worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;

                if (worksheetPart != null)
                {
                    Row lastRow = worksheetPart.Worksheet.Descendants<Row>().LastOrDefault();

                    //Row firstRow = worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault();
                    //if (firstRow != null)
                    //{

                    //    foreach (Cell c in firstRow.ChildElements)
                    //    {
                    //        value = GetValue(c, stringTablePart);
                    //        dt.Columns.Add(value);
                    //    }
                    //}

                    if (lastRow != null)
                    {
                        for (int i = 2; i <= lastRow.RowIndex; i++)
                        {
                            T obj = (T)Activator.CreateInstance(typeOfObject);
                            //DataRow dr = dt.NewRow();
                            bool empty = true;
                            Row row = worksheetPart.Worksheet.Descendants<Row>().Where(r => i == r.RowIndex).FirstOrDefault();
                            int j = 0;
                            if (row != null)
                            {
                                foreach (Cell c in row.ChildElements)
                                {
                                    //Get cell value
                                    value = GetValue(c, stringTablePart);
                                    if (value != null && value != string.Empty && value != "")
                                    {
                                        empty = false;
                                    }
                                    var type = properties[j].PropertyType;
                                    properties[j].SetValue(obj, Convert.ChangeType(value, type));
                                    //dr[j] = value;
                                    j++;
                                    if (j == dt.Columns.Count)
                                    {
                                        break;
                                    }
                                }
                                if (empty)
                                {
                                    break;
                                }
                                //dt.Rows.Add(dr);
                                Console.WriteLine($"{i} - {watch.Elapsed}");
                                list.Add(obj);
                            }
                        }
                    }
                }
                myWorkbook.Close();
            }
            watch.Stop();
            Console.WriteLine($"Tempo levado: {watch.Elapsed}");
            return list;
        }

        public List<T> ImportExcel<T>(string excelFilePath)
        {
            var watch = Stopwatch.StartNew();
            List<T> list = new List<T>();
            Type typeOfObject = typeof(T);
            using (IXLWorkbook workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheets.First();
                var properties = typeOfObject.GetProperties();
                //header column texts
                var columns = worksheet.FirstRow().Cells().Select((v, i) => new { Value = v.Value, Index = i + 1 });//indexing in closedxml starts with 1 not from 0

                int rowCount = 0;
                foreach (IXLRow row in worksheet.RowsUsed().Skip(1))//Skip first row which is used for column header texts
                {
                    T obj = (T)Activator.CreateInstance(typeOfObject);

                    foreach (var prop in properties)
                    {
                        int colIndex = columns.SingleOrDefault(c => c.Value.ToString() == prop.Name.ToString()).Index;
                        var val = row.Cell(colIndex).Value;
                        var type = prop.PropertyType;
                        prop.SetValue(obj, Convert.ChangeType(val, type));
                    }

                    list.Add(obj);
                    rowCount++;
                    Console.WriteLine($"{rowCount} - {watch.Elapsed}");
                }

            }
            watch.Stop();
            Console.WriteLine($"Tempo levado: {watch.Elapsed}");
            return list;
        }

        public DataTable ExtractExcelSheetValuesToDataTable(string xlsxFilePath)
        {
            string value;
            DataTable dt = new DataTable();
            using (SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(xlsxFilePath, true))
            {
                //Access the main Workbook part, which contains dataWorkbookPart 
                WorkbookPart workbookPart = myWorkbook.WorkbookPart;

                WorksheetPart worksheetPart = null;
                Sheet ss = workbookPart.Workbook.Descendants<Sheet>().SingleOrDefault();
                worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;

                if (worksheetPart != null)
                {
                    Row lastRow = worksheetPart.Worksheet.Descendants<Row>().LastOrDefault();

                    Row firstRow = worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault();
                    if (firstRow != null)
                    {

                        foreach (Cell c in firstRow.ChildElements)
                        {
                            value = GetValue(c, stringTablePart);
                            dt.Columns.Add(value);
                        }
                    }

                    if (lastRow != null)
                    {
                        for (int i = 2; i <= lastRow.RowIndex; i++)
                        {
                            DataRow dr = dt.NewRow();
                            bool empty = true;
                            Row row = worksheetPart.Worksheet.Descendants<Row>().Where(r => i == r.RowIndex).FirstOrDefault();
                            int j = 0;
                            if (row != null)
                            {
                                foreach (Cell c in row.ChildElements)
                                {
                                    //Get cell value
                                    value = GetValue(c, stringTablePart);
                                    if (value != null && value != string.Empty && value != "")
                                    {
                                        empty = false;
                                    }
                                    dr[j] = value;
                                    j++;
                                    if (j == dt.Columns.Count)
                                    {
                                        break;
                                    }
                                }
                                if (empty)
                                {
                                    break;
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }
                }


                myWorkbook.Close();
            }

            return dt;
        }

        public DataTable ExtractExcelSAXToDataTable(string path)
        {
            //i want to import excel to data table
            var dt = new DataTable();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                //row counter
                int rcnt = 0;
                while (reader.Read())
                {
                    //find xml row element type 
                    //to understand the element type you can change your excel file eg : test.xlsx to test.zip
                    //and inside that you may observe the elements in xl/worksheets/sheet.xml
                    //that helps to understand openxml better
                    if (reader.ElementType == typeof(Row))
                    {
                        //create data table row type to be populated by cells of this row
                        DataRow tempRow = dt.NewRow();
                        //***** HANDLE THE SECOND SENARIO*****
                        //if row has attribute means it is not a empty row
                        if (reader.HasAttributes)
                        {
                            //read the child of row element which is cells
                            //here first element
                            reader.ReadFirstChild();
                            do
                            {
                                //find xml cell element type 
                                if (reader.ElementType == typeof(Cell))
                                {
                                    Cell c = (Cell)reader.LoadCurrentElement();

                                    string cellValue;
                                    int actualCellIndex = GetColumnIndex(c);

                                    if (c.DataType != null && c.DataType == CellValues.SharedString)
                                    {
                                        SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));
                                        cellValue = ssi.Text.Text;
                                    }
                                    else
                                    {
                                        cellValue = c.CellValue.InnerText;
                                    }

                                    //if row index is 0 its header so columns headers are added & also can do some headers check incase
                                    if (rcnt == 0)
                                    {
                                        dt.Columns.Add(cellValue);
                                    }
                                    else
                                    {
                                        // instead of tempRow[c.CellReference] = cellValue;
                                        tempRow[actualCellIndex -1] = cellValue;
                                    }
                                }
                            }
                            while (reader.ReadNextSibling());
                            //if its not the header row so append rowdata to the datatable
                            if (rcnt != 0)
                            {
                                dt.Rows.Add(tempRow);
                            }
                            rcnt++;
                        }
                    }
                }
            }
            return dt;
        }

        private static int GetColumnIndex(Cell cell)
        {
            string cellReference = cell.CellReference.ToString().ToUpper();
            if (string.IsNullOrEmpty(cellReference)) return 0;

            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber + 1;
        }

        public string GetValue(Cell cell, SharedStringTablePart stringTablePart)
        {
            if (cell.ChildElements.Count == 0)
            {
                return null;
            }

            //get cell value
            string value = cell.ElementAt(0).InnerText;//CellValue.InnerText;
            //Look up real value from shared string table
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            return value;
        }

        public List<T> BindList<T>(DataTable dt)
        {
            // Example 1:
            // Get private fields + non properties
            //var fields = typeof(T).GetFields(BindingFlags.NonPublic | BindingFlags.Instance);

            // Example 2: Your case
            // Get all public fields
            var fields = typeof(T).GetProperties();

            List<T> lst = new List<T>();

            foreach (DataRow dr in dt.Rows)
            {
                // Create the object of T
                var ob = Activator.CreateInstance<T>();

                foreach (var fieldInfo in fields)
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        // Matching the columns with fields
                        if (fieldInfo.Name == dc.ColumnName)
                        {
                            // Get the value from the datatable cell
                            object value = dr[dc.ColumnName];

                            // Set the value into the object
                            fieldInfo.SetValue(ob, value);
                            break;
                        }
                    }
                }

                lst.Add(ob);
            }

            return lst;
        }
    }
}
