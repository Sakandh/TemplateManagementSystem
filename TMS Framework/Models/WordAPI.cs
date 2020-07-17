using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Management.Automation;
using TMS_Framework.Models;

namespace TMS_Framework.Models
{
    public class WordAPI
    {
        // constructor of the Application to open Word
        Application application;

        //Closes all word running instances
        public void KillWord()
        {
            PowerShell.Create().AddCommand("Stop-Process").AddParameter("Name", "Winword").Invoke();
        }

        //find list of keyword and table names
        public List<string> getExistingKeywords(UnitOwner unitOwner)
        {
            //start new word instace
            application = new Application();
            Document document;

            // opens word and goes the file location
            if (File.Exists(unitOwner.baseTemplateFileName))
            {
                document = application.Documents.Open(unitOwner.baseTemplateFileName);
            }
            else
            {
                application.Quit();
                throw new FileNotFoundException("File does not exist: " + unitOwner.baseTemplateFileName);
            }

            List<string> foundKeyWord = new List<string>();

            // loops through the document to find all the keyword words, 
            // comparing the words from the full json list
            foreach (ClientData clientData in unitOwner.clientDataList)
            {
                object findText = clientData.Key;
                object missing = Missing.Value;

                application.Selection.Find.ClearFormatting();
                application.Selection.Find.Wrap = WdFindWrap.wdFindContinue;

                if (application.Selection.Find.Execute(ref findText,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing))
                {
                    foundKeyWord.Add(findText.ToString());
                }
            }

            document.Close();

            application.Quit();

            return foundKeyWord;
        }

        public void ClientDataFindAndReplace(UnitOwner unitOwner)
        {
            // first, check to see if the exists
            if (!File.Exists(unitOwner.baseTemplateFileName))
            {
                throw new FileNotFoundException("File does not exist: " + unitOwner.baseTemplateFileName);
            }

            try
            {
                //start new word instace
                application = new Application();
                Document document;

                //read the document
                document = application.Documents.Open(unitOwner.baseTemplateFileName);

                //save this as target document right away to avoid template corruption
                document.SaveAs2(unitOwner.mergedFileName);
                document.Close();

                //now, read this document again to manipulate
                document = application.Documents.Open(unitOwner.mergedFileName);

                // finds the block keywords, sends the corresponding filepath for the block documents
                foreach (BlockData blockData in unitOwner.blockDataList)
                {
                    Range range = document.Range(0, 1);
                    for (int i = 1; i <= document.Paragraphs.Count; i++)
                    {
                        string paraContent = document.Paragraphs[i].Range.Text;
                        if (paraContent.Length >= blockData.Key.Length && paraContent.Substring(0, blockData.Key.Length) == blockData.Key)
                        {
                            range = document.Paragraphs[i].Range;
                            break;
                        }
                    }

                    range.InsertFile(blockData.Value);
                }

                // loops through all the keywords that need to be replaced
                foreach (ClientData clientData in unitOwner.clientDataList)
                {
                    // finds the text to be replaced and what text will be there
                    Find findObject = application.Selection.Find;
                    findObject.Text = clientData.Key;

                    if (string.IsNullOrWhiteSpace(clientData.Value))
                    {
                        findObject.Replacement.Text = "MISSING";
                    }
                    else
                    {
                        findObject.Replacement.Text = clientData.Value;
                    }

                    object missing = Missing.Value;
                    object replaceAll = WdReplace.wdReplaceAll;

                    findObject.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }

                // do the table replacement
                foreach (TableData tableData in unitOwner.tableDataList)
                {
                    CreateTable(tableData, document);
                }

                //create the barcode picture
                BarCode barCode = new BarCode();
                string barCodeImageFileName = Path.GetTempPath() + unitOwner.barCode + ".png";
                barCode.WriteBarCode(unitOwner.barCode, barCodeImageFileName);

                InlineShape bcImage = application.Selection.InlineShapes.AddPicture(barCodeImageFileName, Type.Missing, Type.Missing, Type.Missing);

                //converts the image to shape then rotates and positions it
                Shape bcShape = bcImage.ConvertToShape();
                bcShape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                bcShape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                bcShape.Top = 3.5F * 72;
                bcShape.Left = 6.5F * 72;
                bcShape.Rotation = 90;

                document.SaveAs2(unitOwner.mergedFileName);

                document.Close();
                application.Quit();
            }

            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (application != null)
                {

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
                    GC.Collect();
                }
            }
        }

        private void CreateTable(TableData tableData, Document document)
        {
            //find the range first
            Range range = document.Range(0, 1);
            for (int i = 1; i <= document.Paragraphs.Count; i++)
            {
                string paraContent = document.Paragraphs[i].Range.Text;
                if (paraContent.Length >= tableData.Key.Length && paraContent.Substring(0, tableData.Key.Length) == tableData.Key)
                {
                    range = document.Paragraphs[i].Range;
                    break;
                }
            }

            //create the table in the range
            Table table = range.Tables.Add(range, tableData.rowCount, tableData.columnCount);

            //these 2 lines put borders both inside & outside - see result image at end
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            table.Range.Font.Name = "Calibri";
            if (tableData.columnCount > 6)
            {
                table.Range.Font.Size = 7.0F;
            }
            else
            {
                table.Range.Font.Size = 10.0F;
            }

            table.Range.Font.Bold = 0;

            //populate the table now
            char rowDivider = '|';
            char fieldDivider = ';';
            List<string> dataSet = tableData.Value.Trim().Split(rowDivider).ToList();   // Splits the data into rows
            dataSet.RemoveAt(dataSet.Count - 1); // Removes the last row, that is empty

            int currentRow = 0;
            int currentField = 0;
            foreach (string record in dataSet)
            {
                currentRow = currentRow + 1;                    // Increments the number of rows
                currentField = 0;
                string[] fields = record.Split(fieldDivider);   // Splits each individual data from each row
                foreach (string field in fields)
                {
                    currentField = currentField + 1;            // Increments the nunmber of columns
                    table.Cell(currentRow, currentField).Range.Text = field.Trim();     // Inserts the data into each cell in the table
                    if (currentRow == 1)
                    {
                        table.Cell(currentRow, currentField).Range.Font.Bold = 1;
                    }
                    if (currentField == tableData.columnCount)
                    {
                        break;
                    }
                }
                if (currentRow == tableData.rowCount)
                {
                    break;
                }
            }
        }
    }
}