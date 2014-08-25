using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Xml;

namespace TetherApp
{
    class MailMerge
    {
        public static void Execute(string template, XmlDocument data, string output)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = new Word.Document();
            object missing = System.Reflection.Missing.Value;
            object sectionBreak = Word.WdBreakType.wdSectionBreakNextPage;
            object oTemplate = template;
            Word.Selection selection = null;

            try
            {
                wordDoc = wordApp.Documents.Add(template, ref missing, ref missing, ref missing);

                selection = wordApp.Selection;

                //Add a section for each record.
                int ndx = 0;
                foreach (XmlNode row in data.SelectNodes("/root/row"))
                {
                    if (ndx > 0)
                    {
                        selection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast);
                        selection.InsertBreak(ref sectionBreak);
                        selection.InsertFile(template, ref missing, ref missing, ref missing, ref missing);
                    }
                    selection.Document.Select();
                    MergeIt(selection, row, wordApp);
                    ndx++;
                }
            }
            catch(Exception ex)
            {
                // Application
                object saveOptionsObject = Word.WdSaveOptions.wdDoNotSaveChanges;
                wordApp.Application.Quit(ref saveOptionsObject, ref missing, ref missing); 
                throw ex;
            }

            selection.GoTo(Word.WdGoToItem.wdGoToLine, 1);
            wordDoc.SaveAs(output);
            wordApp.Documents.Open(output);
            
        }
        public static void MergeIt(Word.Selection selection, XmlNode row, Word.Application wordApp)
        {
            //Replace merge fields.
            Word.Range range = selection.Range;

            foreach (Word.Field myMergeField in range.Fields)
            {
                Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();

                    //Check the data for fieldname matches.
                    //Look for multiple replacements.
                    foreach (XmlNode col in row.ChildNodes)
                    {
                        if (fieldName.ToLower() == col.Name.ToLower())
                        {
                            //Check for secondary records
                            if (col.FirstChild.HasChildNodes)
                            {
                                XmlNodeList secondaryRows = row.SelectNodes("./" + col.Name);
                                //Select and delete the child container merge field.
                                myMergeField.Select();
                                myMergeField.Delete();
                                //Select and copy the table row.
                                selection.SelectRow();
                                selection.Copy();
                                try
                                {
                                    //Select the table.
                                    Word.Table table = selection.Tables[1];
                                    //Paste a row for each record in the secondary source.
                                    int ndx = 0;
                                    foreach (XmlNode secRow in secondaryRows)
                                    {
                                        if (ndx > 0)
                                        {
                                            selection.Paste();
                                        }
                                        ndx++;
                                    }
                                    //Select each row in the table and perform a merge.
                                    ndx = 0;
                                    foreach (Word.Row tr in table.Rows)
                                    {
                                        if (tr.Range.Fields.Count > 0)
                                        {
                                            if (ndx < secondaryRows.Count)
                                            {
                                                XmlNode child = secondaryRows[ndx];
                                                tr.Select();
                                                MergeIt(selection, child, wordApp);
                                            }
                                            ndx++;
                                        }
                                    }
                                }
                                catch{ }
                            }
                            else
                            {
                                //Standard merge field.
                                myMergeField.Select();
                                wordApp.Selection.TypeText(col.InnerText);
                            }
                            break;
                        }
                    }
                }
            }
        }
    }
}
