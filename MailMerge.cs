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
                wordDoc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                selection = wordApp.Selection;

                //Add a section for each record.
                int ndx = 0;
                foreach (XmlNode row in data.SelectNodes("/root/row"))
                {
                    if (ndx > 0)
                    {
                        selection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast);
                        selection.InsertBreak(ref sectionBreak);
                    }
                    selection.InsertFile(template, ref missing, ref missing, ref missing, ref missing);
                    selection.Document.Select();
                    MergeIt(selection, row, wordApp);
                    ndx++;
                }
            }
            catch(Exception ex)
            {
                wordApp.Application.Quit();
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
                            //Check for child records
                            XmlNodeList children = row.SelectNodes("./" + col.Name);
                            if (children.Count > 1)
                            {
                                //Merge child records.
                                //Select and delete the child container merge field.
                                myMergeField.Select();
                                myMergeField.Delete();
                                //Select and copy the table row.
                                selection.SelectRow();
                                selection.Copy();

                                //Select the table.
                                Word.Table table = selection.Tables[1];

                                //Append rows to the table.
                                int max = children.Count - 1;
                                for (int cnt = 0; cnt < max; cnt++)
                                {
                                    selection.Paste();
                                }
                                //Select each row in the table and perform a merge.
                                int ndx = 0;
                                foreach (Word.Row tr in table.Rows)
                                {
                                    if (tr.Range.Fields.Count > 0)
                                    {
                                        XmlNode child = children[ndx];
                                        tr.Select();
                                        MergeIt(selection, child, wordApp);
                                        ndx++;
                                    }
                                }
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
