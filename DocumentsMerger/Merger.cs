using Independentsoft.Office.Odf;
using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentsMerger
{
    public static class Merger
    {
        public static void MergeTxt(Docs docs, string resultfilepath)
        {
            List<string> inputfilespaths = docs.GetList();
            StreamReader sr;
            StreamWriter sw;
            string text = String.Empty;

            for (int i = 0; i < inputfilespaths.Count; i++)
            {
                sr = File.OpenText(inputfilespaths[i]);
                text = String.Empty;
                text = sr.ReadToEnd();
                sw = File.AppendText(resultfilepath);
                sw.WriteLine(text);
                sw.Close();
                Console.WriteLine(text);
            }
        }

        public static void MergeOdt(Docs docs, string outputfilepath)
        {
            List<string> filepaths = new List<string>(docs.GetList());
            TextDocument temp = new TextDocument(outputfilepath);
            List<TextDocument> textdocuments = ListToTextDocuments(filepaths);

            for (int i = 0; i < textdocuments.Count; i++)
            {
                IList<ITextContent> textContents = textdocuments[i].Body.Content;

                foreach (var content in textContents)
                {
                    temp.Body.Add(content);
                }
            }

            temp.Save(outputfilepath, true);
        }

        private static List<TextDocument> ListToTextDocuments(List<string> inputfilepaths)
        {
            List<TextDocument> textdocuments = new List<TextDocument>();

            foreach (var path in inputfilepaths)
            {
                textdocuments.Add(new TextDocument(path));
            }

            return textdocuments;
        }

        public static void Merge(string[] filesToMerge, string outputFilename, bool insertPageBreaks, string documentTemplate)
        {
            object defaultTemplate = documentTemplate;
            object missing = System.Type.Missing;
            object pageBreak = Word.WdBreakType.wdPageBreak;
            object outputFile = outputFilename;

            Word._Application wordApplication = new Word.Application();

            try
            {
                Word._Document wordDocument = wordApplication.Documents.Add(ref defaultTemplate, ref missing, ref missing, ref missing);
                Word.Selection selection = wordApplication.Selection;

                foreach (string file in filesToMerge)
                {
                    //selection.InsertFile(file, ref missing, ref missing, ref missing, ref missing);
                    selection.InsertFile(file);

                    if (insertPageBreaks)
                    {
                        selection.InsertBreak(ref pageBreak);
                    }
                }

                wordDocument.SaveAs(ref outputFile, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, 
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                wordDocument = null;
            }
            catch (Exception ex)
            {
                //TODO
                throw ex;
            }
            finally
            {
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }
        }
    }
}
