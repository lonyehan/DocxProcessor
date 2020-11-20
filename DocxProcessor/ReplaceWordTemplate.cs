using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxProcessor
{
    public class ReplaceWordTemplate
    {
        #region Replace: WordTemplate Replace Function
        /// <summary>
        /// 取代WordTemplate的內容字串
        /// </summary>
        /// <param name="TemplateFilePath">模板來源路徑</param>
        /// <param name="OutputFilePath">目標路徑</param>
        /// <param name="ReplaceItems">
        ///                             用來取代的字典樹{key: string, value: string}
        ///                            key: SearchString
        ///                            value: ReplaceString
        /// </param>                                 
        /// <returns>bool</returns>        
        public bool Replace(string TemplateFilePath, string OutputFilePath, Dictionary<string, string> ReplaceItems)
        {
            try
            {
                // Template needs exist.
                if (File.Exists(TemplateFilePath) == true)
                {
                    File.Copy(TemplateFilePath, OutputFilePath, true);

                    #region Read docx's content and replace them by ReplaceItems                    
                    using(WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(OutputFilePath, true))
                    {
                        
                        string docText = null;

                        using (StreamReader sr = new StreamReader(wordprocessingDocument.MainDocumentPart.GetStream()))
                        {
                            docText = sr.ReadToEnd();                            
                        }

                        foreach(KeyValuePair<string, string> keyValuePair in ReplaceItems)
                        {
                            string SearchString = keyValuePair.Key;
                            string ReplaceString = keyValuePair.Value.Replace("\r\n", "<w:br/>").Replace("\n", "<w:br/>");                            
                            Regex regexText = new Regex(SearchString);
                            docText = regexText.Replace(docText, ReplaceString);
                        }

                        using (StreamWriter sw = new StreamWriter(wordprocessingDocument.MainDocumentPart.GetStream(FileMode.Create)))
                        {
                            sw.Write(docText);                            
                        } 
                    }
                    #endregion
                }
                else
                {
                    throw new FileNotFoundException("Template File is not exist!");                    
                }

                return true;
            }
            catch (InvalidDataException e)
            {
                throw e;                
            }            
        }
        #endregion
      
    }
}
