using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace DocxProcessor
{
    public class GetTextWordTemplate
    {
        #region
        public List<string> getAllText(string TemplateFilePath)
        {
            List<string> result = new List<string>();

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(TemplateFilePath, false))
            {

                var paragraphs = wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();                

                foreach (var paragraph in paragraphs)
                {
                    result.Add(paragraph.InnerText);
                }                
            }
            return result;
        }
        #endregion
    }
}
