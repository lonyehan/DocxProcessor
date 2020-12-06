using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OpenXmlPowerTools;
using System.Xml.Linq;
using System.Text;
//using DinkToPdf;
namespace DocxProcessor
{
    public class ConvertWordTemplate
    {
        /*
        public void ConvertWordToPDF(string InFilePath, string OutFilePath)
        {
            MemoryStream memoryStream = new MemoryStream();
            //byte[] byteArray = File.ReadAllBytes(InFilePath);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(InFilePath, true))
            {
                HtmlConverterSettings settings = new HtmlConverterSettings()
                {
                    PageTitle = "My Page Title"
                };
                XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                //File.WriteAllText("C:\\Users\\JasonJian\\Desktop\\sideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\a.html", html.ToStringNewLineOnAttributes());
                File.WriteAllText(OutFilePath, html.ToStringNewLineOnAttributes());
                doc.Close();
            }

        }
        */

        public string ConvertWordToHTML(string InFilePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(InFilePath, true))
            {
                HtmlConverterSettings settings = new HtmlConverterSettings()
                {
                    PageTitle = "My Page Title"
                };
                XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                doc.Close();

                string htmlString = html.ToStringNewLineOnAttributes();
                byte[] bytes = Encoding.UTF8.GetBytes(htmlString);
                //return bytes;
                return htmlString;
            }
        }
        /*
        public void ConvertHTMLToPDF(string pdfString, string OutFilePath)
        {
            var converter = new SynchronizedConverter(new PdfTools());
            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Portrait,
                    PaperSize = PaperKind.A4,
                    Margins = new MarginSettings() { Top = 10 },
                    Out = OutFilePath,
                },
                Objects = {
                    new ObjectSettings()
                    {
                        Page = "http://google.com/",
                        //HtmlContent = pdfString
                    },
                }
            };
            byte[] pdf = converter.Convert(doc);
        }
        */
    }
}