using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OpenXmlPowerTools;
using System.Xml.Linq;
using PuppeteerReportCsharp;

namespace DocxProcessor
{
    public class ConvertWordTemplate
    {
        public byte[] WordToHtml(string FilePath)
        {
            // Filepath to Byte Array
            MemoryStream stream = new MemoryStream();

            using (FileStream fs = new FileStream(FilePath, FileMode.Open))
            {
                fs.CopyTo(stream);
            }

            stream.Seek(0, SeekOrigin.Begin);              

            return WordToHtml(stream.ToArray());
        }

        public byte[] WordToHtml(byte[] bytes)
        {
            MemoryStream origin = new MemoryStream(bytes, true);
            MemoryStream destination = new MemoryStream();

            origin.CopyToAsync(destination);
            origin.Close();
            
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(destination, true))
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                    MemoryStream result = new MemoryStream();
                    html.Save(result);
                    result.Position = 0;
                    return result.ToArray();
                }
            }
        }
        public byte[] HtmlToPdf(byte[] bytes)
        {
            MemoryStream origin = new MemoryStream(bytes, true);
            MemoryStream destination = new MemoryStream();

            origin.CopyToAsync(destination);
            origin.Close();

            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(destination, true))
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        FabricateCssClasses = true,
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                    MemoryStream result = new MemoryStream();
                    html.Save(result);
                    result.Position = 0;
                    return result.ToArray();
                }
            }
        }
        public async System.Threading.Tasks.Task<byte[]> HtmlToPdf(string Filepath)
        {
            var puppeteer = new PuppeteerReport();           
            return await puppeteer.PDF(Filepath,
                new PuppeteerSharp.PdfOptions
                {
                    Format = PuppeteerSharp.Media.PaperFormat.A4,
                    PreferCSSPageSize = true,
                    MarginOptions = new PuppeteerSharp.Media.MarginOptions
                    {
                        Top = "10mm",
                        Left = "10mm",
                        Right = "10mm",
                        Bottom = "10mm"
                    }
                }) ;
        }
        /*
           public File ConvertWordToPDF(string FilePath)
            {

            }
        */
    }
}
