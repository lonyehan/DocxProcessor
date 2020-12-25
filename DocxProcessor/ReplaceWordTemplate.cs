using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System;
using System.Text.RegularExpressions;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Web;
using System.Drawing;
using DocumentFormat.OpenXml;

namespace DocxProcessor
{


    public class ReplaceWordTemplate
    {
        #region ReplaceStringToString: Replace String At WordTemplate by String
        /// <summary>
        /// 用字串取代字串
        /// </summary>
        /// <param name="doc">WmlDocument 實質上Word的內容</param>
        /// <param name="SearchString">查詢用的字串</param>
        /// <param name="ReplaceString">取代用的字串</param>
        /// <returns>WmlDocument 取代後的字串</returns>
        private WmlDocument ReplaceStringToString(ref WmlDocument doc, string SearchString, string ReplaceString)
        {
            if (string.IsNullOrEmpty(ReplaceString)) ReplaceString = "\r";
            
            return doc.SearchAndReplace(SearchString, ReplaceString, true);
        }
        #endregion

        public FileStream ReplaceByImage(string TemplateFilePath, Dictionary<string, string> ReplaceItems)
        {
            var wordDoc = WordprocessingDocument.Open(TemplateFilePath, true);
            
            var body = wordDoc.MainDocumentPart.Document.Body;
                        
            var paragraphs = body.Elements<Paragraph>();
                        
            foreach (KeyValuePair<string, string> pair in ReplaceItems)
            {
                string SearchString = pair.Key;

                var ReplaceImage = new ImageData(pair.Value)
                {

                    Width = 1,

                    Height = 1

                };                              

                foreach (Paragraph p in paragraphs)
                {                                                                      
                    if (p.InnerText.Contains(SearchString))
                    {
                        p.RemoveAllChildren<Run>();                        

                        p.AppendChild(GenerateImageRun(wordDoc, ReplaceImage));
                    }                                           
                }
            }
            // 執行後儲存
            return (FileStream)wordDoc.MainDocumentPart.GetStream();
           
            /*
            #region 輸出檔案Byte[]

            var stream = wordDoc.MainDocumentPart.GetStream(FileMode.Open);            

            byte[] bytes = new byte[stream.Length];

            stream.Read(bytes, 0, bytes.Length);
            
            stream.Seek(0, SeekOrigin.Begin);

            #endregion

            return bytes;            
            */
        }
        #region Replace: WordTemplate Replace Function ([Core] Replace To Byte[] By Dictionary)
        /// <summary>
        /// 取代WordTemplate的內容字串
        /// </summary>
        /// <param name="TemplateFilePath">模板來源路徑</param>
        /// <param name="ReplaceItems">
        ///                             用來取代的字典樹{key: string, value: string}
        ///                            key: SearchString
        ///                            value: ReplaceString
        /// </param>                                 
        /// <returns>byte[]</returns>        
        public byte[] Replace(string TemplateFilePath, Dictionary<string, string> ReplaceItems)
        {
            if (File.Exists(TemplateFilePath) == true)
            {
                var stream = new MemoryStream();

                #region 字典取代部分
                WmlDocument doc = new WmlDocument(TemplateFilePath);
                
                foreach (KeyValuePair<string, string> keyValuePair in ReplaceItems)
                {
                    
                    string SearchString = keyValuePair.Key;
                    string ReplaceString = keyValuePair.Value.Replace("\r\n", "\n").Replace("\n", "\r\n").Replace("\r\n", "</w:t><w:br/><w:t>");  // 解決換行問題     

                    #region 字串替代                    
                    doc = ReplaceStringToString(ref doc, SearchString, ReplaceString);
                    #endregion
                }

                stream.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
                #endregion

                #region 取代後字串格式整理
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    string docText = wordDoc.MainDocumentPart.GetXDocument().ToString();
                    
                    docText = docText.Replace("\n", "").Replace("\r\n", ""); // 去除未替換的換行字串

                    docText = docText.Replace("\t", "  "); // 將tab字串 換成真正的tab

                    XDocument mainDocumentXDoc =  XDocument.Parse(HttpUtility.HtmlDecode(docText.Replace("\n", "").Replace("\r\n", "")));
                    
                    mainDocumentXDoc.Save(wordDoc.MainDocumentPart.GetStream(FileMode.Create));                    
                    
                }
                #endregion

                stream.Position = 0;

                return stream.ToArray();
            }
            else
            {
                throw new FileNotFoundException("Template File is not exist!");
            }            
        }
        #endregion

        #region Replace: WordTemplate Replace Function (Replace To File By Dictionay)
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
        /// <returns>byte[]</returns>        
        public bool Replace(string TemplateFilePath, string OutputFilePath, Dictionary<string, string> ReplaceItems)
        {
            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Replace(TemplateFilePath, ReplaceItems));

            bw.Close();
            fs.Close();

            return true;
        }
        #endregion

        #region Replace: WordTemplate Replace Function (Replace To Byte[] By Model)
        /// <summary>
        /// 取代WordTemplate的內容字串
        /// </summary>
        /// <param name="TemplateFilePath">模板來源路徑</param>
        /// <param name="ReplaceModel">
        ///                             用來取代的Model        
        /// </param>                                 
        /// <returns>byte[]</returns>        
        public byte[] Replace<T>(string TemplateFilePath, T ReplaceModel) where T : class
        {
            PropertyInfo[] infos = ReplaceModel.GetType().GetProperties();

            Dictionary<string, string> ReplaceItems = new Dictionary<string, string>();

            foreach (PropertyInfo info in infos)
            {
                string Key = "#" + info.Name + "#";

                string Value = info.GetValue(ReplaceModel, null) == null ? "" : info.GetValue(ReplaceModel, null).ToString();

                ReplaceItems.Add(Key, Value);
            }

            return Replace(TemplateFilePath, ReplaceItems);
        }
        #endregion



        public class ImageData

        {

            public string FileName = string.Empty;

            public byte[] BinaryData;

            public Stream DataStream => new MemoryStream(BinaryData);

            public ImagePartType ImageType

            {

                get

                {

                    var ext = Path.GetExtension(FileName).TrimStart('.').ToLower();

                    switch (ext)

                    {

                        case "jpg":

                            return ImagePartType.Jpeg;

                        case "png":

                            return ImagePartType.Png;

                        case "gif":

                            return ImagePartType.Gif;

                        case "bmp":

                            return ImagePartType.Bmp;

                    }

                    throw new ApplicationException($"Unsupported image type: {ext}");

                }

            }

            public int SourceWidth;

            public int SourceHeight;

            public decimal Width;

            public decimal Height;

            public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);

            public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);

            private const decimal INCH_TO_CM = 2.54M;

            private const decimal CM_TO_EMU = 360000M;

            public string ImageName;

            public ImageData(string fileName, byte[] data, int dpi = 300)

            {

                FileName = fileName;

                BinaryData = data;

                Bitmap img = new Bitmap(new MemoryStream(data));

                SourceWidth = img.Width;

                SourceHeight = img.Height;

                Width = ((decimal)SourceWidth) / dpi * INCH_TO_CM;

                Height = ((decimal)SourceHeight) / dpi * INCH_TO_CM;

                ImageName = $"IMG_{Guid.NewGuid().ToString().Substring(0, 8)}";

            }

            public ImageData(string fileName, int dpi = 300) :

                this(fileName, File.ReadAllBytes(fileName), dpi)

            {

            }

        }
       
        private static Run GenerateImageRun(WordprocessingDocument wordDoc, ImageData img)

            {

                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(img.ImageType);

                var relationshipId = mainPart.GetIdOfPart(imagePart);

                imagePart.FeedData(img.DataStream);

                // Define the reference of the image.

                var element =

                     new Drawing(

                         new DW.Inline(

                             //Size of image, unit = EMU(English Metric Unit)

                             //1 cm = 360000 EMUs

                             new DW.Extent() { Cx = img.WidthInEMU, Cy = img.HeightInEMU },

                             new DW.EffectExtent()

                             {

                                 LeftEdge = 0L,

                                 TopEdge = 0L,

                                 RightEdge = 0L,

                                 BottomEdge = 0L

                             },

                             new DW.DocProperties()

                             {

                                 Id = (UInt32Value)1U,

                                 Name = img.ImageName

                             },

                             new DW.NonVisualGraphicFrameDrawingProperties(

                                 new A.GraphicFrameLocks() { NoChangeAspect = true }),

                             new A.Graphic(

                                 new A.GraphicData(

                                     new PIC.Picture(

                                         new PIC.NonVisualPictureProperties(

                                             new PIC.NonVisualDrawingProperties()

                                             {

                                                 Id = (UInt32Value)0U,

                                                 Name = img.FileName

                                             },

                                             new PIC.NonVisualPictureDrawingProperties()),

                                         new PIC.BlipFill(

                                             new A.Blip(

                                                 new A.BlipExtensionList(

                                                     new A.BlipExtension()

                                                     {

                                                         Uri =

                                                            "{28A0092B-C50C-407E-A947-70E740481C1C}"

                                                     })

                                             )

                                             {

                                                 Embed = relationshipId,

                                                 CompressionState =

                                                 A.BlipCompressionValues.Print

                                             },

                                             new A.Stretch(

                                                 new A.FillRectangle())),

                                         new PIC.ShapeProperties(

                                             new A.Transform2D(

                                                 new A.Offset() { X = 0L, Y = 0L },

                                                 new A.Extents()
                                                 {

                                                     Cx = img.WidthInEMU,
                                                     Cy = img.HeightInEMU
                                                 }),

                                             new A.PresetGeometry(

                                                 new A.AdjustValueList()

                                             )

                                             { Preset = A.ShapeTypeValues.Rectangle }))

                                 )

                                 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })

                         )

                         {

                             DistanceFromTop = (UInt32Value)0U,

                             DistanceFromBottom = (UInt32Value)0U,

                             DistanceFromLeft = (UInt32Value)0U,

                             DistanceFromRight = (UInt32Value)0U,

                             EditId = "50D07946"

                         });

                return new Run(element);

            }

        }


    }
