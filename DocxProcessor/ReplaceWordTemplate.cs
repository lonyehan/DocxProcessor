﻿using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Web;
using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml;

namespace DocxProcessor
{
    #region ImageData: Get Image To Replace
    public class ImageData
    {
        public string FilePath { get; set; }
        public decimal WidthInEMU { get; set; }
        public decimal HeightInEMU { get; set; }

        private decimal CM_TO_EMU = 360000M;
        private decimal PIXEL_TO_CM = 0.0264583333M;

        public ImageData(string FilePath)
        {
            this.FilePath = FilePath;
            Image img;
            using (var bmpTemp = new Bitmap(FilePath))
            {
                img = new Bitmap(bmpTemp);
            }
            this.WidthInEMU = img.Width * PIXEL_TO_CM * CM_TO_EMU;
            this.HeightInEMU = img.Height * PIXEL_TO_CM * CM_TO_EMU;
        }
        public ImageData(string FilePath, decimal Width)
        {
            this.FilePath = FilePath;
            Image img;
            using (var bmpTemp = new Bitmap(FilePath))
            {
                img = new Bitmap(bmpTemp);
            }
            this.WidthInEMU = Width * CM_TO_EMU;
            this.HeightInEMU = img.Height * PIXEL_TO_CM * (Width / (img.Width * PIXEL_TO_CM)) * CM_TO_EMU;
        }

        public ImageData(string FilePath, decimal Width = 1.0M, decimal Height = 1.0M)
        {
            this.FilePath = FilePath;
            this.WidthInEMU = Width * CM_TO_EMU;
            this.HeightInEMU = Height * CM_TO_EMU;
        }
    };
    #endregion
    public class ReplaceWordTemplate
    {
        #region Add: Add Image To Table Cell
        private static void AddImageToCell(TableCell cell, string relationshipId, decimal Cx = 1, decimal Cy = 1)
        {
            var element =
              new Drawing(
                new DW.Inline(
                  new DW.Extent() { Cx = Convert.ToInt64(Cx), Cy = Convert.ToInt64(Cy) },
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
                      Name = "Picture 1"
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
                              Name = "New Bitmap Image.jpg"
                          },
                          new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                          new A.Blip(
                            new A.BlipExtensionList(
                              new A.BlipExtension()
                              {
                                  Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
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
                              new A.Extents() { Cx = Convert.ToInt64(Cx), Cy = Convert.ToInt64(Cy) }),
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
                    DistanceFromRight = (UInt32Value)0U
                });

            cell.Append(new Paragraph(new Run(element)));
        }
        #endregion

        #region Get: Get Image Type
        /// <summary>
        /// 獲得Input Image的Type
        /// </summary>
        /// <param name="TargetPath"></param>
        /// <returns></returns>
        private ImagePartType GetImageType(string TargetPath)
        {
            var ext = Path.GetExtension(TargetPath).TrimStart('.').ToLower();
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
        #endregion

        #region Converter: Byte[] To MemoryStream
        private MemoryStream ByteArrayToMemoryStream(byte[] bytes)
        {
            MemoryStream origin = new MemoryStream(bytes, true);
            MemoryStream destination = new MemoryStream();

            origin.CopyToAsync(destination);
            origin.Close();

            return destination;
        }
        #endregion 

        #region Converter: Filepath To Byte[]
        private byte[] FilePathToByteArray(string FilePath)
        {
            // Filepath to Byte Array
            MemoryStream stream = new MemoryStream();

            using (FileStream fs = new FileStream(FilePath, FileMode.Open))
            {
                fs.CopyTo(stream);
            }

            stream.Seek(0, SeekOrigin.Begin);
            return stream.ToArray();
        }
        #endregion

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

        #region Replace: Replace Table Cell By Image (Byte[] to Byte[])
        public byte[] ReplaceTableCellByImage(byte[] Source, Dictionary<string, ImageData> ReplaceItems)
        {
            MemoryStream originFile = new MemoryStream(Source, true);
            MemoryStream destination = new MemoryStream();

            originFile.CopyTo(destination);
            originFile.Close();
            using (var document = WordprocessingDocument.Open(destination, isEditable: true))
            {
                var mainPart = document.MainDocumentPart;
                foreach (var table in mainPart.Document.Body.Descendants<Table>())
                {
                    foreach (var keyValuePair in ReplaceItems)
                    {

                        string SearchString = keyValuePair.Key;

                        foreach (var pictureCell in table.Descendants<TableCell>())
                        {
                            if (pictureCell.InnerText.Contains(SearchString))
                            {
                                ImageData ReplacedImage = keyValuePair.Value;

                                ImagePart imagePart = mainPart.AddImagePart(GetImageType(ReplacedImage.FilePath));

                                using (FileStream stream = new FileStream(ReplacedImage.FilePath, FileMode.Open))
                                {
                                    imagePart.FeedData(stream);
                                }
                                pictureCell.RemoveAllChildren<Paragraph>();
                                AddImageToCell(pictureCell, mainPart.GetIdOfPart(imagePart), ReplacedImage.WidthInEMU, ReplacedImage.HeightInEMU);
                            }
                        }
                    }
                }
            }

            destination.Position = 0;

            return destination.ToArray();
        }
        #endregion

        #region Replace: Replace Table Cell By Image(Filepath to Byte[])
        public byte[] ReplaceTableCellByImage(string TemplateFilePath, Dictionary<string, ImageData> ReplaceItems)
        {

            FileStream originFile = new FileStream(TemplateFilePath, FileMode.Open);
            MemoryStream destination = new MemoryStream();

            originFile.CopyTo(destination);
            originFile.Close();

            using (var document = WordprocessingDocument.Open(destination, isEditable: true))
            {
                var mainPart = document.MainDocumentPart;

                foreach (var table in mainPart.Document.Body.Descendants<Table>())
                {
                    foreach (var keyValuePair in ReplaceItems)
                    {

                        string SearchString = keyValuePair.Key;

                        foreach (var pictureCell in table.Descendants<TableCell>())
                        {
                            if (pictureCell.InnerText.Contains(SearchString))
                            {
                                ImageData ReplacedImage = keyValuePair.Value;

                                ImagePart imagePart = mainPart.AddImagePart(GetImageType(ReplacedImage.FilePath));

                                using (FileStream stream = new FileStream(ReplacedImage.FilePath, FileMode.Open))
                                {
                                    imagePart.FeedData(stream);
                                }
                                pictureCell.RemoveAllChildren<Paragraph>();
                                AddImageToCell(pictureCell, mainPart.GetIdOfPart(imagePart), ReplacedImage.WidthInEMU, ReplacedImage.HeightInEMU);
                            }
                        }
                    }
                }
            }

            destination.Position = 0;

            return destination.ToArray();
        }
        #endregion        

        #region Replace: WordTemplate Replace Function (Replace Byte[] To Byte[] By Dictionary)
        public byte[] Replace(byte[] Source, Dictionary<string, string> ReplaceItems)
        {
            var stream = new MemoryStream();
            #region 字典取代部分
            WmlDocument doc = new WmlDocument("TemplateFile", Source);

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

                XDocument mainDocumentXDoc = XDocument.Parse(HttpUtility.HtmlDecode(docText.Replace("\n", "").Replace("\r\n", "")));

                mainDocumentXDoc.Save(wordDoc.MainDocumentPart.GetStream(FileMode.Create));

            }
            #endregion

            stream.Position = 0;

            return stream.ToArray();
        }
        #endregion

        #region Replace: WordTemplate Replace Function ([Core] Replace FilePath To Byte[] By Dictionary)
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

                    XDocument mainDocumentXDoc = XDocument.Parse(HttpUtility.HtmlDecode(docText.Replace("\n", "").Replace("\r\n", "")));

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

        #region Replace: WordTemplate Replace Function (Replace FilePath To Byte[] By Model)
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

        #region Replace: WordTemplate Replace Function (Replace Byte[] To Byte[] By Model)
        /// <summary>
        /// 取代WordTemplate的內容字串
        /// </summary>
        /// <param name="TemplateFilePath">模板來源路徑</param>
        /// <param name="ReplaceModel">
        ///                             用來取代的Model        
        /// </param>                                 
        /// <returns>byte[]</returns>        
        public byte[] Replace<T>(byte[] bytes, T ReplaceModel) where T : class
        {
            PropertyInfo[] infos = ReplaceModel.GetType().GetProperties();

            Dictionary<string, string> ReplaceItems = new Dictionary<string, string>();

            foreach (PropertyInfo info in infos)
            {
                string Key = "#" + info.Name + "#";

                string Value = info.GetValue(ReplaceModel, null) == null ? "" : info.GetValue(ReplaceModel, null).ToString();

                ReplaceItems.Add(Key, Value);
            }

            return Replace(bytes, ReplaceItems);
        }
        #endregion

        #region Replace: WordTemplate Replace Function (Replce To TableRow By Dictionary<String, String>)
        public TableRow Replace(TableRow tableRow, Dictionary<string, string> ReplaceItems)
        {
            TableRow targetTableRow = (TableRow)tableRow.Clone();

            foreach (KeyValuePair<string, string> keyValuePair in ReplaceItems)
            {

                string SearchString = keyValuePair.Key;
                string ReplaceString = keyValuePair.Value;  // 解決換行問題     

                #region 字串替代                                                    
                foreach (Paragraph para in targetTableRow.Descendants<Paragraph>())
                {
                    if (para.InnerText.Contains(SearchString))
                    {                        
                        Run newRun = (Run)para.Descendants<Run>().First().Clone();

                        newRun.Elements<Text>().First().Text = para.InnerText.Replace(SearchString, ReplaceString);

                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(newRun);
                    }
                }
                #endregion
            }

            return targetTableRow;
        }
        #endregion

        #region Replace: WordTemplate Replace Function (Replace To TableRow By Model)
        public TableRow Replace<T>(TableRow tableRow, T ReplaceModel) where T : class
        {
            PropertyInfo[] infos = ReplaceModel.GetType().GetProperties();

            Dictionary<string, string> ReplaceItems = new Dictionary<string, string>();

            foreach (PropertyInfo info in infos)
            {
                string Key = "#" + info.Name + "#";

                string Value = info.GetValue(ReplaceModel, null) == null ? "" : info.GetValue(ReplaceModel, null).ToString();

                ReplaceItems.Add(Key, Value);
            }

            return Replace(tableRow, ReplaceItems);
        }
        #endregion

        #region Replace: WordTemplate Replace Function (Replace From Byte[] To Byte[] By ModelList)
        public byte[] Replace<T>(byte[] bytes, List<T> ReplaceModelList) where T : class
        {
            // 處理Model List 資訊
            PropertyInfo[] infos = ReplaceModelList.First().GetType().GetProperties();

            List<string> ReplaceTags = new List<string>();

            foreach (PropertyInfo info in infos)
            {
                ReplaceTags.Add("#" + info.Name + "#");
            }

            MemoryStream destination = ByteArrayToMemoryStream(bytes);

            using (var wordDoc = WordprocessingDocument.Open(destination, true))
            {
                TableRow TargetRow = wordDoc.MainDocumentPart.Document.Body.Descendants<TableRow>().First((target) => ReplaceTags.All(target.InnerText.Contains));
                TableRow CopyRow = (TableRow)TargetRow.Clone();

                var Inserter = new InsertWordTemplate();
                foreach (T ReplaceItem in ReplaceModelList)
                {
                    Inserter.InsertTableRow(TargetRow, Replace(CopyRow, ReplaceItem));
                }
                TargetRow.Remove();

                wordDoc.Save();
            }

            destination.Position = 0;

            return destination.ToArray();
        }
        #endregion

        #region Replace: WordTemplate Replace Function (Replace From String To Byte[] By ModelList)
        public byte[] Replace<T>(string TemplateFilePath, List<T> ReplaceModelList) where T : class
        {
            byte[] bytes = FilePathToByteArray(TemplateFilePath);

            return Replace(bytes, ReplaceModelList);
        }
        #endregion
    }


}
