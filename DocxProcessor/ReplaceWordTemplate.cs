using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Xceed.Words.NET;

namespace DocxProcessor
{
    public class ReplaceWordTemplate
    {
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
            try
            {
                if (File.Exists(TemplateFilePath) == true)
                {
                    byte[] byteArray = File.ReadAllBytes(TemplateFilePath);
                    using (var stream = new MemoryStream())
                    {
                        stream.Write(byteArray, 0, byteArray.Length);
                        using (DocX document = DocX.Load(stream))
                        {
                            

                            foreach (KeyValuePair<string, string> keyValuePair in ReplaceItems)
                            {
                                string SearchString = keyValuePair.Key;
                                string ReplaceString = keyValuePair.Value;
                                document.ReplaceText(SearchString, ReplaceString);
                            }                            

                            document.Save();
                        }

                        // Save the file with the new name
                        stream.Position = 0;
                        return stream.ToArray();
                    }
                }
                else
                {
                    throw new FileNotFoundException("Template File is not exist!");
                }
            }
            catch (InvalidDataException e)
            {
                throw e;
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
            try
            {
                FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

                BinaryWriter bw = new BinaryWriter(fs);

                bw.Write(Replace(TemplateFilePath, ReplaceItems));

                bw.Close();
                fs.Close();

                return true;
            }
            catch (InvalidDataException e)
            {
                throw e;
            }
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
            try
            {
                PropertyInfo[] infos = ReplaceModel.GetType().GetProperties();

                Dictionary<string, string> ReplaceItems = new Dictionary<string, string>();

                foreach (PropertyInfo info in infos)
                {
                    ReplaceItems.Add("#" + info.Name + "#", info.GetValue(ReplaceModel, null).ToString());
                }

                return Replace(TemplateFilePath, ReplaceItems);
            }
            catch (InvalidDataException e)
            {
                throw e;
            }
        }
        #endregion
    }
}
