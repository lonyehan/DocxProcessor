using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using DocxProcessor;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class ConvertWordTemplateTests
    {   
        [TestMethod]
        public void Case1()
        {
            //string InFilePath = GetTestDataFolder("WordTemplate\\test.docx");
            string InFilePath = @"C:\\Users\\JasonJian\\Desktop\\sideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\test2.docx";
            string OutFilePath = @"C:\\Users\\JasonJian\\Desktop\\sideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\test2.html";

            //convert Word to PDF

            var Coverter = new ConvertWordTemplate();
            //Coverter.ConvertWordToPDF(InFilePath);

            //Assert.IsTrue(Coverter.ConvertWordToPDF(InFilePath) == InFilePath);
            string tmp = Coverter.ConvertWordToHTML(InFilePath);
            //Coverter.ConvertHTMLToPDF(tmp, OutFilePath);
            
            FileStream fs = new FileStream(OutFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Coverter.ConvertWordToHTML(InFilePath));
            //Coverter.ConvertWordToHTML(InFilePath);

            bw.Close();

            fs.Close();
            

        }
        public string GetTestDataFolder(string testDataFolder)
        {
            string startupPath = AppDomain.CurrentDomain.BaseDirectory;
            string[] pathItems = startupPath.Split(Path.DirectorySeparatorChar);
            List<string> resultPath = new List<string>();
            foreach (string pathItem in pathItems)
            {
                if (pathItem == "bin")
                    break;
                resultPath.Add(pathItem);
            }
            string projectPath = String.Join(Path.DirectorySeparatorChar.ToString(), resultPath);
            string finalResult = Path.Combine(projectPath, testDataFolder);
            return finalResult;
        }

    }
    //test
}
