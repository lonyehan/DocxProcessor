using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class GetTextWordTemplateTests
    {
        [TestMethod]
        public void getAllTextCase1()
        {
            string TemplateFilePath = "C:\\Users\\lonye\\Desktop\\SideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\GetTextCase1.docx";
            List<string> TemplateContent = new List<string>();
            TemplateContent.Add("Test");
            var wordProccessor = new GetTextWordTemplate();
            List<string> ResultList = wordProccessor.getAllText(TemplateFilePath);

            Assert.IsTrue(TemplateContent.Count == ResultList.Count);

            for (int i = 0; i < ResultList.Count; i++)
            {
                string compareString = TemplateContent[i];
                string resultString = ResultList[i];
                Assert.IsTrue(resultString == compareString);
            }                        
        }
        [TestMethod]
        public void getAllTextCase2()
        {
            string TemplateFilePath = "C:\\Users\\lonye\\Desktop\\SideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\GetTextCase2.docx";
            List<string> TemplateContent = new List<string>();
            TemplateContent.Add("Test");
            TemplateContent.Add("再測試");
            var wordProccessor = new GetTextWordTemplate();
            List<string> ResultList = wordProccessor.getAllText(TemplateFilePath);

            Assert.IsTrue(TemplateContent.Count == ResultList.Count);

            for (int i = 0; i < ResultList.Count; i++)
            {
                string compareString = TemplateContent[i];
                string resultString = ResultList[i];
                Assert.IsTrue(resultString == compareString);
            }
        }
        [TestMethod]
        public void getAllTextCase3()
        {
            string TemplateFilePath = "C:\\Users\\lonye\\Desktop\\SideProject\\WordProcessor\\DocxProcessorTests\\WordTemplate\\GetTextCase3.docx";
            List<string> TemplateContent = new List<string>();
            TemplateContent.Add("Test");
            TemplateContent.Add("再測試");
            TemplateContent.Add("再測試");
            TemplateContent.Add("再測試");
            var wordProccessor = new GetTextWordTemplate();
            List<string> ResultList = wordProccessor.getAllText(TemplateFilePath);

            Assert.IsTrue(TemplateContent.Count == ResultList.Count);

            for (int i = 0; i < ResultList.Count; i++)
            {
                string compareString = TemplateContent[i];
                string resultString = ResultList[i];
                Assert.IsTrue(resultString == compareString);
            }
        }
    }
}
