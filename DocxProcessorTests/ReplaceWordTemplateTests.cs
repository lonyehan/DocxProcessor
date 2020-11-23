using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class ReplaceWordTemplateTests
    {
        [TestMethod]
        public void Replace()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\DocxProcessor\\DocxProcessorTests\\WordTemplateTest\\test.docx";
            string OutputFilePath2 = "C:\\Users\\歐家豪\\source\\repos\\DocxProcessor\\DocxProcessorTests\\WordTemplateTest\\test3.docx";
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            string TestStr = @"1.	符合報名條件及門檻者，依選校登記序號現場分發。
2.	本校升學績效優質，超高國立大學錄取率：109年第17屆畢業班，國立大學錄取率高達96%。
3.	課程以分組教學，並包含多種適性多元課程。擁有全新數位藝術與設計教室，設計與電繪課程、版畫課程、插畫創意風格課程與素描、水彩、水墨書畫等專業課程；設備、師資與課程規劃最健全，教學與輔導最用心!
4.	備有縝密的專車路線與4人一寢冷氣宿舍，優質環境歡迎蒞校參觀或來電詢問(037-868680分機204)。
";

            keyValuePairs.Add("#Name#", TestStr);
            keyValuePairs.Add("#NO#", "測試");
            var Replacer = new ReplaceWordTemplate();            

            FileStream fs = new FileStream(OutputFilePath2, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Replacer.Replace(TemplateFilePath, keyValuePairs));

            bw.Close();

            fs.Close();
        }       
        [TestMethod]
        public void ReplaceToFile()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\DocxProcessor\\DocxProcessorTests\\WordTemplateTest\\test.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\DocxProcessor\\DocxProcessorTests\\WordTemplateTest\\test2.docx";
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            string TestStr = @"1.	符合報名條件及門檻者，依選校登記序號現場分發。
2.	本校升學績效優質，超高國立大學錄取率：109年第17屆畢業班，國立大學錄取率高達96%。
3.	課程以分組教學，並包含多種適性多元課程。擁有全新數位藝術與設計教室，設計與電繪課程、版畫課程、插畫創意風格課程與素描、水彩、水墨書畫等專業課程；設備、師資與課程規劃最健全，教學與輔導最用心!
4.	備有縝密的專車路線與4人一寢冷氣宿舍，優質環境歡迎蒞校參觀或來電詢問(037-868680分機204)。
";

            keyValuePairs.Add("#Name#", TestStr);
            keyValuePairs.Add("#1#", TestStr);
            keyValuePairs.Add("#NO#", "測試");
            var Replacer = new ReplaceWordTemplate();
            Replacer.Replace(TemplateFilePath, OutputFilePath, keyValuePairs);
        }
        [TestMethod]
        public void ReplaceByModel()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\source\\repos\\DocxProcessor\\DocxProcessorTests\\WordTemplateTest\\test.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\source\\repos\\DocxProcessor\\DocxProcessorTests\\WordTemplateTest\\testByModel.docx";
            TestModel test = new TestModel();
            test.NO = 200;
            test.Name = "Test";
            test.手機 = "0905337291";
            test.Date = new DateTime(2020,01,28);

            var Replacer = new ReplaceWordTemplate();            
            FileStream fs = new FileStream(OutputFilePath, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(Replacer.Replace(TemplateFilePath, test));

            bw.Close();

            fs.Close();
        }
        [TestClass]
        public class TestModel
        {
            [Display(Name = "編號")]
            public int NO { get; set; }
            public string Name { get; set; }
            public string 手機 { get; set; }
            public DateTime? Date { get; set; }

        }

    }
}
