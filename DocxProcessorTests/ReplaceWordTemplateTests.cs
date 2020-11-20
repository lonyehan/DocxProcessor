using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace DocxProcessor.Tests
{
    [TestClass]
    public class ReplaceWordTemplateTests
    {
        [TestMethod]
        public void Replace()
        {
            string TemplateFilePath = "C:\\Users\\歐家豪\\Desktop\\side-project\\WordTemplate\\test.docx";
            string OutputFilePath = "C:\\Users\\歐家豪\\Desktop\\side-project\\WordTemplate\\test2.docx";
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            string TestStr = @"一、甄選成績優秀者可請領本校獎學金，請領條件如下：進一步詢問者，請電洽本校。
1. 學費比照新北市立高中音樂班(代辦費另收\n)：
(1)西樂主修：主修成績達82分以上，聽寫成績達70分以上，國文及英文均達 B
+以上者。
(2)國樂主修：主修達86分以上，聽寫成績達60分以上，國文及英文均達 B
+以上者。
2. 可申請免主修學費：主修達82分以上，聽寫達60分以上，國文及英文均達 B
+以上者。
二、為全國第一所中學音樂班，創班50餘年，秉承全人教育理念，教學嚴謹，學、術科並
重，重視生活教育。音樂教學內容豐富多元，定期舉辦國際大師講座、出國巡演。
三、音樂班有專屬大樓，含授課及練習琴房共計80餘間、設備新穎之音樂欣賞教室1間、
大型演奏廳、小型演奏廳。亦備有設備新穎、管理完善之宿舍。
四、本校位於新北市板橋區，可乘捷運板南線於龍山寺站轉乘公車，或於臺北車站、板橋
車站換乘公車，於埔墘站下車即至，交通相當便捷。";

            keyValuePairs.Add("Note", TestStr);
            var Replacer = new ReplaceWordTemplate();
            Assert.IsTrue(Replacer.Replace(TemplateFilePath, OutputFilePath, keyValuePairs));
        }
        [TestMethod]
        public void NewLineString()
        {
            string TestStr = @"一、甄選成績優秀者可請領本校獎學金，請領條件如下：進一步詢問者，請電洽本校。
1. 學費比照新北市立高中音樂班(代辦費另收)：
(1)西樂主修：主修成績達82分以上，聽寫成績達70分以上，國文及英文均達 B
+以上者。
(2)國樂主修：主修達86分以上，聽寫成績達60分以上，國文及英文均達 B
+以上者。
2. 可申請免主修學費：主修達82分以上，聽寫達60分以上，國文及英文均達 B
+以上者。
二、為全國第一所中學音樂班，創班50餘年，秉承全人教育理念，教學嚴謹，學、術科並
重，重視生活教育。音樂教學內容豐富多元，定期舉辦國際大師講座、出國巡演。
三、音樂班有專屬大樓，含授課及練習琴房共計80餘間、設備新穎之音樂欣賞教室1間、
大型演奏廳、小型演奏廳。亦備有設備新穎、管理完善之宿舍。
四、本校位於新北市板橋區，可乘捷運板南線於龍山寺站轉乘公車，或於臺北車站、板橋
車站換乘公車，於埔墘站下車即至，交通相當便捷。";
            Assert.IsFalse(TestStr.IndexOf("\r\n") == -1);
        }
    }
}
