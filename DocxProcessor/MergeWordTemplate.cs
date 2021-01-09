using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace DocxProcessor
{
    public class MergeWordTemplate
    {
        public byte[] MergeDocxsIntoOne(List<Stream> documents)
        {
            var result = new MemoryStream();

            var sources = new List<Source>();

            foreach (var stream in documents)
            {
                var tempms = new MemoryStream();
                stream.CopyToAsync(tempms);
                sources.Add(new Source(new WmlDocument(stream.Length.ToString(), tempms), true));
            }

            var mergedDoc = DocumentBuilder.BuildDocument(sources);

            result.Write(mergedDoc.DocumentByteArray, 0, mergedDoc.DocumentByteArray.Length);

            result.Position = 0;

            return result.ToArray();
        }
    }
}
