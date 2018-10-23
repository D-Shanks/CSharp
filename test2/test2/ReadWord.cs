using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test2
{
    class ReadWord
    {
        public void read()
        {
            string fileName = @"D:\Program Files\VS2017_workplace\C#_Git_Work\CSharp\test2\科研细则.docx";
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(fileName, false))
            {
                // Create a Body object.
                DocumentFormat.OpenXml.Wordprocessing.Body body =
                    wordprocessingDocument.MainDocumentPart.Document.Body;

                // Get the Run elements after the specified element.
                Console.WriteLine("Run elements after the first child are: ");
                foreach (var paragraph in body.Elements())
                {
                    Console.WriteLine(paragraph.InnerText);
                }
                Console.ReadKey();
            }
        }
    }
}
