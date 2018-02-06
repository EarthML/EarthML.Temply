using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EarthML.Temply.Core;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;


namespace EarthML.Temply.ParseWordDocumentSample
{

    public class MyProvider : BaseProcessorProvider
    {
        public MyProvider()
        {
            Name = nameof(MyProvider);
        }
        public override Task UpdateElement(MainDocumentPart mainPart, SdtElement element, TemplateReplacement tag)
        {
            if (tag is TemplateImageReplacement image)
            {
                mainPart.UpdateImageFromPath(element, "../../../../../data/Hello-Im-Awesome.jpg");
            }
            else
            {
                 
                if (tag.TagName == $"{nameof(MyProvider)}:Table")
                {
                    element.WriteJsonTable("../../../tablejson.json", mainPart); 

                 //   InsertAPicture(mainPart, "../../../../../data/Hello-Im-Awesome.jpg",element);

                }
                else
                {  
                    element.Descendants<Text>().First().Text = "Hello World";
                    element.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove());
                }
            }

            return base.UpdateElement(mainPart, element, tag);
        }


      

    }


    class Program
    {

        static async Task Main(string[] args)
        {

            var p = new Processor();
            p.Providers.Add(new MyProvider());

            //Allow to read a word file currently opened
            using (var fs = new FileStream(args[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var bytes = new byte[fs.Length];
                fs.Read(bytes, 0, (int)fs.Length);
                await p.ProcessDocument(bytes);

                await p.stream.FlushAsync();

                Console.WriteLine(JToken.FromObject(p.Metadata));


                File.WriteAllBytes(Path.ChangeExtension(args[0], ".updated.docx"), p.stream.ToArray());

            }



        }
    }
}
