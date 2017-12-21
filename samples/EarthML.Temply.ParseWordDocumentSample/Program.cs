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
    


    class Program
    {
      
        static async Task Main(string[] args)
        {
            
            var p = new Processor();
           

            using (var fs = new FileStream(args[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var bytes = new byte[fs.Length];
                fs.Read(bytes, 0, (int)fs.Length);
                await p.ProcessDocument(bytes);

                await p.stream.FlushAsync();

                Console.WriteLine(JToken.FromObject(p.Metadata));


                File.WriteAllBytes(Path.ChangeExtension(args[0], ".updated.docx"),p.stream.ToArray());

            }

            

        }
    }
}
