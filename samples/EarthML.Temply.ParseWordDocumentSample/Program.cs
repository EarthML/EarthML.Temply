using EarthML.Temply.Core;
using System;
using System.IO;
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
            }

            

        }
    }
}
