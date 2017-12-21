# EarthML.Temply
A small framework for creating templates with word and replacing content

A small demo is given below and input, data and output is all found in the data folder. 

The main goal of this framework is to quickly turn a word file into a template, run the tooling to extract metadata that can be feed into other tooling to generate images and data needed for the providers. 

Then running the second time with the providers given to actually update the word file.

Btw, its just a few hours of work at this stage so try it out and give some feedback.

## Samples
Run the sample and see the output
```
MyProvider:ReportTitle
MyProvider:CoolImage
        MyProvider:ReportTitle,
        MyProvider:CoolImage, image=1500x1200
[
  {
    "TagName": "MyProvider:ReportTitle",
    "Format": ""
  },
  {
    "IsImage": true,
    "PxWidth": 1500,
    "PxHeight": 1200,
    "TagName": "MyProvider:CoolImage",
    "Format": ""
  }
]
```

using the following demo provider
```
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
                mainPart.UpdateImageFromPath(element, "../../data/Hello-Im-Awesome.jpg");
            }
            else
            {
                element.Descendants<Text>().First().Text = "Hello World";
                element.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove());
            }

            return base.UpdateElement(mainPart, element, tag);
        }
    }
```
the sample will replace the text templates with hellow world and image template parts with a Im-Awesome image.

```
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


                File.WriteAllBytes(Path.ChangeExtension(args[0], ".updated.docx"),p.stream.ToArray());

            }

            

        }
    }
```
