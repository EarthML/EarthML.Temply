using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EarthML.Temply.Core
{
    public class Processor
    {
        public List<TemplateReplacement> Metadata = new List<TemplateReplacement>();

        public List<IProcessorProvider> Providers = new List<IProcessorProvider>();

        public async Task ProcessElements(MainDocumentPart mainPart, IEnumerable<SdtElement> elements)
        {
            var providers = Providers.ToLookup(k => k.Name.ToLower());
            foreach (SdtElement sdt in elements)
            {
                var tag = sdt.SdtProperties.GetFirstChild<Tag>()?.Val;
                if (tag != null && tag.HasValue)
                {
                    //Console.WriteLine(sdt.InnerXml);

                    //Console.WriteLine();
                    //Console.WriteLine();
                    //Console.WriteLine();

                    var tagname = tag.Value;
                    var idx = tagname.IndexOf('|');
                    var format = string.Empty;
                    if (idx != -1)
                    {
                        format = tagname.Substring(idx + 1);
                        tagname = tagname.Substring(0, idx);
                    }





                    var picture = sdt.SdtProperties.GetFirstChild<SdtContentPicture>();
                    if (picture != null)
                    {
                        var dr = sdt.Descendants<Drawing>().FirstOrDefault();
                        if (dr != null)
                        {
                            const double emusPerInch = 914400;
                            const double emusPerCm = 360000;
                            //Resize picture placeholder


                            Metadata.Add(new TemplateImageReplacement
                            {
                                TagName = tagname,
                                Format = format,
                                IsImage = true,
                                PxHeight = (int)(dr.Inline.Extent.Cy / emusPerInch * 300),
                                PxWidth = (int)(dr.Inline.Extent.Cx / emusPerInch * 300)
                            });

                        }
                    }
                    else
                    {
                        Metadata.Add(new TemplateReplacement { TagName = tagname, Format = format });
                    }

                    Console.WriteLine(tagname);
                    foreach (var provider in providers[tagname.Split(':').First().ToLower()] ?? Enumerable.Empty<IProcessorProvider>())
                    {
                        await provider.UpdateElement(mainPart, sdt, Metadata.Last());
                    }




                }
            }

          


        }
        public MemoryStream stream { get; set; }
        public async Task ProcessDocument(byte[] data)
        {

            stream = new MemoryStream();
            using (var copyFrom = new MemoryStream(data))
            {
                await copyFrom.CopyToAsync(stream);
                stream.Seek(0, SeekOrigin.Begin);
            }

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(stream, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;


                foreach (var header in mainPart.HeaderParts)
                {
                    await ProcessElements(mainPart, header.RootElement.Descendants<SdtElement>());
                }

                foreach (var footer in mainPart.FooterParts)
                {
                    await ProcessElements(mainPart, footer.RootElement.Descendants<SdtElement>());
                }

                await ProcessElements(mainPart, mainPart.Document.Body.Descendants<SdtElement>());

                var tags = Metadata.ToLookup(k => k.TagName);
                foreach (var tag in tags)
                {
                    //Console.WriteLine($"{tag.Key}");
                    foreach (var content in tag)
                    {
                        if (content is TemplateImageReplacement imageTag)
                        {
                            Console.WriteLine($"\t{tag.Key},{(string.IsNullOrEmpty(content.Format) ? "" : $"format={content.Format}")} image={imageTag.PxWidth}x{imageTag.PxHeight}");

                        }
                        else
                        {
                            Console.WriteLine($"\t{tag.Key},{(string.IsNullOrEmpty(content.Format) ? "" : $"format={content.Format}")}");
                        }

                    }
                }

                foreach (var a in mainPart.ImageParts)
                {
                    Console.WriteLine(a.GetStream().Length);
                    Console.WriteLine(a.ContentType);
                }

                wdDoc.MainDocumentPart.Document.Save();
            
                wdDoc.Save();
                wdDoc.Close();
                stream.Flush();

               
            }

        }
    }
}
