﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EarthML.Temply.Core
{
    public class TemplateReplacement
    {
        public string TagName { get; set; }       
        public string Format { get; set; }
    }
    public class TemplateImageReplacement : TemplateReplacement
    {
        public bool IsImage { get; set; }
        public int PxWidth { get; set; }
        public int PxHeight { get; set; }
    }

    public class Processor
    {
        public List<TemplateReplacement> Metadata = new List<TemplateReplacement>();

        public async Task ProcessElements(MainDocumentPart mainPart, IEnumerable<SdtElement> elements)
        {
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
                            var blip = dr.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                            if (blip != null)
                            {
                                var embed = blip.Embed;
                                if (embed != null)
                                {
                                  
                                    IdPartPair idpp = mainPart.Parts
                                        .Where(pa => pa.RelationshipId == embed).FirstOrDefault();
                                    if (idpp != null)
                                    {

                                        Drawing sdtImage = sdt.Descendants<Drawing>().First();

                                        const int emusPerInch = 914400;
                                        const int emusPerCm = 360000;
                                        //Resize picture placeholder
                                        

                                        Metadata.Add(new TemplateImageReplacement { TagName = tagname, Format = format, IsImage = true,
                                            PxHeight = (int)( sdtImage.Inline.Extent.Cy / emusPerInch * 300 ),
                                            PxWidth =(int)( sdtImage.Inline.Extent.Cx / emusPerInch * 300)
                                        });

                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Metadata.Add(new TemplateReplacement { TagName = tagname, Format = format});
                    }





                  
                }
            }


        }

        public async Task ProcessDocument(byte[] data)
        {
            var stream = new MemoryStream(data);
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(stream, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;


                foreach (var header in mainPart.HeaderParts)
                {
                    await ProcessElements(mainPart,header.RootElement.Descendants<SdtElement>());
                }

                foreach (var footer in mainPart.FooterParts)
                {
                    await ProcessElements(mainPart,footer.RootElement.Descendants<SdtElement>());
                }

                await ProcessElements(mainPart,mainPart.Document.Body.Descendants<SdtElement>());

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

            }

        }
    }
}