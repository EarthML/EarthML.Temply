using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Collections.Generic;

namespace EarthML.Temply.Core
{
    public class Size
    {
        public double width { get; set; }
        public double height { get; set; }

        public Size(int width, int height)
        {
            this.width = width;
            this.height = height;
        }
    }
    public static class ImageHelper
    {
        const string errorMessage = "Could not recognize image format.";

        private static Dictionary<byte[], Func<BinaryReader, Size>> imageFormatDecoders = new Dictionary<byte[], Func<BinaryReader, Size>>()
        {
            { new byte[]{ 0x42, 0x4D }, DecodeBitmap},
            { new byte[]{ 0x47, 0x49, 0x46, 0x38, 0x37, 0x61 }, DecodeGif },
            { new byte[]{ 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, DecodeGif },
            { new byte[]{ 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, DecodePng },
            { new byte[]{ 0xff, 0xd8 }, DecodeJfif },
        };

        /// <summary>
        /// Gets the dimensions of an image.
        /// </summary>
        /// <param name="path">The path of the image to get the dimensions of.</param>
        /// <returns>The dimensions of the specified image.</returns>
        /// <exception cref="ArgumentException">The image was of an unrecognized format.</exception>
        public static Size GetDimensions(string path)
        {
            using (BinaryReader binaryReader = new BinaryReader(File.OpenRead(path)))
            {
                try
                {
                    return GetDimensions(binaryReader);
                }
                catch (ArgumentException e)
                {
                    if (e.Message.StartsWith(errorMessage))
                    {
                        throw new ArgumentException(errorMessage, "path", e);
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the dimensions of an image.
        /// </summary>
        /// <param name="path">The path of the image to get the dimensions of.</param>
        /// <returns>The dimensions of the specified image.</returns>
        /// <exception cref="ArgumentException">The image was of an unrecognized format.</exception>    
        public static Size GetDimensions(BinaryReader binaryReader)
        {
            int maxMagicBytesLength = imageFormatDecoders.Keys.OrderByDescending(x => x.Length).First().Length;

            byte[] magicBytes = new byte[maxMagicBytesLength];

            for (int i = 0; i < maxMagicBytesLength; i += 1)
            {
                magicBytes[i] = binaryReader.ReadByte();

                foreach (var kvPair in imageFormatDecoders)
                {
                    if (magicBytes.StartsWith(kvPair.Key))
                    {
                        return kvPair.Value(binaryReader);
                    }
                }
            }

            throw new ArgumentException(errorMessage, "binaryReader");
        }

        private static bool StartsWith(this byte[] thisBytes, byte[] thatBytes)
        {
            for (int i = 0; i < thatBytes.Length; i += 1)
            {
                if (thisBytes[i] != thatBytes[i])
                {
                    return false;
                }
            }
            return true;
        }

        private static short ReadLittleEndianInt16(this BinaryReader binaryReader)
        {
            byte[] bytes = new byte[sizeof(short)];
            for (int i = 0; i < sizeof(short); i += 1)
            {
                bytes[sizeof(short) - 1 - i] = binaryReader.ReadByte();
            }
            return BitConverter.ToInt16(bytes, 0);
        }

        private static int ReadLittleEndianInt32(this BinaryReader binaryReader)
        {
            byte[] bytes = new byte[sizeof(int)];
            for (int i = 0; i < sizeof(int); i += 1)
            {
                bytes[sizeof(int) - 1 - i] = binaryReader.ReadByte();
            }
            return BitConverter.ToInt32(bytes, 0);
        }

        private static Size DecodeBitmap(BinaryReader binaryReader)
        {
            binaryReader.ReadBytes(16);
            int width = binaryReader.ReadInt32();
            int height = binaryReader.ReadInt32();
            return new Size(width, height);
        }

        private static Size DecodeGif(BinaryReader binaryReader)
        {
            int width = binaryReader.ReadInt16();
            int height = binaryReader.ReadInt16();
            return new Size(width, height);
        }

        private static Size DecodePng(BinaryReader binaryReader)
        {
            binaryReader.ReadBytes(8);
            int width = binaryReader.ReadLittleEndianInt32();
            int height = binaryReader.ReadLittleEndianInt32();
            return new Size(width, height);
        }

        private static Size DecodeJfif(BinaryReader binaryReader)
        {
            while (binaryReader.ReadByte() == 0xff)
            {
                byte marker = binaryReader.ReadByte();
                short chunkLength = binaryReader.ReadLittleEndianInt16();

                if (marker == 0xc0)
                {
                    binaryReader.ReadByte();

                    int height = binaryReader.ReadLittleEndianInt16();
                    int width = binaryReader.ReadLittleEndianInt16();
                    return new Size(width, height);
                }

                binaryReader.ReadBytes(chunkLength - 2);
            }

            throw new ArgumentException(errorMessage);
        }
    }

    public static class Extensions
    {
        public static void InsertAPicture(MainDocumentPart mainPart, string fileName, OpenXmlElement element, string width)
        {


            var size = ImageHelper.GetDimensions(fileName);


            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            const double emusPerInch = 914400;
            const double emusPerCm = 360000;
            double widthInCm = double.Parse(width.Replace("cm", ""));
            double cx = (widthInCm * emusPerCm);
            double cy = cx / (size.width / size.height);

            AddImageToBody(mainPart, mainPart.GetIdOfPart(imagePart), element, (long)cx, (long)cy);

        }

        private static void AddImageToBody(MainDocumentPart mainPart, string relationshipId, OpenXmlElement elementwrapper, long cx, long cy)
        {

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = cx, Cy = cy },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = cx, Cy = cy }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            elementwrapper.RemoveAllChildren();
            elementwrapper.AppendChild(element);
          
          //  elementwrapper.AppendChild(new Paragraph(new Run(element)));
            // Append the reference to body, the element should be in a Run.

        }


        public static void WriteJsonTable(this SdtElement element, string path, MainDocumentPart mainPart)
        {
            var tablejson = JObject.Parse(File.ReadAllText(path));



            var boderSize = tablejson.SelectToken("$.border.outer.size").ToObject<UInt32>();
            var borderColor = tablejson.SelectToken("$.border.outer.color").ToObject<string>();
            var borderType = (BorderValues)tablejson.SelectToken("$.border.outer.type").ToObject<int>();
            // Create an empty table.
            Table table = new Table();

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(borderType),
                        Size = boderSize,
                        Color = borderColor,
                    },
                    new BottomBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(borderType),
                        Size = boderSize,
                        Color = borderColor,
                    },
                    new LeftBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(borderType),
                        Size = boderSize,
                        Color = borderColor,
                    },
                    new RightBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(borderType),
                        Size = boderSize,
                        Color = borderColor,
                    },
                    new InsideHorizontalBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(borderType),
                        Size = boderSize,
                        Color = borderColor,
                    },
                    new InsideVerticalBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(borderType),
                        Size = boderSize,
                        Color = borderColor,
                    }
                )
            );

            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tblProp);

            CreateTableHeader(tablejson, table);

            var rows = tablejson.SelectToken("$.body.rows").ToArray();
            var i = 0;
            foreach (var row in rows)
            {

                TableRow tr = new TableRow();
                foreach (var column in row)
                {
                    var run = new Run(new RunProperties(new Bold(), new Color() { Val = "#000000" }), new Text(column.ToString()));

                    if (column is JObject columnObj && column.SelectToken("$.type").ToString() == "image")
                    {


                      



                        InsertAPicture(mainPart, column.SelectToken("$.path").ToString(), run, column.SelectToken("$.width").ToString());

                    


                    }
                    

                    tr.Append(new TableCell(
                          new TableCellProperties(
                               new Shading()
                               {
                                   Color = "auto",
                                   Fill = i % 2 == 0 ? "auto" : "#eeeeee",
                                   Val = ShadingPatternValues.Clear
                               }),
                          new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), run
                  )
                          ));
                }
                i++;
                table.Append(tr);
            }



            element.RemoveAllChildren();
            element.AppendChild(table);
        }

        private static void CreateTableHeader(JObject tablejson, Table table)
        {
            // Create a row.
            TableRow tr = new TableRow(new TableRowProperties(new TableRowHeight()
            {
                HeightType = new EnumValue<HeightRuleValues>(HeightRuleValues.AtLeast),
                Val = (UInt32)tablejson.SelectTokens("$.header.columns[*].text").Max(c => c.ToString().Length) * 130
            }));
            TableRow trg = new TableRow();

            var headers = tablejson.SelectToken("$.header.columns").ToArray();
            if (!headers.Any())
                return;

            var groupsan = 0;
            var last = "";
            foreach (var header in headers)
            {
                var group = tablejson.SelectTokens($"$.header.groups[?(@.id == '{header.SelectToken("$.group").ToString()}')]").FirstOrDefault() ?? header;

                if ((header.SelectToken("$.group")?.ToString() ?? "") != last)
                {

                    last = header.SelectToken("$.group")?.ToString() ?? "";
                    groupsan = 1;

                }
                else
                {
                    groupsan++;
                }


                {
                    // Create a cell.
                    TableCell tc1 = new TableCell();

                    // Specify the width property of the table cell.




                    var tcp = new TableCellProperties(
                        new TextDirection() { Val = TextDirectionValues.BottomToTopLeftToRight },
                        new TableCellWidth() { Type = TableWidthUnitValues.Auto, Width = Math.Floor(1.0 / headers.Length * 100).ToString() },

                        new Shading()
                        {
                            Color = "auto",
                            Fill = group.SelectToken("$.bg")?.ToString() ?? "#3344aa",
                            Val = ShadingPatternValues.Clear
                        });





                    tc1.Append(tcp);
                    tc1.Append(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(new RunProperties(new Bold(), new Color() { Val = "#ffffff" }), new Text(header.SelectToken("$.text").ToString()))));




                    tr.Append(tc1);
                }
                if (tablejson.SelectToken("$.header.groups")?.Any() ?? false)
                {
                    // Create a cell.
                    TableCell tc1 = new TableCell();

                    // Specify the width property of the table cell.




                    var tcp = new TableCellProperties(
                        new TableCellWidth() { Type = TableWidthUnitValues.Auto, Width = Math.Floor(1.0 / headers.Length * 100).ToString() },
                        new Shading()
                        {
                            Color = "auto",
                            Fill = group.SelectToken("$.bg").ToString(),
                            Val = ShadingPatternValues.Clear
                        },
                        new HorizontalMerge() { Val = new EnumValue<MergedCellValues>(groupsan == 1 ? MergedCellValues.Restart : MergedCellValues.Continue) });





                    tc1.Append(tcp);
                    {
                        tc1.Append(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), groupsan == 1 ? new Run(
                            new RunProperties(
                                new Bold(),

                                new Color() { Val = "#ffffff" }),
                            new Text(group.SelectToken("$.text")?.ToString())) : new Run(new Text())));
                    }




                    trg.Append(tc1);
                }


            }

            table.Append(tr);
            if (tablejson.SelectToken("$.header.groups")?.Any() ?? false)
                table.Append(trg);
        }

        public static void UpdateImageFromPath(this MainDocumentPart mainPart, SdtElement element, string imgPath)
        {

            var picture = element.Descendants<SdtContentPicture>().FirstOrDefault();

            if (picture != null)
            {

                var dr = element.Descendants<Drawing>().FirstOrDefault();
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

                            ImagePart ip = (ImagePart)idpp.OpenXmlPart;

                            using (FileStream fileStream = File.Open(imgPath, FileMode.Open))
                                ip.FeedData(fileStream);
                        }
                        else
                        {
                            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                            blip.Embed = mainPart.GetIdOfPart(imagePart);
                            System.IO.Stream data = new System.IO.MemoryStream(File.ReadAllBytes(imgPath));
                            //  using (FileStream stream = new FileStream(imgPath, FileMode.Open))
                            {
                                imagePart.FeedData(data);
                                data.Close();
                            }
                            Console.WriteLine("test");

                            // element.InnerXml = "<w:rPr><w:noProof /></w:rPr><w:tag w:val=\"MyProvider:CoolImage\" /><w:id w:val=\"2083329445\" /><w:picture /></w:sdtPr><w:sdtEndPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" /><w:sdtContent xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p w:rsidR=\"00646939\" w:rsidP=\"00646939\" w:rsidRDefault=\"00646939\"><w:pPr><w:rPr><w:noProof /></w:rPr></w:pPr><w:r w:rsidRPr=\"00646939\"><w:rPr><w:noProof /></w:rPr><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" wp14:anchorId=\"40D5DA71\" wp14:editId=\"48213C65\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:extent cx=\"4605338\" cy=\"3673658\" /><wp:effectExtent l=\"0\" t=\"0\" r=\"5080\" b=\"3175\" /><wp:docPr id=\"2\" name=\"Picture 2\" /><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\" /></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"\" /><pic:cNvPicPr /></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId6\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"4613233\" cy=\"3679955\" /></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:sdtContent>";
                            // Console.WriteLine(element.InnerXml);
                            // element.Parent.Remove();
                            //  element.Remove();
                            //  mainPart.Document.AppendChild(element);

                        }
                    }
                }
            }

        }
    }
}
