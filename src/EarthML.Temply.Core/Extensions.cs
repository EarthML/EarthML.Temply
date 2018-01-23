using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;

namespace EarthML.Temply.Core
{
    public static class Extensions
    {
        public static void WriteJsonTable(this SdtElement element, string path)
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

                    if (column.Type == JTokenType.Object)
                    {
                        var type = column.SelectToken("$.type").ToString();
                        if (type == "image")
                        {

                            var imageElement =
                            new Drawing(
                                 new Inline(
                                     new Extent() { Cx = 4605338, Cy = 3673658 },
                                     new EffectExtent()
                                     {
                                         LeftEdge = 0L,
                                         TopEdge = 0L,
                                         RightEdge = 5080L,
                                         BottomEdge = 3175L
                                     },
                                     new DocProperties()
                                     {
                                         Id = (UInt32Value)1U,
                                         Name = "Picture 1"
                                     },
                                     new NonVisualGraphicFrameDrawingProperties(
                                         new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                                     new DocumentFormat.OpenXml.Drawing.Graphic(
                                         new DocumentFormat.OpenXml.Drawing.GraphicData(
                                             new Picture(
                                                 new DocumentFormat.OpenXml.Drawing.NonVisualPictureProperties(
                                                     new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties()
                                                     {
                                                         Id = (UInt32Value)0U,
                                                         Name = "New Bitmap Image.jpg"
                                                     },
                                                     new DocumentFormat.OpenXml.Drawing.NonVisualPictureDrawingProperties()),
                                                 new DocumentFormat.OpenXml.Drawing.BlipFill(
                                                     new DocumentFormat.OpenXml.Drawing.Blip(
                                                         new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                                             new DocumentFormat.OpenXml.Drawing.BlipExtension()
                                                             {
                                                                 Uri =
                                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                             })
                                                     )
                                                     {
                                                         CompressionState =
                                                         DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                                     },
                                                     new DocumentFormat.OpenXml.Drawing.Stretch(
                                                         new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                                 new DocumentFormat.OpenXml.Drawing.ShapeProperties(
                                                     new DocumentFormat.OpenXml.Drawing.Transform2D(
                                                         new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                                         new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 4605338, Cy = 3673658L }),
                                                     new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                                         new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                                     )
                                                     { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
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

                            var block = new SdtBlock(
                                new SdtProperties(
                                    new SdtContentPicture(),
                                    new Tag { Val = column.SelectToken("$.tag").ToString() }
                                ),
                               new Paragraph(new Run(imageElement))
                                    );
                            // var tag = new TemplateReplacement { TagName = column.SelectToken("$.tag").ToString() };
                            tr.Append(new TableCell(block));
                          
                        }
                        else
                        {
                            var block = new SdtBlock(
                            new SdtProperties(new Tag { Val = column.SelectToken("$.tag").ToString() }),
                          new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                    new Run(new RunProperties(new Bold(), new Color() { Val = "#000000" }), new Text(column.ToString())))
                            );

                            // var tag = new TemplateReplacement { TagName = column.SelectToken("$.tag").ToString() };
                            tr.Append(new TableCell(block));
                            //,
                        }
                    }
                    else
                    {


                        tr.Append(new TableCell(
                            new TableCellProperties(
                                 new Shading()
                                 {
                                     Color = "auto",
                                     Fill = i % 2 == 0 ? "auto" : "#eeeeee",
                                     Val = ShadingPatternValues.Clear
                                 }),
                            new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                    new Run(new RunProperties(new Bold(), new Color() { Val = "#000000" }), new Text(column.ToString())))
                            ));
                    }
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
                            mainPart.GetIdOfPart(imagePart);
                            using (FileStream stream = new FileStream(imgPath, FileMode.Open))
                            {
                                imagePart.FeedData(stream);
                            }
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
