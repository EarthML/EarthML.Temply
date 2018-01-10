using DocumentFormat.OpenXml;
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
                Val = (UInt32)tablejson.SelectTokens("$.header.groups[*].text").Max(c => c.ToString().Length) * 120
            }));
            TableRow trg = new TableRow();

            var headers = tablejson.SelectToken("$.header.columns").ToArray();
            var groupsan = 0;
            var last = "";
            foreach (var header in headers)
            {
                var group = tablejson.SelectTokens($"$.header.groups[?(@.id == '{header.SelectToken("$.group").ToString()}')]").FirstOrDefault();

                if (header.SelectToken("$.group").ToString() != last)
                {

                    last = header.SelectToken("$.group").ToString();
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
                            Fill = group.SelectToken("$.bg").ToString(),
                            Val = ShadingPatternValues.Clear
                        });





                    tc1.Append(tcp);
                    tc1.Append(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(new RunProperties(new Bold(), new Color() { Val = "#ffffff" }), new Text(header.SelectToken("$.text").ToString()))));




                    tr.Append(tc1);
                }

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
                    }
                }
            }
        }
    }
}
