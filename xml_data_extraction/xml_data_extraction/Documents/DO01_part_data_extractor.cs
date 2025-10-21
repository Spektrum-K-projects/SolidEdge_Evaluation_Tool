using System.Runtime.InteropServices;
using System.Xml.Linq;
using SolidEdgePart;
using xml_data_extraction.Features;
using xml_data_extraction.Properties;

namespace xml_data_extraction.Documents
{
    internal class DO01_part_data_extractor
    {
        public static XElement PartExtract(PartDocument partDoc)
        {
            Models models = null;
            Model model = null;
            SolidEdgePart.Features features = null;

            var partElements = new XElement("Part");

            try
            {
                ////File Properties Extract
                //Console.WriteLine("  Extracting metadata...");
                //FilePropertiesExtract.Properties(partDoc);

                //Loading the Models
                Console.WriteLine("  Extracting features...");
                models = partDoc.Models;

                for (int m = 1; m <= models.Count; m++)
                {
                    //Loading each Model
                    model = models.Item(m);
                    Console.WriteLine($"   Model [{m}]: {model.Name}");

                    partElements.Add(new XElement (PR02_part_physical_properties_extractor.Properties(model)));

                    //Identifying the features in the Model
                    features = model.Features;

                    // Call each feature extractor in turn:
                    for (int f = 1; f <= features.Count; f++)
                    {
                        var feat = features.Item(f);
                        int output = feat.Type;
                        Console.WriteLine(output);

                        if (output == 462094706)
                        {
                            var extrusions = model.ExtrudedProtrusions;
                            partElements.Add(new XElement(FE01_extruded_protrusion_extractor.Protrusions((ExtrudedProtrusions)extrusions)));
                            continue;
                        }
                        else if (output == 462094722)
                        {
                            var holes = model.Holes;
                            partElements.Add(new XElement(FE02_hole_extractor.Holes((Holes)holes)));
                            continue;
                        }
                        else if (output == 462094714)
                        {
                            var cutouts = model.ExtrudedCutouts;
                            partElements.Add(new XElement(FE03_cutout_extractor.Cutouts((ExtrudedCutouts)cutouts)));
                            continue;
                        }
                        else if (output == 462094742)
                        {
                            partElements.Add(new XElement(FE04_edge_features_extractor.Chamfer((Chamfer)feat)));
                            continue;
                        }
                        else if (output == 462094710)
                        {
                            partElements.Add(new XElement(FE05_revolve_extractor.Revolved_protrusion((RevolvedProtrusion)feat)));
                            continue;
                        }
                        else if (output == 462094730)
                        {
                            partElements.Add(new XElement(FE06_rib_extractor.Rib((Rib)feat)));
                            continue;
                        }
                        else if (output == 462094738)
                        {
                            partElements.Add(new XElement(FE04_edge_features_extractor.Round((Round) feat)));
                            continue;
                        }
                        else if (output == -416228998)
                        {
                            partElements.Add(new XElement(FE07_pattern_extractor.Pattern((Pattern)feat)));
                            continue;
                        }
                        else if (output == 462094734)
                        {
                            partElements.Add(new XElement(FE08_thinwall_extractor.ThinWall((Thinwall)feat)));
                            continue;
                        }
                        else if (output == 438630050)
                        {
                            partElements.Add(new XElement(FE09_thinregion_extractor.ThinRegion((Thin)feat)));
                            continue;
                        }
                        else if (output == 1908287958)
                        {
                            partElements.Add(new XElement(FE10_mirror_extractor.MirrorPart((MirrorPart)feat)));
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($"Skipping feature");
                            continue;
                        }
                    }
                    //---add more as needed...

                    Marshal.ReleaseComObject(features);
                    Marshal.ReleaseComObject(model);
                }
                Marshal.ReleaseComObject(models);
            }
            catch (Exception ex)
            {
                Console.WriteLine("  PartDataExtractor error: " + ex.Message);
            }

            Console.WriteLine($"The part data has been extracted");
            return partElements;
        }
    }
}
