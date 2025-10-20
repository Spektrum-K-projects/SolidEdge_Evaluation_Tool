using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
                            partElements.Add(new XElement (FE01_extruded_protrusion_extractor.Protrusion((ExtrudedProtrusion)feat)));
                            continue;
                        }
                        else if (output == 462094722)
                        {
                            partElements.Add(new XElement(FE02_hole_extractor.Hole((Hole)feat)));
                            continue;
                        }
                        else if (output == 462094714)
                        {
                            partElements.Add(new XElement(FE03_cutout_extractor.Cutout((ExtrudedCutout)feat)));
                            continue;
                        }
                        else if (output == 462094742)
                        {
                            partElements.Add(new XElement(FE04_chamfer_extractor.Chamfer((Chamfer)feat)));
                            continue;
                        }
                        else if (output == 462094710)
                        {
                            partElements.Add(new XElement(FE05_revolved_protrusion_extractor.Revolve((RevolvedProtrusion)feat)));
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

            return partElements;
        }
    }
}
