using SolidEdgeConstants;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Documents;
using xml_data_extraction.Features;
using xml_data_extraction.Properties;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction
{
    class program_xml_data_extraction
    {
        [STAThread]
        static void Main(string[] args)
        {
            string rootFolder = "";
            string outputFolder = "";

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--input" && i + 1 < args.Length)
                {
                    rootFolder = args[i + 1];
                }
                else if (args[i] == "--output" && i + 1 < args.Length)
                {
                    outputFolder = args[i + 1];
                }
            }

            try
            {
                if (string.IsNullOrEmpty(rootFolder) || string.IsNullOrEmpty(outputFolder))
                {
                    Console.Error.WriteLine("ERROR: Both --input and --output arguments are required.");
                    System.Environment.Exit(1);
                }

                if (!Directory.Exists(rootFolder))
                {
                    Console.Error.WriteLine($"ERROR: Input folder not found: {rootFolder}");
                    System.Environment.Exit(1);
                }

                if (!Directory.Exists(outputFolder))
                {
                    Console.Error.WriteLine($"ERROR: Output folder not found: {outputFolder}");
                    System.Environment.Exit(1);
                }

                SolidEdgeFramework.Application seApp = null;

                try
                {
                    seApp = (SolidEdgeFramework.Application)MarshalHelper.GetActiveObject("SolidEdge.Application");
                    seApp.DisplayAlerts = false;
                }
                catch
                {
                    Console.Error.WriteLine("ERROR: Could not attach to Solid Edge. Ensure it is running.");
                    System.Environment.Exit(1);
                }

                try
                {
                    var subFiles = Directory.GetFiles(rootFolder, "*.par", SearchOption.AllDirectories);

                    foreach (var subFile in subFiles)
                    {
                        Console.WriteLine($"\n--- Processing: {subFile}");
                        SolidEdgeDocument doc = null;

                        try
                        {
                            List<XElement> featureXmlList = new List<XElement>();
                            Console.WriteLine("  Extracting metadata...");

                            var prop_report = PR01_file_properties_extract.Properties(subFile);
                            featureXmlList.Add(prop_report);

                            var fileInfo = new FileInfo(subFile);
                            if (fileInfo.IsReadOnly)
                            {
                                fileInfo.IsReadOnly = false;
                            }

                            doc = seApp.Documents.Open(subFile);

                            if (doc is SolidEdgePart.PartDocument partDoc)
                            {
                                featureXmlList.Insert(0, PR01_file_properties_extract.UnitsOfMeasure_Extract(partDoc));
                                featureXmlList.Insert(1, PR01_file_properties_extract.BaseStyle_Extract(partDoc));

                                featureXmlList.Add(Documents.DO01_part_data_extractor.PartExtract(partDoc));
                                featureXmlList.Add(MM02_variable_extractor.Variables_extract(partDoc));

                                Sketchs sketches = null;
                                try
                                {
                                    sketches = partDoc.Sketches;
                                    var sketchesElement = new XElement("Sketches", new XAttribute("Count", sketches.Count));

                                    for (int s = 1; s <= sketches.Count; s++)
                                    {
                                        Sketch sketch = null;
                                        try
                                        {
                                            sketch = sketches.Item(s);
                                            sketchesElement.Add(FE15_sketch_extractor.Sketch_Extract(sketch));
                                        }
                                        finally
                                        {
                                            if (sketch != null)
                                            {
                                                Marshal.ReleaseComObject(sketch);
                                                sketch = null;
                                            }
                                        }
                                    }

                                    featureXmlList.Add(sketchesElement);
                                }
                                catch (Exception ex)
                                {
                                    Console.Error.WriteLine($"  Error extracting Sketches: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                }
                                finally
                                {
                                    if (sketches != null)
                                    {
                                        Marshal.ReleaseComObject(sketches);
                                        sketches = null;
                                    }
                                }
                            }
                            else if (doc is SolidEdgePart.SheetMetalDocument sheetDoc)
                            {
                                continue;
                            }
                            else
                            {
                                Console.WriteLine("  Skipped: not a PartDocument");
                            }

                            string parentFolder = Path.GetFileName(Path.GetDirectoryName(subFile));
                            string fileName = Path.GetFileNameWithoutExtension(subFile);
                            string baseName = $"{parentFolder}_{fileName}.xml";
                            string xmlFullPath = Path.Combine(outputFolder, baseName);

                            var rootElement = new XElement("Evaluation");
                            foreach (var elements in featureXmlList)
                            {
                                rootElement.Add(elements);
                            }

                            rootElement.Save(xmlFullPath);
                            Console.WriteLine($"  Saved XML: {xmlFullPath}");
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine($"  Error opening/extracting {subFile}: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (doc != null)
                            {
                                doc.Close(false, Type.Missing, Type.Missing);
                                Marshal.ReleaseComObject(doc);
                                doc = null;
                            }
                        }
                    }

                    Console.WriteLine("\nAll files processed. Disconnecting from Solid Edge.");
                }
                finally
                {
                    if (seApp != null)
                    {
                        Marshal.ReleaseComObject(seApp);
                        seApp = null;
                    }
                }

                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"FATAL ERROR: {ex.Message} | Inner: {ex.InnerException?.Message}");
                System.Environment.Exit(1);
            }
        }
    }
}