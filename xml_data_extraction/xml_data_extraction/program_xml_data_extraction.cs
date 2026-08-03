using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Collections.Generic;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using SolidEdgeGeometry;
using SolidEdgeConstants;
using xml_data_extraction.Properties;
using xml_data_extraction.Documents;

namespace xml_data_extraction
{
    class program_xml_data_extraction
    {
        [STAThread]
        static void Main(string[] args)
        {
            // --- MODIFICATION: Add variables for our new arguments ---
            string rootFolder = "";
            string outputFolder = "";

            // --- MODIFICATION: Simple loop to parse command-line args ---
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

            // --- MODIFICATION: Add full try...catch block for error handling ---
            try
            {
                // --- MODIFICATION: Validate arguments ---
                if (string.IsNullOrEmpty(rootFolder) || string.IsNullOrEmpty(outputFolder))
                {
                    Console.Error.WriteLine("ERROR: Both --input and --output arguments are required.");
                    System.Environment.Exit(1); // --- MODIFICATION: Exit with failure code
                }

                // --- MODIFICATION: Console.ReadLine() calls REMOVED ---

                if (!Directory.Exists(rootFolder))
                {
                    // --- MODIFICATION: Write errors to Console.Error ---
                    Console.Error.WriteLine($"ERROR: Input folder not found: {rootFolder}");
                    System.Environment.Exit(1); // --- MODIFICATION: Exit with failure code
                }

                if (!Directory.Exists(outputFolder))
                {
                    Console.Error.WriteLine($"ERROR: Output folder not found: {outputFolder}");
                    System.Environment.Exit(1); // --- MODIFICATION: Exit with failure code
                }

                SolidEdgeFramework.Application seApp = null;

                try
                {
                    seApp = (SolidEdgeFramework.Application)MarshalHelper.GetActiveObject("SolidEdge.Application");
                }
                catch
                {
                    Console.Error.WriteLine("ERROR: Could not attach to Solid Edge. Ensure it is running.");
                    System.Environment.Exit(1); // --- MODIFICATION: Exit with failure code
                }

                var subFiles = Directory.GetFiles(rootFolder, "*.par", SearchOption.AllDirectories);

                foreach (var subFile in subFiles)
                {
                    Console.WriteLine($"\n--- Processing: {subFile}");
                    SolidEdgeDocument doc = null;

                    try
                    {
                        //File Properties Extract
                        List<XElement> featureXmlList = new List<XElement>();
                        Console.WriteLine("  Extracting metadata...");
                        // --- MODIFICATION: Assuming these are in the same namespace or project
                        var prop_report = PR01_file_properties_extract.Properties(subFile);
                        featureXmlList.Add(prop_report);

                        doc = seApp.Documents.Open(subFile);

                        if (doc is SolidEdgePart.PartDocument partDoc)
                        {

                            featureXmlList.Add(Documents.DO01_part_data_extractor.PartExtract(partDoc));
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
                        // --- MODIFICATION: Write errors to Console.Error ---
                        Console.Error.WriteLine($"  Error opening/extracting {subFile}: " + ex.Message);
                        // Don't exit here, just continue to the next file
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            doc.Close();
                            Marshal.ReleaseComObject(doc);
                        }
                    }
                }

                Console.WriteLine("\nAll files processed. Disconnecting from Solid Edge.");
                Marshal.ReleaseComObject(seApp);

                // --- MODIFICATION: Exit with success code ---
                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {
                // --- MODIFICATION: Catch any unexpected fatal errors ---
                Console.Error.WriteLine($"FATAL ERROR: {ex.Message}");
                System.Environment.Exit(1);
            }
        }
    }
}
