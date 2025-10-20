using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Collections.Generic;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using xml_data_extraction.Properties;
using xml_data_extraction.Documents;

namespace xml_data_extraction
{
    class program_xml_data_extraction
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Enter Source Folder Path: ");
            string rootFolder = Console.ReadLine().Trim();

            if (!Directory.Exists(rootFolder))
            {
                Console.WriteLine("Folder not found.");
                return;
            }

            Console.WriteLine("Enter Destination Folder Path: ");
            string outputFolder = Console.ReadLine().Trim();

            if (!Directory.Exists(outputFolder))
            {
                Console.WriteLine("Folder not found.");
                return;
            }

            SolidEdgeFramework.Application seApp = null;

            try
            {
                seApp = (SolidEdgeFramework.Application)MarshalHelper.GetActiveObject("SolidEdge.Application");
            }
            catch
            {
                Console.WriteLine("Could not attach to Solid Edge. Ensure it is running.");
                return;
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
                    var prop_report = PR01_file_properties_extract.Properties(subFile);
                    featureXmlList.Add(prop_report);

                    doc = seApp.Documents.Open(subFile);

                    if (doc is SolidEdgePart.PartDocument partDoc)
                    {

                        featureXmlList.Add(DO01_part_data_extractor.PartExtract(partDoc));
                    }

                    else if (doc is SolidEdgePart.SheetMetalDocument sheetDoc)
                    {
                        //---Add Sheet Document Extractor function code here...
                        continue;
                    }

                    //---More Document functions to be added here

                    else
                    {
                        Console.WriteLine("  Skipped: not a PartDocument");
                    }

                    string parentFolder = Path.GetFileName(Path.GetDirectoryName(subFile));
                    string fileName = Path.GetFileNameWithoutExtension(subFile);
                    string baseName = $"{parentFolder}_{fileName}.xml";
                    //string xmlName = baseName + "_Properties.xml";
                    string xmlFullPath = Path.Combine(outputFolder, baseName);

                    // now save to that file
                    //prop_report.Save(xmlFullPath);

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
                    Console.WriteLine("  Error opening/extracting: " + ex.Message);
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


        }
    }
}
