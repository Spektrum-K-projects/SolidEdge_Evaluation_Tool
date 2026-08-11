using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    //Lofted Protrusion & Lofted Cutout Data Extraction
    internal class FE17_loft_extractor
    {
        public static XElement Lofted_Protrusion(LoftedProtrusion loftedProtrusion)
        {
            XElement loftedProtrusionElements = new XElement("LoftedProtrusion", new XAttribute("Type", 2057842144));

            try
            {
                try { loftedProtrusionElements.Add(new XElement("name", loftedProtrusion.Name)); }
                catch (Exception ex) { Console.WriteLine($"LoftedProtrusion Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { loftedProtrusionElements.Add(new XElement("type", loftedProtrusion.Type)); }
                catch (Exception ex) { Console.WriteLine($"LoftedProtrusion Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { loftedProtrusionElements.Add(new XElement("modelingModeType", loftedProtrusion.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"LoftedProtrusion ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- Cross-sections ----
                try
                {
                    Array crossSections = Array.CreateInstance(typeof(object), 0);
                    loftedProtrusion.GetCrossSections(out int numCrossSections, ref crossSections);

                    XElement crossSectionsElement = new XElement("CrossSections", new XAttribute("Count", numCrossSections));
                    foreach (object csObj in crossSections)
                    {
                        if (csObj is Profile csProfile)
                        {
                            crossSectionsElement.Add(GE04_getProfiles_extractor.Profile_Data(csProfile));
                        }
                    }
                    loftedProtrusionElements.Add(crossSectionsElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetCrossSections: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    loftedProtrusion.GetDimensions(out int numDims, ref dims);
                    loftedProtrusionElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    loftedProtrusion.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    loftedProtrusionElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = loftedProtrusion.Edges[edgeTyp];

                    XElement edgeElements = new XElement("edges", new XAttribute("count", edges.Count));

                    for (int e = 1; e <= edges.Count; e++)
                    {
                        var edge = (Edge)edges.Item(e);
                        edgeElements.Add(new XElement($"type{e}", edge.Type.ToString()));

                        edge.GetEndPoints(ref startPoint, ref endPoint);
                        edgeElements.Add(new XElement($"endPoints{e}",
                            new XAttribute("startpoint", string.Join(" ", (double[])startPoint)),
                            new XAttribute("endPoint", string.Join(" ", (double[])endPoint))));
                    }

                    loftedProtrusionElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = loftedProtrusion.Faces[faceTyp];
                    loftedProtrusionElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynLoftedProtrusion = loftedProtrusion;
                    loftedProtrusionElements.Add(new XElement("status", dynLoftedProtrusion.Status));
                    loftedProtrusionElements.Add(new XElement("suppress", dynLoftedProtrusion.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = loftedProtrusion.GetStatusEx(out object description);
                    loftedProtrusionElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedProtrusion GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LoftedProtrusion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("LoftedProtrusion", "Error");
            }
            finally
            {
                if (loftedProtrusion != null)
                {
                    Marshal.ReleaseComObject(loftedProtrusion);
                    loftedProtrusion = null;
                }
            }

            Console.WriteLine("Created Lofted Protrusion XML list");
            return loftedProtrusionElements;
        }

        public static XElement Lofted_Cutout(LoftedCutout loftedCutout)
        {
            XElement loftedCutoutElements = new XElement("LoftedCutout", new XAttribute("Type", 2057842149));

            try
            {
                try { loftedCutoutElements.Add(new XElement("name", loftedCutout.Name)); }
                catch (Exception ex) { Console.WriteLine($"LoftedCutout Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { loftedCutoutElements.Add(new XElement("type", loftedCutout.Type)); }
                catch (Exception ex) { Console.WriteLine($"LoftedCutout Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { loftedCutoutElements.Add(new XElement("modelingModeType", loftedCutout.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"LoftedCutout ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- Cross-sections ----
                try
                {
                    Array crossSections = Array.CreateInstance(typeof(object), 0);
                    loftedCutout.GetCrossSections(out int numCrossSections, ref crossSections);

                    XElement crossSectionsElement = new XElement("CrossSections", new XAttribute("Count", numCrossSections));
                    foreach (object csObj in crossSections)
                    {
                        if (csObj is Profile csProfile)
                        {
                            crossSectionsElement.Add(GE04_getProfiles_extractor.Profile_Data(csProfile));
                        }
                    }
                    loftedCutoutElements.Add(crossSectionsElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetCrossSections: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    loftedCutout.GetDimensions(out int numDims, ref dims);
                    loftedCutoutElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    loftedCutout.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    loftedCutoutElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = loftedCutout.Edges[edgeTyp];

                    XElement edgeElements = new XElement("edges", new XAttribute("count", edges.Count));

                    for (int e = 1; e <= edges.Count; e++)
                    {
                        var edge = (Edge)edges.Item(e);
                        edgeElements.Add(new XElement($"type{e}", edge.Type.ToString()));

                        edge.GetEndPoints(ref startPoint, ref endPoint);
                        edgeElements.Add(new XElement($"endPoints{e}",
                            new XAttribute("startpoint", string.Join(" ", (double[])startPoint)),
                            new XAttribute("endPoint", string.Join(" ", (double[])endPoint))));
                    }

                    loftedCutoutElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = loftedCutout.Faces[faceTyp];
                    loftedCutoutElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynLoftedCutout = loftedCutout;
                    loftedCutoutElements.Add(new XElement("status", dynLoftedCutout.Status));
                    loftedCutoutElements.Add(new XElement("suppress", dynLoftedCutout.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = loftedCutout.GetStatusEx(out object description);
                    loftedCutoutElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoftedCutout GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LoftedCutout: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("LoftedCutout", "Error");
            }
            finally
            {
                if (loftedCutout != null)
                {
                    Marshal.ReleaseComObject(loftedCutout);
                    loftedCutout = null;
                }
            }

            Console.WriteLine("Created Lofted Cutout XML list");
            return loftedCutoutElements;
        }
    }
}