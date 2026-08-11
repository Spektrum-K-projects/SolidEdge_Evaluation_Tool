using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE06_rib_extractor
    {
        public static XElement Rib(Rib rib)
        {
            XElement ribElements = new XElement("Rib", new XAttribute("Type", 462094730));

            try
            {
                try { ribElements.Add(new XElement("name", rib.Name)); }
                catch (Exception ex) { Console.WriteLine($"Rib Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("type", rib.Type)); }
                catch (Exception ex) { Console.WriteLine($"Rib Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("modelingModeType", rib.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Rib ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("showDimensions", rib.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Rib ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("visible", rib.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Rib Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("thickness", rib.Thickness)); }
                catch (Exception ex) { Console.WriteLine($"Rib Thickness: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("thicknessSide", rib.ThicknessSide)); }
                catch (Exception ex) { Console.WriteLine($"Rib ThicknessSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("thicknessType", rib.ThicknessType)); }
                catch (Exception ex) { Console.WriteLine($"Rib ThicknessType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("profileExtensionType", rib.ProfileExtensionType)); }
                catch (Exception ex) { Console.WriteLine($"Rib ProfileExtensionType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ribElements.Add(new XElement("materialSide", rib.MaterialSide)); }
                catch (Exception ex) { Console.WriteLine($"Rib MaterialSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- Profiles ----
                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(rib);
                    ribElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    rib.GetDimensions(out int numDims, ref dims);
                    ribElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    rib.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    ribElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = rib.Edges[edgeTyp];

                    XElement edgeElements = new XElement("edges", new XAttribute("count", edges.Count));

                    for (int e = 1; e <= edges.Count; e++)
                    {
                        var edge = (Edge)edges.Item(e);
                        edgeElements.Add(new XElement($"type{e}", edge.Type.ToString()));

                        edge.GetEndPoints(ref startPoint, ref endPoint);
                        edgeElements.Add(new XElement($"endPoints{e}",
                            new XAttribute("startpoint", string.Join(" ", (double[])startPoint)),
                            new XAttribute("endPoint", string.Join(" ", (double[])endPoint))));

                        Marshal.ReleaseComObject(edge);
                    }

                    ribElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = rib.Faces[faceTyp];
                    ribElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynRib = rib;
                    ribElements.Add(new XElement("status", dynRib.Status));
                    ribElements.Add(new XElement("suppress", dynRib.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = rib.GetStatusEx(out object description);
                    ribElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Rib GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Rib: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (rib != null)
                {
                    Marshal.ReleaseComObject(rib);
                    rib = null;
                }
            }

            Console.WriteLine($"\t Created Rib XML list");
            return ribElements;
        }
    }
}