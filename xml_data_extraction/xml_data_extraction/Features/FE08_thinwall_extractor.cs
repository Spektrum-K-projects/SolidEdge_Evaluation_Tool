using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE08_thinwall_extractor
    {
        public static XElement ThinWall(ThinWall thinwall)
        {
            XElement thinwallElements = new XElement("Thin_Wall", new XAttribute("Type", 462094734));

            try
            {
                try { thinwallElements.Add(new XElement("name", thinwall.Name)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinwallElements.Add(new XElement("type", thinwall.Type)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinwallElements.Add(new XElement("modelingModeType", thinwall.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinwallElements.Add(new XElement("showDimensions", thinwall.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinwallElements.Add(new XElement("visible", thinwall.Visible)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinwallElements.Add(new XElement("thickness", thinwall.Thickness)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall Thickness: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinwallElements.Add(new XElement("thicknessSide", thinwall.ThicknessSide)); }
                catch (Exception ex) { Console.WriteLine($"ThinWall ThicknessSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynThinwall = thinwall;
                    thinwallElements.Add(new XElement("status", dynThinwall.Status));
                    thinwallElements.Add(new XElement("suppress", dynThinwall.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = thinwall.GetStatusEx(out object description);
                    thinwallElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    thinwall.GetDimensions(out int numDims, ref dims);
                    thinwallElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    thinwall.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    thinwallElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = thinwall.Edges[edgeTyp];

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

                    thinwallElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = thinwall.Faces[faceTyp];
                    thinwallElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Unique thicknesses ----
                try
                {
                    Array uniqueThicknesses = Array.CreateInstance(typeof(double), 0);
                    thinwall.GetUniqueThicknesses(out int uniqueCount, ref uniqueThicknesses);

                    var uniqueThicknessesElement = new XElement("UniqueThicknesses", new XAttribute("Count", uniqueCount));
                    for (int i = 0; i < uniqueThicknesses.Length; i++)
                    {
                        uniqueThicknessesElement.Add(new XElement("thickness", Convert.ToDouble(uniqueThicknesses.GetValue(i))));
                    }
                    thinwallElements.Add(uniqueThicknessesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetUniqueThicknesses: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ThinWall: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (thinwall != null)
                {
                    Marshal.ReleaseComObject(thinwall);
                    thinwall = null;
                }
            }

            Console.WriteLine("\t Created Thin Wall XML list");
            return thinwallElements;
        }
    }
}