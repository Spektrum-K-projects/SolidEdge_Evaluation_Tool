using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE09_thinregion_extractor
    {
        public static XElement ThinRegion(Thin thinRegion)
        {
            XElement thinRegionElements = new XElement("Thin_Region", new XAttribute("Type", 438630050));

            try
            {
                try { thinRegionElements.Add(new XElement("name", thinRegion.Name)); }
                catch (Exception ex) { Console.WriteLine($"ThinRegion Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinRegionElements.Add(new XElement("type", thinRegion.Type)); }
                catch (Exception ex) { Console.WriteLine($"ThinRegion Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinRegionElements.Add(new XElement("modelingModeType", thinRegion.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"ThinRegion ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinRegionElements.Add(new XElement("showDimensions", thinRegion.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"ThinRegion ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { thinRegionElements.Add(new XElement("visible", thinRegion.Visible)); }
                catch (Exception ex) { Console.WriteLine($"ThinRegion Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynThinRegion = thinRegion;
                    thinRegionElements.Add(new XElement("status", dynThinRegion.Status));
                    thinRegionElements.Add(new XElement("suppress", dynThinRegion.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = thinRegion.GetStatusEx(out object description);
                    thinRegionElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    thinRegion.GetDimensions(out int numDims, ref dims);
                    thinRegionElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    thinRegion.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    thinRegionElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = thinRegion.Edges[edgeTyp];

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

                    thinRegionElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = thinRegion.Faces[faceTyp];
                    thinRegionElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ThinRegion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (thinRegion != null)
                {
                    Marshal.ReleaseComObject(thinRegion);
                    thinRegion = null;
                }
            }

            Console.WriteLine("Created Thin Region XML list");
            return thinRegionElements;
        }
    }
}