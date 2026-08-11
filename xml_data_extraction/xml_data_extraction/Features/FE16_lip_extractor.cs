using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    //Lip Feature Data Extraction
    internal class FE16_lip_extractor
    {
        public static XElement Lip_Extract(Lip lip)
        {
            XElement lipElements = new XElement("Lip", new XAttribute("Type", -483645223));

            try
            {
                try { lipElements.Add(new XElement("name", lip.Name)); }
                catch (Exception ex) { Console.WriteLine($"Lip Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { lipElements.Add(new XElement("type", lip.Type)); }
                catch (Exception ex) { Console.WriteLine($"Lip Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { lipElements.Add(new XElement("modelingModeType", lip.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Lip ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { lipElements.Add(new XElement("width", lip.Width)); }
                catch (Exception ex) { Console.WriteLine($"Lip Width: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { lipElements.Add(new XElement("height", lip.Height)); }
                catch (Exception ex) { Console.WriteLine($"Lip Height: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- SideFace / CapFace ----
                try
                {
                    object sideFace = lip.SideFace;
                    lipElements.Add(new XElement("SideFace",
                        new XAttribute("UnderlyingType", sideFace?.GetType().Name ?? "null")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetSideFace: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    object capFace = lip.CapFace;
                    lipElements.Add(new XElement("CapFace",
                        new XAttribute("UnderlyingType", capFace?.GetType().Name ?? "null")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetCapFace: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    lip.GetDimensions(out int numDims, ref dims);
                    lipElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    lip.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    lipElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = lip.Edges[edgeTyp];

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

                    lipElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = lip.Faces[faceTyp];
                    lipElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Suppress ----
                try
                {
                    dynamic dynLip = lip;
                    lipElements.Add(new XElement("suppress", dynLip.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetSuppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = lip.GetStatusEx(out object description);
                    lipElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lip GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lip: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Lip", "Error");
            }
            finally
            {
                if (lip != null)
                {
                    Marshal.ReleaseComObject(lip);
                    lip = null;
                }
            }

            Console.WriteLine("\t Created Lip Feature XML list");
            return lipElements;
        }
    }
}