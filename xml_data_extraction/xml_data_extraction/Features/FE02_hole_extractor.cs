using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using xml_data_extraction.Miscellaneous_Methods;
using xml_data_extraction.Properties;

namespace xml_data_extraction.Features
{
    internal class FE02_hole_extractor
    {
        public static XElement Hole(Hole hole)
        {
            XElement holeElements = new XElement("Hole", new XAttribute("Type", 462094722));

            try
            {
                try { holeElements.Add(new XElement("name", hole.Name)); }
                catch (Exception ex) { Console.WriteLine($"Hole Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("type", hole.Type)); }
                catch (Exception ex) { Console.WriteLine($"Hole Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("modelingModeType", hole.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Hole ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("showDimensions", hole.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Hole ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("visible", hole.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Hole Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("extent_side", hole.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Hole ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("extent_type", hole.ExtentType)); }
                catch (Exception ex) { Console.WriteLine($"Hole ExtentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("depth", hole.Depth)); }
                catch (Exception ex) { Console.WriteLine($"Hole Depth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("draftAngle", hole.DraftAngle)); }
                catch (Exception ex) { Console.WriteLine($"Hole DraftAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(new XElement("createPhysicalThread", hole.CreatePhysicalThread)); }
                catch (Exception ex) { Console.WriteLine($"Hole CreatePhysicalThread: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- FromPlane / ToPlane ----
                try { holeElements.Add(MM01_geometry_methods.GetPlaneData(hole.FromPlane, "FromPlane")); }
                catch (Exception ex) { Console.WriteLine($"Hole GetFromPlane: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(MM01_geometry_methods.GetPlaneData(hole.ToPlane, "ToPlane")); }
                catch (Exception ex) { Console.WriteLine($"Hole GetToPlane: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- TopCap / BottomCap / SideFaces ----
                try { holeElements.Add(MM01_geometry_methods.GetPlaneData(hole.TopCap, "TopCap")); }
                catch (Exception ex) { Console.WriteLine($"Hole GetTopCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(MM01_geometry_methods.GetPlaneData(hole.BottomCap, "BottomCap")); }
                catch (Exception ex) { Console.WriteLine($"Hole GetBottomCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeElements.Add(MM01_geometry_methods.GetPlaneData(hole.SideFaces, "SideFaces")); }
                catch (Exception ex) { Console.WriteLine($"Hole GetSideFaces: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- Profiles ----
                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(hole);
                    holeElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- HoleData ----
                try
                {
                    holeElements.Add(PR03_hole_data_extractor.Hole_Data(hole));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetHoleData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    hole.GetDimensions(out int numDims, ref dims);
                    holeElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    hole.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    holeElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = hole.Edges[edgeTyp];

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

                    holeElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = hole.Faces[faceTyp];
                    holeElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Body array  ----
                try
                {
                    Array bodyArray = Array.CreateInstance(typeof(object), 0);
                    hole.GetBodyArray(out bool multiBodyCut, out int numberOfBodies, out bodyArray);
                    holeElements.Add(new XElement("BodyArray",
                        new XAttribute("MultiBodyCut", multiBodyCut),
                        new XAttribute("NumberOfBodies", numberOfBodies)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetBodyArray: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynHole = hole;
                    holeElements.Add(new XElement("status", dynHole.Status));
                    holeElements.Add(new XElement("suppress", dynHole.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = hole.GetStatusEx(out object description);
                    holeElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hole GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hole: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (hole != null)
                {
                    Marshal.ReleaseComObject(hole);
                    hole = null;
                }
            }

            Console.WriteLine($"\t Created Hole XML list");
            return holeElements;
        }
    }
}