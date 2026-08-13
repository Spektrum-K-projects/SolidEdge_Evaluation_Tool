using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using xml_data_extraction.Properties;

namespace xml_data_extraction.Features
{
    //Thread Feature Data Extraction
    internal class FE20_thread_extractor
    {
        public static XElement Thread_Extract(SolidEdgePart.Thread thread)
        {
            XElement threadElements = new XElement("Thread");

            try
            {
                try { threadElements.Add(new XElement("name", thread.Name)); }
                catch (Exception ex) { Console.WriteLine($"Thread Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { threadElements.Add(new XElement("type", thread.Type)); }
                catch (Exception ex) { Console.WriteLine($"Thread Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { threadElements.Add(new XElement("modelingModeType", thread.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Thread ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { threadElements.Add(new XElement("createPhysicalThread", thread.CreatePhysicalThread)); }
                catch (Exception ex) { Console.WriteLine($"Thread CreatePhysicalThread: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { threadElements.Add(new XElement("showDimensions", thread.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Thread ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { threadElements.Add(new XElement("visible", thread.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Thread Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    object holeDataObj = thread.HoleData;
                    if (holeDataObj is HoleData holeData)
                    {
                        threadElements.Add(PR03_hole_data_extractor.HoleData_Extract(holeData));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetHoleData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    thread.GetDimensions(out int numDims, ref dims);
                    threadElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    thread.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    threadElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = thread.Edges[edgeTyp];

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

                    threadElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = thread.Faces[faceTyp];
                    threadElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynThread = thread;
                    threadElements.Add(new XElement("status", dynThread.Status));
                    threadElements.Add(new XElement("suppress", dynThread.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = thread.GetStatusEx(out object description);
                    threadElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Thread GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Thread: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Thread", "Error");
            }
            finally
            {
                if (thread != null)
                {
                    Marshal.ReleaseComObject(thread);
                    thread = null;
                }
            }

            Console.WriteLine("Created Thread XML list");
            return threadElements;
        }
    }
}