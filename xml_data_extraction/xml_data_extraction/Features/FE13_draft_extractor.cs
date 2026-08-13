using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    //Draft Data Extraction
    internal class FE13_draft_extractor
    {
        public static XElement draft_Data(Draft draft)
        {
            XElement draftElements = new XElement("Draft", new XAttribute("Type", 462094746));

            try
            {
                try { draftElements.Add(new XElement("name", draft.Name)); }
                catch (Exception ex) { Console.WriteLine($"Draft Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { draftElements.Add(new XElement("type", draft.Type)); }
                catch (Exception ex) { Console.WriteLine($"Draft Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { draftElements.Add(new XElement("modelingModeType", draft.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Draft ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { draftElements.Add(new XElement("showDimensions", draft.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Draft ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { draftElements.Add(new XElement("visible", draft.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Draft Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { draftElements.Add(new XElement("draftSide", draft.DraftSide.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Draft DraftSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynDraft = draft;
                    draftElements.Add(new XElement("status", dynDraft.Status));
                    draftElements.Add(new XElement("suppress", dynDraft.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array draftAngles = Array.CreateInstance(typeof(double), 0);
                    draft.GetDraftAngles(out int draftAngleCount, ref draftAngles);

                    var draftAnglesElement = new XElement("DraftAngles", new XAttribute("Count", draftAngleCount));
                    for (int i = 0; i < draftAngles.Length; i++)
                    {
                        draftAnglesElement.Add(new XElement("angle", Convert.ToDouble(draftAngles.GetValue(i))));
                    }
                    draftElements.Add(draftAnglesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetDraftAngles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    draft.GetDimensions(out int numDims, ref dims);
                    draftElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    draft.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    draftElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = draft.Edges[edgeTyp];

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

                    draftElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = draft.Faces[faceTyp];
                    draftElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    draftElements.Add(MM01_geometry_methods.GetPlaneData(draft.DraftPlane, "DraftPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetDraftPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = draft.GetStatusEx(out object description);
                    draftElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Draft: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (draft != null)
                {
                    Marshal.ReleaseComObject(draft);
                    draft = null;
                }
            }

            Console.WriteLine("\t Created Draft Feature XML list");
            return draftElements;
        }
    }
}