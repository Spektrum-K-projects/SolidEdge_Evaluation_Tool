using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    //Draft Data Extraction
    internal class FE13_draft_extractor
    {
        public static XElement draft_Data(Draft draft)
        {
            XElement draftElements = new XElement("Draft", new XAttribute("Type", 462094746));

            Array draftAngles = Array.CreateInstance(typeof(double), 0);

            try
            {
                draftElements.Add(new XElement("name", draft.Name));
                draftElements.Add(new XElement("type", draft.Type));
                draftElements.Add(new XElement("modelingModeType", draft.ModelingModeType));
                draftElements.Add(new XElement("draftSide", draft.DraftSide.ToString()));

                try
                {
                    dynamic dynDraft = draft;
                    draftElements.Add(new XElement("status", dynDraft.Status));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetStatus: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                draft.GetDraftAngles(out int draftAngleCount, ref draftAngles);
                if (draftAngles.Length > 0)
                {
                    draftElements.Add(new XElement("draftAngleCount", draftAngleCount));
                    draftElements.Add(new XElement("draftAngles",
                        new XAttribute("values", string.Join(" ", (double[])draftAngles))));
                }

                // ---- Dimensions (count only - matches how Pattern's GetDimensions was handled) ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    draft.GetDimensions(out int numDims, ref dims);
                    draftElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Whole-feature bounding box ----
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

                // ---- Edges (mirrors the already-proven pattern from Round, FE04) ----
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
                    }

                    draftElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces (count only - per-face detail not yet verified against the Face interop) ----
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

                // ---- Suppress (whole-feature suppress flag) ----
                try
                {
                    dynamic dynDraft = draft;
                    draftElements.Add(new XElement("suppress", dynDraft.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Draft GetSuppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx (status code + description text, when Solid Edge provides one) ----
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
                Console.WriteLine($"Draft: Error Message:{ex.Message}");
                return new XElement("Draft", "Error");
            }
            finally
            {
                if (draft != null)
                {
                    Marshal.ReleaseComObject(draft);
                    draft = null;
                }
            }

            Console.WriteLine("Created Draft Feature XML list");
            return draftElements;
        }
    }
}