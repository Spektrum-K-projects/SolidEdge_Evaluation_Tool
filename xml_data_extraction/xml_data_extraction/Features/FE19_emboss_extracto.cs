using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    //Emboss Feature Data Extraction
    internal class FE19_emboss_extractor
    {
        public static XElement Emboss_Extract(EmbossFeature embossFeature)
        {
            XElement embossElements = new XElement("Emboss", new XAttribute("Type", -2101998503));

            try
            {
                try { embossElements.Add(new XElement("name", embossFeature.Name)); }
                catch (Exception ex) { Console.WriteLine($"Emboss Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("type", embossFeature.Type)); }
                catch (Exception ex) { Console.WriteLine($"Emboss Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("modelingModeType", embossFeature.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Emboss ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("clearance", embossFeature.Clearance)); }
                catch (Exception ex) { Console.WriteLine($"Emboss Clearance: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("thicken", embossFeature.Thicken)); }
                catch (Exception ex) { Console.WriteLine($"Emboss Thicken: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("thickness", embossFeature.Thickness)); }
                catch (Exception ex) { Console.WriteLine($"Emboss Thickness: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("side", embossFeature.Side)); }
                catch (Exception ex) { Console.WriteLine($"Emboss Side: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("dieRounding", embossFeature.DieRounding.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Emboss DieRounding: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("punchSideRounding", embossFeature.PunchSideRounding.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Emboss PunchSideRounding: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("dieRadius", embossFeature.DieRadius)); }
                catch (Exception ex) { Console.WriteLine($"Emboss DieRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { embossElements.Add(new XElement("punchSideRadius", embossFeature.PunchSideRadius)); }
                catch (Exception ex) { Console.WriteLine($"Emboss PunchSideRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    object embossTarget = embossFeature.EmbossTarget;
                    embossElements.Add(GetBodyData(embossTarget, "EmbossTarget"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetEmbossTarget: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Emboss tool bodies ----
                try
                {
                    int numTools = 0;
                    Array toolBodies = Array.CreateInstance(typeof(object), 0);
                    embossFeature.GetEmbossToolBodies(ref numTools, ref toolBodies);

                    XElement toolBodiesElement = new XElement("EmbossToolBodies", new XAttribute("Count", numTools));
                    int tbIndex = 0;
                    foreach (object tbObj in toolBodies)
                    {
                        tbIndex++;
                        toolBodiesElement.Add(GetBodyData(tbObj, $"body{tbIndex}"));
                    }
                    embossElements.Add(toolBodiesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetEmbossToolBodies: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    embossFeature.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    embossElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    embossFeature.ExactRange(out double ex1, out double ey1, out double ez1, out double ex2, out double ey2, out double ez2);
                    embossElements.Add(new XElement("ExactRange",
                        new XAttribute("X1", ex1), new XAttribute("Y1", ey1), new XAttribute("Z1", ez1),
                        new XAttribute("X2", ex2), new XAttribute("Y2", ey2), new XAttribute("Z2", ez2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetExactRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = embossFeature.Edges[edgeTyp];

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

                    embossElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = embossFeature.Faces[faceTyp];
                    embossElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynEmboss = embossFeature;
                    embossElements.Add(new XElement("status", dynEmboss.Status));
                    embossElements.Add(new XElement("suppress", dynEmboss.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = embossFeature.GetStatusEx(out object description);
                    embossElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Emboss GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Emboss: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Emboss", "Error");
            }
            finally
            {
                if (embossFeature != null)
                {
                    Marshal.ReleaseComObject(embossFeature);
                    embossFeature = null;
                }
            }

            Console.WriteLine("Created Emboss XML list");
            return embossElements;
        }

        private static XElement GetBodyData(object bodyObj, string elementName)
        {
            var el = new XElement(elementName);
            try
            {
                if (bodyObj is SolidEdgeGeometry.Body body)
                {
                    try { el.Add(new XElement("displayName", body.DisplayName)); } catch { }
                    try { el.Add(new XElement("isSolid", body.IsSolid)); } catch { }
                    try { el.Add(new XElement("volume", body.Volume)); } catch { }

                    try
                    {
                        Array minPoint = Array.CreateInstance(typeof(double), 0);
                        Array maxPoint = Array.CreateInstance(typeof(double), 0);
                        body.GetRange(ref minPoint, ref maxPoint);

                        double[] minArr = (double[])minPoint;
                        double[] maxArr = (double[])maxPoint;

                        el.Add(new XElement("Range",
                            new XAttribute("X1", minArr[0]), new XAttribute("Y1", minArr[1]), new XAttribute("Z1", minArr[2]),
                            new XAttribute("X2", maxArr[0]), new XAttribute("Y2", maxArr[1]), new XAttribute("Z2", maxArr[2])));
                    }
                    catch { }
                }
                else
                {
                    el.Add(new XAttribute("UnderlyingType", bodyObj?.GetType().Name ?? "null"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetBodyData ({elementName}): {ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            return el;
        }
    }
}