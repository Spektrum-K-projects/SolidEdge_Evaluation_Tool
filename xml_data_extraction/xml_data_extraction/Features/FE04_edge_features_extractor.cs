using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    internal class FE04_edge_features_extractor
    {
        private static void TryAdd(XElement parent, string featureLabel, string propertyLabel, Func<object> getter)
        {
            try
            {
                parent.Add(new XElement(propertyLabel, getter()));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{featureLabel} {propertyLabel}: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }
        }

        //Chamfer Data Extraction
        public static XElement Chamfer(Chamfer chamfer)
        {
            XElement chamferElements = new XElement("Chamfer", new XAttribute("Type", 462094742));

            try
            {
                TryAdd(chamferElements, "Chamfer", "Name", () => chamfer.Name);
                TryAdd(chamferElements, "Chamfer", "modelingModeType", () => chamfer.ModelingModeType);
                TryAdd(chamferElements, "Chamfer", "showDimensions", () => chamfer.ShowDimensions);
                TryAdd(chamferElements, "Chamfer", "visible", () => chamfer.Visible);

                FeaturePropertyConstants chamferType = FeaturePropertyConstants.igChamfer45degSetback;
                bool chamferTypeKnown = false;

                try
                {
                    chamferType = chamfer.ChamferType;
                    chamferTypeKnown = true;
                    chamferElements.Add(new XElement("type", chamferType));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chamfer ChamferType: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                if (!chamferTypeKnown || chamferType != FeaturePropertyConstants.igChamfer2Setbacks)
                {
                    TryAdd(chamferElements, "Chamfer", "setback_angle", () => chamfer.ChamferSetbackAngle);
                }
                else
                {
                    chamferElements.Add(new XElement("setback_angle", "not_applicable"));
                }

                TryAdd(chamferElements, "Chamfer", "setback_value_1", () => chamfer.ChamferSetbackValue1);

                if (!chamferTypeKnown || chamferType != FeaturePropertyConstants.igChamfer45degSetback)
                {
                    TryAdd(chamferElements, "Chamfer", "setback_value_2", () => chamfer.ChamferSetbackValue2);
                }
                else
                {
                    chamferElements.Add(new XElement("setback_value_2", "not_applicable"));
                }

                // ---- Reference face ----
                try
                {
                    chamferElements.Add(MM01_geometry_methods.GetPlaneData(chamfer.ChamferReferenceFace, "ChamferReferenceFace"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chamfer GetChamferReferenceFace: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynChamfer = chamfer;
                    chamferElements.Add(new XElement("status", dynChamfer.Status));
                    chamferElements.Add(new XElement("suppress", dynChamfer.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chamfer GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = chamfer.GetStatusEx(out object description);
                    chamferElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chamfer GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    chamfer.GetDimensions(out int numDims, ref dims);
                    chamferElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chamfer GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    chamfer.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    chamferElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chamfer GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chamfer: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (chamfer != null)
                {
                    Marshal.ReleaseComObject(chamfer);
                    chamfer = null;
                }
            }

            Console.WriteLine($"\t Created Chamfer Feature XML list");
            return chamferElements;
        }

        //Round Data Extraction
        public static XElement Round(Round round)
        {
            XElement roundElements = new XElement("Round", new XAttribute("Type", 462094738));

            try
            {
                try { roundElements.Add(new XElement("type", round.Type)); }
                catch (Exception ex) { Console.WriteLine($"Round Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { roundElements.Add(new XElement("name", round.Name)); }
                catch (Exception ex) { Console.WriteLine($"Round Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { roundElements.Add(new XElement("modelingModeType", round.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Round ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { roundElements.Add(new XElement("showDimensions", round.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Round ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { roundElements.Add(new XElement("visible", round.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Round Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- BlendShape ----
                try { roundElements.Add(new XElement("blendShape", round.BlendShape.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Round BlendShape: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { roundElements.Add(new XElement("blendShapeValue", round.BlendShapeValue)); }
                catch (Exception ex) { Console.WriteLine($"Round BlendShapeValue: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                int edgeSetCount = 0;
                try
                {
                    edgeSetCount = round.EdgeSetCount;
                    roundElements.Add(new XElement("edge_set_count", edgeSetCount));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round EdgeSetCount: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Per-edge-set radius data ----
                for (int idx = 1; idx <= edgeSetCount; idx++)
                {
                    try
                    {
                        SolidEdgePart.RoundTypeConstants radiusType = round.RadiusType[idx];

                        roundElements.Add(new XElement($"radius_type{idx}", radiusType));

                        if (radiusType == SolidEdgePart.RoundTypeConstants.igConstantRadius)
                        {
                            Array constant_Radii = Array.CreateInstance(typeof(double), 0);
                            round.GetConstantRadii(out int constant_RadiiCount, ref constant_Radii);

                            if (constant_Radii.Length > 0)
                            {
                                var constantRadiiElement = new XElement($"constantRadii{idx}", new XAttribute("Count", constant_RadiiCount));
                                foreach (var radius in constant_Radii)
                                {
                                    constantRadiiElement.Add(new XElement("radius", radius));
                                }
                                roundElements.Add(constantRadiiElement);
                            }
                        }
                        else if (radiusType == SolidEdgePart.RoundTypeConstants.igVariableRadius)
                        {
                            Array variable_Radii = Array.CreateInstance(typeof(double), 0);
                            round.GetVariableRadii(out int variable_RadiiCount, ref variable_Radii);

                            if (variable_Radii.Length > 0)
                            {
                                var variableRadiiElement = new XElement($"variableRadii{idx}", new XAttribute("Count", variable_RadiiCount));
                                foreach (var radius in variable_Radii)
                                {
                                    variableRadiiElement.Add(new XElement("radius", radius));
                                }
                                roundElements.Add(variableRadiiElement);
                            }
                        }
                        else
                        {
                            roundElements.Add(new XElement($"radiusNote{idx}", "No constant or variable radius found for this edge set"));
                        }
                    }
                    catch (COMException ex) when ((uint)ex.ErrorCode == 0x8002000B)
                    {
                        Console.WriteLine($"Round EdgeSetIndex {idx}: Index out of bounds. No more radius types for this Round feature.");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Round EdgeSetIndex {idx}: {ex.Message} | Inner: {ex.InnerException?.Message}");
                    }
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    round.GetDimensions(out int numDims, ref dims);
                    roundElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    round.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    roundElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = round.Edges[edgeTyp];

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

                    roundElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = round.Faces[faceTyp];
                    roundElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynRound = round;
                    roundElements.Add(new XElement("status", dynRound.Status));
                    roundElements.Add(new XElement("suppress", dynRound.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = round.GetStatusEx(out object description);
                    roundElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Round GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Round: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (round != null)
                {
                    Marshal.ReleaseComObject(round);
                    round = null;
                }
            }

            Console.WriteLine($"\t Created Round Feature XML list");
            return roundElements;
        }
    }
}