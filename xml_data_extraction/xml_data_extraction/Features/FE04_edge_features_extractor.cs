using SolidEdgeGeometry;
using SolidEdgePart;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE04_edge_features_extractor
    {
        //Chamfer Data Extraction
        private static void TryAdd(XElement parent, string label, Func<object> getter)
        {
            try
            {
                parent.Add(new XElement(label, getter()));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chamfer {label}: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }
        }

        //Chamfer Data Extraction
        public static XElement Chamfer(Chamfer chamfer)
        {
            XElement chamferElements = new XElement("Chamfer", new XAttribute("Type", 462094742));

            try
            {
                TryAdd(chamferElements, "Name", () => chamfer.Name);
                TryAdd(chamferElements, "modelingModeType", () => chamfer.ModelingModeType);

                FeaturePropertyConstants chamferType = chamfer.ChamferType;
                chamferElements.Add(new XElement("type", chamferType));

                if (chamferType != FeaturePropertyConstants.igChamfer2Setbacks)
                {
                    TryAdd(chamferElements, "setback_angle", () => chamfer.ChamferSetbackAngle);
                }
                else
                {
                    chamferElements.Add(new XElement("setback_angle", "not_applicable"));
                }

                TryAdd(chamferElements, "setback_value_1", () => chamfer.ChamferSetbackValue1);

                if (chamferType != FeaturePropertyConstants.igChamfer45degSetback)
                {
                    TryAdd(chamferElements, "setback_value_2", () => chamfer.ChamferSetbackValue2);
                }
                else
                {
                    chamferElements.Add(new XElement("setback_value_2", "not_applicable"));
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
                return new XElement("Chamfer", "Error");
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
        public static XElement Round (Round round)
        {
            XElement roundElements = new XElement("Round", new XAttribute("Type", 462094738));

            Array constant_Radii = Array.CreateInstance(typeof(double), 0);
            Array variable_Radii = Array.CreateInstance(typeof(double), 0);

            Array startPoint = Array.CreateInstance(typeof(double), 0);
            Array endPoint = Array.CreateInstance(typeof(double), 0);

            try
            {
                var round_type = round.Type;
                roundElements.Add(new XElement("type", round_type));
                //Console.WriteLine($"Round Type: {round_type}");

                var round_Name = round.Name;
                roundElements.Add(new XElement("name", round_Name));
                //Console.WriteLine($"Round Name: {round_Name}");

                var round_EdgeSetCount = round.EdgeSetCount;
                roundElements.Add(new XElement("edge_set_count", round_EdgeSetCount));
                //Console.WriteLine($"Round Edge Set Count: {round_EdgeSetCount}");

                //SolidEdgePart.BlendShapeConstants round_BlendShape = round.BlendShape;
                ////roundElements.Add(new XElement("blend_shape", round_BlendShape));
                //Console.WriteLine($"Round Blend Shape: {round_BlendShape}");

                //var round_BlendShapeValue = round.BlendShapeValue;
                ////roundElements.Add(new XElement("blend_shape_value", round_BlendShapeValue));
                //Console.WriteLine($"Round Blend Shape Value: {round_BlendShapeValue.ToString()}");

                //Extracting the Radius information for the Round
                for (int idx = 1; idx <= round_EdgeSetCount; idx++)
                {
                    try
                    {
                        SolidEdgePart.RoundTypeConstants radiusType = round.RadiusType[idx];
                        //Console.WriteLine($"Radius Type{idx}: {radiusType}");

                        roundElements.Add(new XElement($"radius_index{idx}", idx));
                        roundElements.Add(new XElement($"radius_type{idx}", radiusType));

                        if (radiusType.ToString() == "igConstantRadius")
                        {
                            round.GetConstantRadii(out int constant_RadiiCount, ref constant_Radii);

                            if (constant_Radii.Length > 0)
                            {
                                roundElements.Add(new XElement("constantradii_count", constant_RadiiCount));

                                //Console.WriteLine("  Constant Radii found:");
                                foreach (var radius in constant_Radii)
                                {
                                    //Console.WriteLine($"    Radius: {radius.ToString()} mm");
                                    roundElements.Add(new XElement("radius", radius.ToString()));
                                }
                            }
                        }

                        else if (radiusType.ToString() == "igVariableRadius")
                        {
                            round.GetVariableRadii(out int variable_RadiiCount, ref variable_Radii);

                            if (variable_Radii.Length > 0)
                            {
                                roundElements.Add(new XElement("variableradii_count", variable_RadiiCount));

                                //Console.WriteLine("  Varaiable Radii found:");
                                foreach (var radius in variable_Radii)
                                {
                                    //Console.WriteLine($"    Radius: {radius.ToString()} mm");
                                    roundElements.Add(new XElement("variable_radii", radius.ToString()));
                                }
                            }
                        }

                        else
                        {
                            Console.WriteLine("No Radius Found");
                        }
                    }

                    catch (COMException ex) when ((uint)ex.ErrorCode == 0x8002000B)
                    {
                        Console.WriteLine($"EdgeSetIndex {idx}: Index out of bounds. No more radius types for this Round feature.");
                        break;
                    }
                }

                //Edges End Point Coordinates Extraction
                FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                var edges = round.Edges[EdgeTyp];
                Console.WriteLine($"Round.Edges []: {edges.Count}");
                XElement edgeElements = new XElement("edges");
                edgeElements.Add(new XAttribute("count", edges.Count));

                for (int e = 1; e <= edges.Count; e++)
                {
                    var edge = (Edge)edges.Item(e);
                    var edgeElement = edge.Type.ToString();
                    edgeElements.Add(new XElement($"type{e}", edgeElement));

                    edge.GetEndPoints(ref startPoint, ref endPoint);

                    edgeElements.Add(new XElement($"endPoints{e}", new XAttribute("startpoint", string.Join(" ", (double[])startPoint)),
                                                                            new XAttribute("endPoint", string.Join(" ", (double[])endPoint))));

                    //Marshal.ReleaseComObject(edge);

                }
                roundElements.Add(edgeElements);
                //Marshal.ReleaseComObject(edges);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Round: Error Message:{ex.Message}");
                return new XElement("Round");
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
