using SolidEdgeGeometry;
using SolidEdgePart;
using SolidEdgeFramework;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE12_vent_extractor
    {

        private static XElement ExtractCurveEntry(object curveObj, int index)
        {
            var curveElement = new XElement("Curve", new XAttribute("Index", index));

            if (curveObj == null)
            {
                curveElement.Add(new XAttribute("Type", "null"));
                return curveElement;
            }

            try
            {
                if (curveObj is Edge singleEdge)
                {
                    curveElement.Add(new XAttribute("Type", "igEdge"));

                    try
                    {
                        Array startPoint = Array.CreateInstance(typeof(double), 0);
                        Array endPoint = Array.CreateInstance(typeof(double), 0);
                        singleEdge.GetEndPoints(ref startPoint, ref endPoint);
                        curveElement.Add(new XAttribute("startpoint", string.Join(" ", (double[])startPoint)));
                        curveElement.Add(new XAttribute("endPoint", string.Join(" ", (double[])endPoint)));
                    }
                    catch (Exception ex)
                    {
                        curveElement.Add(new XAttribute("EndPointsError", ex.Message));
                    }
                }
                else if (curveObj is Edges edgeCollection)
                {
                    curveElement.Add(new XAttribute("Type", "igEdges"), new XAttribute("Count", edgeCollection.Count));

                    for (int e = 1; e <= edgeCollection.Count; e++)
                    {
                        var edge = (Edge)edgeCollection.Item(e);
                        var edgeElement = new XElement("edge", new XAttribute("index", e));

                        try
                        {
                            Array startPoint = Array.CreateInstance(typeof(double), 0);
                            Array endPoint = Array.CreateInstance(typeof(double), 0);
                            edge.GetEndPoints(ref startPoint, ref endPoint);
                            edgeElement.Add(new XAttribute("startpoint", string.Join(" ", (double[])startPoint)));
                            edgeElement.Add(new XAttribute("endPoint", string.Join(" ", (double[])endPoint)));
                        }
                        catch (Exception ex)
                        {
                            edgeElement.Add(new XAttribute("EndPointsError", ex.Message));
                        }

                        curveElement.Add(edgeElement);
                        Marshal.ReleaseComObject(edge);
                    }
                }
                else
                {
                    dynamic dynCurve = curveObj;
                    curveElement.Add(new XAttribute("Type", ((GNTTypePropertyConstants)dynCurve.Type).ToString()));
                }
            }
            catch (Exception ex)
            {
                curveElement.Add(new XAttribute("Error", ex.Message));
            }

            return curveElement;
        }

        public static XElement Vent(Vent vent)
        {
            XElement ventElements = new XElement("Vent", new XAttribute("Type", -85880079));

            try
            {
                try { ventElements.Add(new XElement("name", vent.Name)); }
                catch (Exception ex) { Console.WriteLine($"Vent Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("type", vent.Type)); }
                catch (Exception ex) { Console.WriteLine($"Vent Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("modelingModeType", vent.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Vent ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("showDimensions", vent.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Vent ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("visible", vent.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Vent Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("extentSide", vent.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Vent ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var extentType = vent.ExtentType;
                    ventElements.Add(new XElement("extentType", extentType));

                    if (extentType == VentExtentTypeConstants.seVentExtentTypeFinite)
                    {
                        try
                        {
                            ventElements.Add(new XElement("extentDepth", vent.ExtentDepth));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Vent extentDepth: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                    }
                    else
                    {
                        ventElements.Add(new XElement("extentDepth", new XAttribute("NotApplicable", extentType.ToString())));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent extentType: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { ventElements.Add(new XElement("ribDepth", vent.RibDepth)); }
                catch (Exception ex) { Console.WriteLine($"Vent RibDepth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("ribThickness", vent.RibThickness)); }
                catch (Exception ex) { Console.WriteLine($"Vent RibThickness: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("ribExtension", vent.RibExtension)); }
                catch (Exception ex) { Console.WriteLine($"Vent RibExtension: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("ribOffset", vent.RibOffset)); }
                catch (Exception ex) { Console.WriteLine($"Vent RibOffset: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("roundEnabled", vent.RoundEnabled)); }
                catch (Exception ex) { Console.WriteLine($"Vent RoundEnabled: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("roundRadius", vent.RoundRadius)); }
                catch (Exception ex) { Console.WriteLine($"Vent RoundRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("sparDepth", vent.SparDepth)); }
                catch (Exception ex) { Console.WriteLine($"Vent SparDepth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("sparExtension", vent.SparExtension)); }
                catch (Exception ex) { Console.WriteLine($"Vent SparExtension: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("sparOffset", vent.SparOffset)); }
                catch (Exception ex) { Console.WriteLine($"Vent SparOffset: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("sparThickness", vent.SparThickness)); }
                catch (Exception ex) { Console.WriteLine($"Vent SparThickness: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("draftAngle", vent.DraftAngle)); }
                catch (Exception ex) { Console.WriteLine($"Vent DraftAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("draftEnabled", vent.DraftEnabled)); }
                catch (Exception ex) { Console.WriteLine($"Vent DraftEnabled: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("draftFromOutsideEdges", vent.DraftFromOutsideEdges)); }
                catch (Exception ex) { Console.WriteLine($"Vent DraftFromOutsideEdges: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { ventElements.Add(new XElement("draftSide", vent.DraftSide)); }
                catch (Exception ex) { Console.WriteLine($"Vent DraftSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynVent = vent;
                    ventElements.Add(new XElement("status", dynVent.Status));
                    ventElements.Add(new XElement("suppress", dynVent.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    vent.GetDimensions(out int numDims, ref dims);
                    ventElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    vent.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    ventElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = vent.GetStatusEx(out object description);
                    ventElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = vent.Edges[edgeTyp];

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

                    ventElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = vent.Faces[faceTyp];
                    ventElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Boundary / rib / spar curves ----
                try
                {
                    Array boundaryCurves = Array.CreateInstance(typeof(object), 0);
                    vent.GetBoundaryCurves(out int numBoundary, ref boundaryCurves);
                    var boundaryElement = new XElement("BoundaryCurves", new XAttribute("Count", numBoundary));
                    for (int i = 0; i < boundaryCurves.Length; i++)
                    {
                        boundaryElement.Add(ExtractCurveEntry(boundaryCurves.GetValue(i), i));
                    }
                    ventElements.Add(boundaryElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetBoundaryCurves: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array ribCurves = Array.CreateInstance(typeof(object), 0);
                    vent.GetRibCurves(out int numRib, ref ribCurves);
                    var ribElement = new XElement("RibCurves", new XAttribute("Count", numRib));
                    for (int i = 0; i < ribCurves.Length; i++)
                    {
                        ribElement.Add(ExtractCurveEntry(ribCurves.GetValue(i), i));
                    }
                    ventElements.Add(ribElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetRibCurves: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array sparCurves = Array.CreateInstance(typeof(object), 0);
                    vent.GetSparCurves(out int numSpar, ref sparCurves);
                    var sparElement = new XElement("SparCurves", new XAttribute("Count", numSpar));
                    for (int i = 0; i < sparCurves.Length; i++)
                    {
                        sparElement.Add(ExtractCurveEntry(sparCurves.GetValue(i), i));
                    }
                    ventElements.Add(sparElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Vent GetSparCurves: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Vent: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (vent != null)
                {
                    Marshal.ReleaseComObject(vent);
                    vent = null;
                }
            }

            Console.WriteLine("\t Created Vent XML list");
            return ventElements;
        }
    }
}