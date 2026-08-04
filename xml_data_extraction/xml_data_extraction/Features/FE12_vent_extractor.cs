using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE12_vent_extractor
    {
        private static void TryAdd(XElement parent, string label, Func<object> getter)
        {
            try
            {
                parent.Add(new XElement(label, getter()));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Vent {label}: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }
        }

        public static XElement Vent(Vent vent)
        {
            XElement ventElements = new XElement("Vent", new XAttribute("Type", -85880079));

            try
            {
                TryAdd(ventElements, "name", () => vent.Name);
                TryAdd(ventElements, "type", () => vent.Type);
                TryAdd(ventElements, "modelingModeType", () => vent.ModelingModeType);

                TryAdd(ventElements, "extentSide", () => vent.ExtentSide);

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

                TryAdd(ventElements, "ribDepth", () => vent.RibDepth);
                TryAdd(ventElements, "ribThickness", () => vent.RibThickness);
                TryAdd(ventElements, "ribExtension", () => vent.RibExtension);
                TryAdd(ventElements, "ribOffset", () => vent.RibOffset);

                TryAdd(ventElements, "roundEnabled", () => vent.RoundEnabled);
                TryAdd(ventElements, "roundRadius", () => vent.RoundRadius);

                TryAdd(ventElements, "sparDepth", () => vent.SparDepth);
                TryAdd(ventElements, "sparExtension", () => vent.SparExtension);
                TryAdd(ventElements, "sparOffset", () => vent.SparOffset);
                TryAdd(ventElements, "sparThickness", () => vent.SparThickness);

                TryAdd(ventElements, "draftAngle", () => vent.DraftAngle);
                TryAdd(ventElements, "draftEnabled", () => vent.DraftEnabled);
                TryAdd(ventElements, "draftFromOutsideEdges", () => vent.DraftFromOutsideEdges);
                TryAdd(ventElements, "draftSide", () => vent.DraftSide);

                TryAdd(ventElements, "status", () => { dynamic d = vent; return d.Status; });
                TryAdd(ventElements, "suppress", () => { dynamic d = vent; return d.Suppress; });

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    vent.GetDimensions(out int numDims, ref dims);
                    ventElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
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

                try
                {
                    Array boundaryCurves = Array.CreateInstance(typeof(object), 0);
                    vent.GetBoundaryCurves(out int numBoundary, ref boundaryCurves);
                    var boundaryElement = new XElement("BoundaryCurves", new XAttribute("Count", numBoundary));
                    for (int i = 0; i < boundaryCurves.Length; i++)
                    {
                        var curveObj = boundaryCurves.GetValue(i);
                        boundaryElement.Add(new XElement("Curve", new XAttribute("Index", i), new XAttribute("Type", curveObj?.GetType().Name ?? "null")));
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
                        var curveObj = ribCurves.GetValue(i);
                        ribElement.Add(new XElement("Curve", new XAttribute("Index", i), new XAttribute("Type", curveObj?.GetType().Name ?? "null")));
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
                        var curveObj = sparCurves.GetValue(i);
                        sparElement.Add(new XElement("Curve", new XAttribute("Index", i), new XAttribute("Type", curveObj?.GetType().Name ?? "null")));
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
                // Should rarely fire now - everything above is individually isolated.
                Console.WriteLine($"Vent: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Vent", "Error");
            }
            finally
            {
                if (vent != null)
                {
                    Marshal.ReleaseComObject(vent);
                    vent = null;
                }
            }

            Console.WriteLine("Created Vent XML list");
            return ventElements;
        }
    }
}