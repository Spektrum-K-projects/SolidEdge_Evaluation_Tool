using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    //Swept Protrusion, Swept Cutout, Solid Swept Protrusion & Solid Swept Cutout Data Extraction
    internal class FE18_sweep_extractor
    {
        public static XElement Swept_Protrusion(SweptProtrusion sweptProtrusion)
        {
            XElement sweptProtrusionElements = new XElement("SweptProtrusion", new XAttribute("Type", -2101194894));

            try
            {
                try { sweptProtrusionElements.Add(new XElement("name", sweptProtrusion.Name)); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptProtrusionElements.Add(new XElement("type", sweptProtrusion.Type)); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptProtrusionElements.Add(new XElement("modelingModeType", sweptProtrusion.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptProtrusionElements.Add(new XElement("sectionAlignment", sweptProtrusion.SectionAlignment.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion SectionAlignment: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptProtrusionElements.Add(new XElement("faceMerge", sweptProtrusion.FaceMerge.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion FaceMerge: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptProtrusionElements.Add(new XElement("faceContinuity", sweptProtrusion.FaceContinuity.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion FaceContinuity: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptProtrusionElements.Add(new XElement("maintainCrossSectionGeometry", sweptProtrusion.MaintainCrossSectionGeometry)); }
                catch (Exception ex) { Console.WriteLine($"SweptProtrusion MaintainCrossSectionGeometry: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- LockAxis  ----
                try
                {
                    object lockAxis = sweptProtrusion.LockAxis;
                    sweptProtrusionElements.Add(new XElement("LockAxis",
                        new XAttribute("UnderlyingType", lockAxis?.GetType().Name ?? "null")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetLockAxis: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Trace curves ----
                try
                {
                    Array traceCurves = Array.CreateInstance(typeof(object), 0);
                    sweptProtrusion.GetTraceCurves(out int numTraceCurves, ref traceCurves);

                    XElement traceCurvesElement = new XElement("TraceCurves", new XAttribute("Count", numTraceCurves));
                    int tcIndex = 0;
                    foreach (object tcObj in traceCurves)
                    {
                        tcIndex++;
                        traceCurvesElement.Add(new XElement($"curve{tcIndex}",
                            new XAttribute("UnderlyingType", tcObj?.GetType().Name ?? "null")));
                    }
                    sweptProtrusionElements.Add(traceCurvesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetTraceCurves: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Cross-sections ----
                try
                {
                    Array crossSections = Array.CreateInstance(typeof(object), 0);
                    sweptProtrusion.GetCrossSections(out int numCrossSections, ref crossSections);

                    XElement crossSectionsElement = new XElement("CrossSections", new XAttribute("Count", numCrossSections));
                    foreach (object csObj in crossSections)
                    {
                        if (csObj is Profile csProfile)
                        {
                            crossSectionsElement.Add(GE04_getProfiles_extractor.Profile_Data(csProfile));
                        }
                    }
                    sweptProtrusionElements.Add(crossSectionsElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetCrossSections: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    sweptProtrusion.GetDimensions(out int numDims, ref dims);
                    sweptProtrusionElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    sweptProtrusion.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    sweptProtrusionElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = sweptProtrusion.Edges[edgeTyp];

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

                    sweptProtrusionElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = sweptProtrusion.Faces[faceTyp];
                    sweptProtrusionElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynSweptProtrusion = sweptProtrusion;
                    sweptProtrusionElements.Add(new XElement("status", dynSweptProtrusion.Status));
                    sweptProtrusionElements.Add(new XElement("suppress", dynSweptProtrusion.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = sweptProtrusion.GetStatusEx(out object description);
                    sweptProtrusionElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptProtrusion GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SweptProtrusion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("SweptProtrusion", "Error");
            }
            finally
            {
                if (sweptProtrusion != null)
                {
                    Marshal.ReleaseComObject(sweptProtrusion);
                    sweptProtrusion = null;
                }
            }

            Console.WriteLine("Created Swept Protrusion XML list");
            return sweptProtrusionElements;
        }

        public static XElement Swept_Cutout(SweptCutout sweptCutout)
        {
            XElement sweptCutoutElements = new XElement("SweptCutout", new XAttribute("Type", -398746894));

            try
            {
                try { sweptCutoutElements.Add(new XElement("name", sweptCutout.Name)); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptCutoutElements.Add(new XElement("type", sweptCutout.Type)); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptCutoutElements.Add(new XElement("modelingModeType", sweptCutout.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptCutoutElements.Add(new XElement("sectionAlignment", sweptCutout.SectionAlignment.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout SectionAlignment: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptCutoutElements.Add(new XElement("faceMerge", sweptCutout.FaceMerge.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout FaceMerge: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptCutoutElements.Add(new XElement("faceContinuity", sweptCutout.FaceContinuity.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout FaceContinuity: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sweptCutoutElements.Add(new XElement("maintainCrossSectionGeometry", sweptCutout.MaintainCrossSectionGeometry)); }
                catch (Exception ex) { Console.WriteLine($"SweptCutout MaintainCrossSectionGeometry: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // ---- LockAxis ----
                try
                {
                    object lockAxis = sweptCutout.LockAxis;
                    sweptCutoutElements.Add(new XElement("LockAxis",
                        new XAttribute("UnderlyingType", lockAxis?.GetType().Name ?? "null")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetLockAxis: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Trace curves ----
                try
                {
                    Array traceCurves = Array.CreateInstance(typeof(object), 0);
                    sweptCutout.GetTraceCurves(out int numTraceCurves, ref traceCurves);

                    XElement traceCurvesElement = new XElement("TraceCurves", new XAttribute("Count", numTraceCurves));
                    int tcIndex = 0;
                    foreach (object tcObj in traceCurves)
                    {
                        tcIndex++;
                        traceCurvesElement.Add(new XElement($"curve{tcIndex}",
                            new XAttribute("UnderlyingType", tcObj?.GetType().Name ?? "null")));
                    }
                    sweptCutoutElements.Add(traceCurvesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetTraceCurves: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Cross-sections ----
                try
                {
                    Array crossSections = Array.CreateInstance(typeof(object), 0);
                    sweptCutout.GetCrossSections(out int numCrossSections, ref crossSections);

                    XElement crossSectionsElement = new XElement("CrossSections", new XAttribute("Count", numCrossSections));
                    foreach (object csObj in crossSections)
                    {
                        if (csObj is Profile csProfile)
                        {
                            crossSectionsElement.Add(GE04_getProfiles_extractor.Profile_Data(csProfile));
                        }
                    }
                    sweptCutoutElements.Add(crossSectionsElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetCrossSections: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Dimensions ----
                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    sweptCutout.GetDimensions(out int numDims, ref dims);
                    sweptCutoutElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    sweptCutout.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    sweptCutoutElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Edges ----
                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = sweptCutout.Edges[edgeTyp];

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

                    sweptCutoutElements.Add(edgeElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Faces ----
                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = sweptCutout.Faces[faceTyp];
                    sweptCutoutElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Body array ----
                try
                {
                    Array bodyArray = Array.CreateInstance(typeof(object), 0);
                    sweptCutout.GetBodyArray(out bool multiBodyCut, out int numberOfBodies, out bodyArray);
                    sweptCutoutElements.Add(new XElement("BodyArray",
                        new XAttribute("MultiBodyCut", multiBodyCut),
                        new XAttribute("NumberOfBodies", numberOfBodies)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetBodyArray: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Status / Suppress ----
                try
                {
                    dynamic dynSweptCutout = sweptCutout;
                    sweptCutoutElements.Add(new XElement("status", dynSweptCutout.Status));
                    sweptCutoutElements.Add(new XElement("suppress", dynSweptCutout.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- GetStatusEx ----
                try
                {
                    FeatureStatusConstants statusEx = sweptCutout.GetStatusEx(out object description);
                    sweptCutoutElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"SweptCutout GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SweptCutout: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("SweptCutout", "Error");
            }
            finally
            {
                if (sweptCutout != null)
                {
                    Marshal.ReleaseComObject(sweptCutout);
                    sweptCutout = null;
                }
            }

            Console.WriteLine("\t Created Swept Cutout XML list");
            return sweptCutoutElements;
        }
    }
}