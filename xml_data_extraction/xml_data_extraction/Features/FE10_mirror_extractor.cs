using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    internal class FE10_mirror_extractor
    {
        public static XElement MirrorPart(MirrorPart mirrorPart)
        {
            XElement mirrorPartElements = new XElement("Mirror_Part", new XAttribute("Type", 1908287958));

            try
            {
                try { mirrorPartElements.Add(new XElement("name", mirrorPart.Name)); }
                catch (Exception ex) { Console.WriteLine($"MirrorPart Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { mirrorPartElements.Add(new XElement("type", mirrorPart.Type)); }
                catch (Exception ex) { Console.WriteLine($"MirrorPart Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { mirrorPartElements.Add(new XElement("modelingModeType", mirrorPart.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"MirrorPart ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynMirror = mirrorPart;
                    mirrorPartElements.Add(new XElement("status", dynMirror.Status));
                    mirrorPartElements.Add(new XElement("suppress", dynMirror.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = mirrorPart.GetStatusEx(out object description);
                    mirrorPartElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    mirrorPartElements.Add(new XElement("removeOriginal", mirrorPart.RemoveOriginal));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart RemoveOriginal: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    mirrorPart.GetDimensions(out int numDims, ref dims);
                    mirrorPartElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    mirrorPart.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    mirrorPartElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    mirrorPart.ExactRange(out double ex1, out double ey1, out double ez1, out double ex2, out double ey2, out double ez2);
                    mirrorPartElements.Add(new XElement("ExactRange",
                        new XAttribute("X1", ex1), new XAttribute("Y1", ey1), new XAttribute("Z1", ez1),
                        new XAttribute("X2", ex2), new XAttribute("Y2", ey2), new XAttribute("Z2", ez2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetExactRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array bodies = Array.CreateInstance(typeof(object), 0);
                    mirrorPart.GetBodies(out int numBodies, ref bodies);
                    mirrorPartElements.Add(new XElement("Bodies", new XAttribute("Count", numBodies)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetBodies: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    mirrorPartElements.Add(MM01_geometry_methods.GetPlaneData(mirrorPart.MirrorPlane, "MirrorPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorPart GetMirrorPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"MirrorPart: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (mirrorPart != null)
                {
                    Marshal.ReleaseComObject(mirrorPart);
                    mirrorPart = null;
                }
            }

            Console.WriteLine("Created Mirror Part XML list");
            return mirrorPartElements;
        }

        public static XElement MirrorCopy(MirrorCopy mirrorCopy)
        {
            XElement mirrorCopyElements = new XElement("MirrorCopy", new XAttribute("Type", 66247736));

            try
            {
                try { mirrorCopyElements.Add(new XElement("name", mirrorCopy.Name)); }
                catch (Exception ex) { Console.WriteLine($"MirrorCopy Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { mirrorCopyElements.Add(new XElement("type", mirrorCopy.Type)); }
                catch (Exception ex) { Console.WriteLine($"MirrorCopy Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { mirrorCopyElements.Add(new XElement("isVisible", mirrorCopy.Visible)); }
                catch (Exception ex) { Console.WriteLine($"MirrorCopy Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { mirrorCopyElements.Add(new XElement("modelingModeType", mirrorCopy.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"MirrorCopy ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { mirrorCopyElements.Add(new XElement("NumberOfInputFeatures", mirrorCopy.NumberInputFeatures)); }
                catch (Exception ex) { Console.WriteLine($"MirrorCopy NumberInputFeatures: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    mirrorCopy.GetNumberOfOccurrences(out int numOcc, out int numFea);
                    mirrorCopyElements.Add(new XElement("numberOfOccurrences", numOcc));
                    mirrorCopyElements.Add(new XElement("numberOfFeaturesPerOccurrences", numFea));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetNumberOfOccurrences: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynMirrorCopy = mirrorCopy;
                    mirrorCopyElements.Add(new XElement("status", dynMirrorCopy.Status));
                    mirrorCopyElements.Add(new XElement("suppress", dynMirrorCopy.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = mirrorCopy.GetStatusEx(out object description);
                    mirrorCopyElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    mirrorCopy.GetDimensions(out int numDims, ref dims);
                    mirrorCopyElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    mirrorCopy.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    mirrorCopyElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    mirrorCopyElements.Add(MM01_geometry_methods.GetPlaneData(mirrorCopy.PatternPlane, "PatternPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetPatternPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    object transforms = mirrorCopy.Transforms;
                    if (transforms is Array occurrenceArray && occurrenceArray.Rank == 1)
                    {
                        var transformsElement = new XElement("Transforms", new XAttribute("Count", occurrenceArray.Length));

                        for (int i = 0; i < occurrenceArray.Length; i++)
                        {
                            var occurrenceElement = new XElement("Occurrence", new XAttribute("Index", i));

                            try
                            {
                                if (occurrenceArray.GetValue(i) is Array matrix && matrix.Rank == 2)
                                {
                                    int rows = matrix.GetLength(0);
                                    int cols = matrix.GetLength(1);
                                    occurrenceElement.Add(new XAttribute("Rows", rows), new XAttribute("Columns", cols));

                                    for (int r = 0; r < rows; r++)
                                    {
                                        var rowElement = new XElement($"Row{r}");
                                        for (int c = 0; c < cols; c++)
                                        {
                                            rowElement.Add(new XAttribute($"C{c}", matrix.GetValue(r, c)));
                                        }
                                        occurrenceElement.Add(rowElement);
                                    }
                                }
                                else
                                {
                                    occurrenceElement.Add(new XAttribute("UnexpectedShape",
                                        occurrenceArray.GetValue(i)?.GetType().Name ?? "null"));
                                }
                            }
                            catch (Exception ex)
                            {
                                occurrenceElement.Add(new XAttribute("Error", ex.Message));
                            }

                            transformsElement.Add(occurrenceElement);
                        }

                        mirrorCopyElements.Add(transformsElement);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MirrorCopy GetTransforms: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"MirrorCopy: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (mirrorCopy != null)
                {
                    Marshal.ReleaseComObject(mirrorCopy);
                    mirrorCopy = null;
                }
            }

            Console.WriteLine("\t Created Mirror Copy XML list");
            return mirrorCopyElements;
        }
    }
}