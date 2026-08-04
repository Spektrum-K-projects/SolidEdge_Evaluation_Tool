using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE07_pattern_extractor
    {
        private static void TryAdd(XElement parent, string label, Func<object> getter)
        {
            try
            {
                parent.Add(new XElement(label, getter()));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pattern {label}: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }
        }

        public static XElement Pattern(Pattern pattern)
        {
            XElement patternElements = new XElement("Pattern", new XAttribute("Type", -416228998));

            try
            {
                TryAdd(patternElements, "name", () => pattern.Name);
                TryAdd(patternElements, "pattern_type", () => pattern.PatternType);
                TryAdd(patternElements, "modelingModeType", () => pattern.ModelingModeType);
                TryAdd(patternElements, "no_of_input_features", () => pattern.NumberOfInputFeatures);
                TryAdd(patternElements, "no_of_occurences", () => pattern.NumberOfOccurrences);
                TryAdd(patternElements, "status", () => { dynamic d = pattern; return d.Status; });
                TryAdd(patternElements, "arcPattern", () => pattern.ArcPattern);
                TryAdd(patternElements, "skipCount", () => pattern.SkipCount);



                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    pattern.GetDimensions(out int numDims, ref dims);
                    patternElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array originPoint = Array.CreateInstance(typeof(double), 0);
                    pattern.GetOriginPosition(ref originPoint);
                    if (originPoint.Length >= 3)
                        patternElements.Add(new XElement("OriginPosition",
                            new XAttribute("X", originPoint.GetValue(0)),
                            new XAttribute("Y", originPoint.GetValue(1)),
                            new XAttribute("Z", originPoint.GetValue(2))));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetOriginPosition: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array axisPoint = Array.CreateInstance(typeof(double), 0);
                    pattern.GetAxisPosition(ref axisPoint);
                    if (axisPoint.Length >= 3)
                        patternElements.Add(new XElement("AxisPosition",
                            new XAttribute("X", axisPoint.GetValue(0)),
                            new XAttribute("Y", axisPoint.GetValue(1)),
                            new XAttribute("Z", axisPoint.GetValue(2))));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetAxisPosition: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    pattern.GetRectangularPatternData(out PatternOffsetTypeConstants method, out int xCount, out int yCount,
                                                        out double xSpacing, out double ySpacing, out double angle);
                    patternElements.Add(new XElement("RectangularPattern",
                        new XElement("method", method), new XElement("x_count", xCount), new XElement("y_count", yCount),
                        new XElement("x_spacing", xSpacing), new XElement("y_spacing", ySpacing), new XElement("angle", angle)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern: not rectangular (expected if this pattern is a different kind) - {ex.Message}");
                }

                try
                {
                    pattern.GetCircularPatternData(out PatternOffsetTypeConstants method, out int radialCount, out double angleSpacing);
                    patternElements.Add(new XElement("CircularPattern",
                        new XElement("method", method), new XElement("radial_count", radialCount), new XElement("angle_spacing", angleSpacing)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern: not circular (expected if this pattern is a different kind) - {ex.Message}");
                }

                try
                {
                    pattern.GetFillPatternData(out FillPatternMethodConstants method, out double xSpacing, out double ySpacingOrAngle,
                                                out double staggerOffset, out double orientAngle, out bool centerOrientOn,
                                                out double centerOrientAngle, out double regionOffset);
                    patternElements.Add(new XElement("FillPattern",
                        new XElement("method", method), new XElement("x_spacing", xSpacing), new XElement("y_spacing_or_angle", ySpacingOrAngle),
                        new XElement("stagger_complex_linear_offset", staggerOffset), new XElement("orient_vector_angle", orientAngle),
                        new XElement("center_orient_on", centerOrientOn), new XElement("center_orient_angle", centerOrientAngle),
                        new XElement("region_offset", regionOffset)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern: not fill-type (expected if this pattern is a different kind) - {ex.Message}");
                }

                try
                {
                    pattern.GetCurve1Data(out PatternCurveAnchorSideConstants anchorSide, out double anchorDistance,
                                            out PatternOffsetTypeConstants method, out int count, out double spacing);
                    patternElements.Add(new XElement("Curve1Pattern",
                        new XElement("anchor_side", anchorSide), new XElement("anchor_at_distance", anchorDistance),
                        new XElement("method", method), new XElement("count", count), new XElement("spacing", spacing)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern: no Curve1 data (expected if this pattern is a different kind) - {ex.Message}");
                }

                try
                {
                    pattern.GetCurve2Data(out PatternCurveAnchorSideConstants anchorSide, out double anchorDistance,
                                            out PatternOffsetTypeConstants method, out int count, out double spacing);
                    patternElements.Add(new XElement("Curve2Pattern",
                        new XElement("anchor_side", anchorSide), new XElement("anchor_at_distance", anchorDistance),
                        new XElement("method", method), new XElement("count", count), new XElement("spacing", spacing)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern: no Curve2 data (expected if this pattern is a different kind) - {ex.Message}");
                }

                try
                {
                    dynamic dynPattern = pattern;
                    int occurrenceCount = pattern.NumberOfOccurrences;
                    var occurrencesElement = new XElement("Occurrences", new XAttribute("Count", occurrenceCount));

                    for (int occ = 1; occ <= occurrenceCount; occ++)
                    {
                        try
                        {
                            bool isSuppressed = dynPattern.Suppressed(occ);
                            occurrencesElement.Add(new XElement("Occurrence",
                                new XAttribute("Index", occ), new XAttribute("Suppressed", isSuppressed)));
                        }
                        catch (Exception ex)
                        {
                            occurrencesElement.Add(new XElement("Occurrence",
                                new XAttribute("Index", occ), new XAttribute("Error", ex.Message)));
                        }
                    }

                    patternElements.Add(occurrencesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern Suppressed loop: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    var profile = pattern.Profile;
                    if (profile != null)
                    {
                        patternElements.Add(GE04_getProfiles_extractor.Profile_Data(profile));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetProfile: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    pattern.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    patternElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pattern: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Pattern", "Error");
            }
            finally
            {
                if (pattern != null)
                {
                    Marshal.ReleaseComObject(pattern);
                    pattern = null;
                }
            }

            Console.WriteLine("Created Pattern XML list");
            return patternElements;
        }
    }
}