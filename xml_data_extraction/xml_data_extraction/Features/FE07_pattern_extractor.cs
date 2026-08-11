using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using SolidEdgeGeometry;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    internal class FE07_pattern_extractor
    {

        private static XElement ExtractTransforms(object transforms, string featureLabel)
        {
            if (transforms is object[] transformsArray)
            {
                var transformsElement = new XElement("Transforms", new XAttribute("Count", transformsArray.Length));

                for (int t = 0; t < transformsArray.Length; t++)
                {
                    var matrixElement = new XElement("Transform", new XAttribute("Occurrence", t));

                    try
                    {
                        if (transformsArray[t] is Array matrixValues)
                        {
                            if (matrixValues.Rank == 2 && matrixValues.GetLength(0) == 4 && matrixValues.GetLength(1) == 4)
                            {
                                for (int row = 0; row < 4; row++)
                                {
                                    matrixElement.Add(new XElement($"Row{row}",
                                        new XAttribute("C0", matrixValues.GetValue(row, 0)),
                                        new XAttribute("C1", matrixValues.GetValue(row, 1)),
                                        new XAttribute("C2", matrixValues.GetValue(row, 2)),
                                        new XAttribute("C3", matrixValues.GetValue(row, 3))));
                                }
                            }
                            else if (matrixValues.Rank == 1)
                            {
                                for (int v = 0; v < matrixValues.Length; v++)
                                {
                                    matrixElement.Add(new XAttribute($"v{v}", matrixValues.GetValue(v)));
                                }
                            }
                            else
                            {
                                matrixElement.Add(new XAttribute("UnexpectedShape",
                                    $"Rank={matrixValues.Rank}, Dimensions=[{string.Join(",", Enumerable.Range(0, matrixValues.Rank).Select(d => matrixValues.GetLength(d)))}]"));
                            }
                        }
                        else
                        {
                            matrixElement.Add(new XAttribute("RawValue", transformsArray[t]?.ToString() ?? "null"));
                        }
                    }
                    catch (Exception ex)
                    {
                        matrixElement.Add(new XAttribute("Error", ex.Message));
                    }

                    transformsElement.Add(matrixElement);
                }

                return transformsElement;
            }

            return new XElement("Transforms", new XAttribute("RawValue", transforms?.ToString() ?? "null"));
        }

        private static XElement ExtractInputFeatures(Array inputFeatures)
        {
            var inputFeaturesElement = new XElement("InputFeatures", new XAttribute("Count", inputFeatures.Length));

            foreach (object featObj in inputFeatures)
            {
                try
                {
                    dynamic dynFeat = featObj;
                    inputFeaturesElement.Add(new XElement("InputFeature",
                        new XAttribute("Name", dynFeat.Name ?? ""),
                        new XAttribute("Type", dynFeat.Type)));
                }
                catch (Exception ex)
                {
                    inputFeaturesElement.Add(new XElement("InputFeature", new XAttribute("Error", ex.Message)));
                }
            }

            return inputFeaturesElement;
        }

        public static XElement Pattern(Pattern pattern)
        {
            XElement patternElements = new XElement("Pattern", new XAttribute("Type", -416228998));

            try
            {
                try { patternElements.Add(new XElement("name", pattern.Name)); }
                catch (Exception ex) { Console.WriteLine($"Pattern Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("pattern_type", pattern.PatternType)); }
                catch (Exception ex) { Console.WriteLine($"Pattern PatternType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("modelingModeType", pattern.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Pattern ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("showDimensions", pattern.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Pattern ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("visible", pattern.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Pattern Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("no_of_input_features", pattern.NumberOfInputFeatures)); }
                catch (Exception ex) { Console.WriteLine($"Pattern NumberOfInputFeatures: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("no_of_occurences", pattern.NumberOfOccurrences)); }
                catch (Exception ex) { Console.WriteLine($"Pattern NumberOfOccurrences: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("arcPattern", pattern.ArcPattern)); }
                catch (Exception ex) { Console.WriteLine($"Pattern ArcPattern: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { patternElements.Add(new XElement("skipCount", pattern.SkipCount)); }
                catch (Exception ex) { Console.WriteLine($"Pattern SkipCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynPattern = pattern;
                    patternElements.Add(new XElement("status", dynPattern.Status));
                    patternElements.Add(new XElement("suppress", dynPattern.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = pattern.GetStatusEx(out object description);
                    patternElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    pattern.GetDimensions(out int numDims, ref dims);
                    patternElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    patternElements.Add(MM01_geometry_methods.GetPlaneData(pattern.PatternPlane, "PatternPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetPatternPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array inputFeatures = Array.CreateInstance(typeof(object), 0);
                    pattern.GetInputFeatures(ref inputFeatures);
                    patternElements.Add(ExtractInputFeatures(inputFeatures));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetInputFeatures: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    patternElements.Add(ExtractTransforms(pattern.Transforms, "Pattern"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetTransforms: {ex.Message} | Inner: {ex.InnerException?.Message}");
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
                    dynamic dynPattern2 = pattern;
                    int occurrenceCount = pattern.NumberOfOccurrences;
                    var occurrencesElement = new XElement("Occurrences", new XAttribute("Count", occurrenceCount));

                    for (int occ = 1; occ <= occurrenceCount; occ++)
                    {
                        try
                        {
                            bool isSuppressed = dynPattern2.Suppressed(occ);
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

                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = pattern.Edges[edgeTyp];

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

                    patternElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = pattern.Faces[faceTyp];
                    patternElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Pattern GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pattern: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (pattern != null)
                {
                    Marshal.ReleaseComObject(pattern);
                    pattern = null;
                }
            }

            Console.WriteLine("\t Created Pattern XML list");
            return patternElements;
        }

        public static XElement UserDefinedPattern(UserDefinedPattern userDefinedPattern)
        {
            XElement udPatternElements = new XElement("UserDefinedPattern", new XAttribute("Type", -1468087919));

            try
            {
                try { udPatternElements.Add(new XElement("name", userDefinedPattern.Name)); }
                catch (Exception ex) { Console.WriteLine($"UserDefinedPattern Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { udPatternElements.Add(new XElement("type", userDefinedPattern.Type)); }
                catch (Exception ex) { Console.WriteLine($"UserDefinedPattern Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { udPatternElements.Add(new XElement("modelingModeType", userDefinedPattern.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"UserDefinedPattern ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { udPatternElements.Add(new XElement("showDimensions", userDefinedPattern.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"UserDefinedPattern ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { udPatternElements.Add(new XElement("visible", userDefinedPattern.Visible)); }
                catch (Exception ex) { Console.WriteLine($"UserDefinedPattern Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { udPatternElements.Add(new XElement("numberInputFeatures", userDefinedPattern.NumberInputFeatures)); }
                catch (Exception ex) { Console.WriteLine($"UserDefinedPattern NumberInputFeatures: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynUdPattern = userDefinedPattern;
                    udPatternElements.Add(new XElement("status", dynUdPattern.Status));
                    udPatternElements.Add(new XElement("suppress", dynUdPattern.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    udPatternElements.Add(MM01_geometry_methods.GetPlaneData(userDefinedPattern.PatternPlane, "PatternPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetPatternPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array inputFeatures = Array.CreateInstance(typeof(object), 0);
                    userDefinedPattern.GetInputFeatures(ref inputFeatures);
                    udPatternElements.Add(ExtractInputFeatures(inputFeatures));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetInputFeatures: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    userDefinedPattern.GetNumberOfOccurrences(out int numOccurrences, out int numFeaturesPerOccurrence);
                    udPatternElements.Add(new XElement("NumberOfOccurrences",
                        new XAttribute("Occurrences", numOccurrences),
                        new XAttribute("FeaturesPerOccurrence", numFeaturesPerOccurrence)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetNumberOfOccurrences: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    udPatternElements.Add(ExtractTransforms(userDefinedPattern.Transforms, "UserDefinedPattern"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetTransforms: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    userDefinedPattern.GetDimensions(out int numDims, ref dims);
                    udPatternElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    userDefinedPattern.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    udPatternElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = userDefinedPattern.Faces[faceTyp];
                    udPatternElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = userDefinedPattern.GetStatusEx(out object description);
                    udPatternElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx),
                        new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UserDefinedPattern GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"UserDefinedPattern: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (userDefinedPattern != null)
                {
                    Marshal.ReleaseComObject(userDefinedPattern);
                    userDefinedPattern = null;
                }
            }

            Console.WriteLine("\t Created UserDefinedPattern XML list");
            return udPatternElements;
        }
    }
}