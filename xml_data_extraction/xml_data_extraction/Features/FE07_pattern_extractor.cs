using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE07_pattern_extractor
    {
        public static XElement Pattern(Pattern pattern)
        {
            XElement patternElements = new XElement("Pattern", new XAttribute("Type", -416228998));

            try
            {
                var pattern_Name = pattern.Name;
                patternElements.Add(new XElement("name", pattern_Name));
                Console.WriteLine($"Pattern Name: {pattern_Name}");

                //var pattern_Method = pattern.PatternMethod;
                //patternElements.Add(new XElement("method", pattern_Method.GetType()));
                //Console.WriteLine($"Pattern Method: {pattern_Method.GetType()}");

                var pattern_Type = pattern.PatternType;
                patternElements.Add(new XElement("type", pattern_Type));
                Console.WriteLine($"Pattern type: {pattern_Type}");

                var pattern_NoOfInputFeatures = pattern.NumberOfInputFeatures;
                patternElements.Add(new XElement("no_of_input_features", pattern_NoOfInputFeatures));
                Console.WriteLine($"Pattern Input Features: {pattern_NoOfInputFeatures}");

                var pattern_NoOfOccurences = pattern.NumberOfOccurrences;
                patternElements.Add(new XElement("no_of_occurences", pattern_NoOfOccurences));
                Console.WriteLine($"Pattern Number of Occurences: {pattern_NoOfOccurences}");

                var pattern_direction = pattern.PatternDirection.GetType().ToString;
                patternElements.Add(new XElement("direction", pattern_direction));
                Console.WriteLine($"Pattern Direction: {pattern_direction}");

                //var pattern_Plane = pattern.PatternPlane;
                //patternElements.Add(new XElement("plane", pattern_Plane));
                //Console.WriteLine($"Pattern Plane: {pattern_Plane}");

                //var pattern_ArcPattern = pattern.ArcPattern.ToString();
                //patternElements.Add(new XElement("arc_pattern", pattern_ArcPattern));
                //Console.WriteLine($"Pattern Arc: {pattern_ArcPattern}");

                //Rectangular Pattern Data Extract
                pattern.GetRectangularPatternData(out PatternOffsetTypeConstants Rect_Pattern_Method, out int xDirCount, out int yDirCount, 
                                                                            out double xDirSpacing,out double yDirSpacing, out double RectAngle);
                patternElements.Add(new XElement("rectanglar_pattern", new XElement ("method", Rect_Pattern_Method),
                                                                        new XElement("x_Direction_Count", xDirCount),
                                                                        new XElement("y_Direction_Count", yDirCount),
                                                                        new XElement("x_Direction_Spacing", xDirSpacing),
                                                                        new XElement("y_Direction_Spacing", yDirSpacing),
                                                                        new XElement("rectangle_angle", RectAngle)));

                //Circular Pattern Data Extract
                pattern.GetCircularPatternData(out PatternOffsetTypeConstants Circ_Pattern_Method, out int radial_Count, out double angle_Spacing);
                patternElements.Add(new XElement("circular_pattern", new XElement("method", Circ_Pattern_Method),
                                                                        new XElement("radial_count", radial_Count),
                                                                        new XElement("angle_spacing", angle_Spacing)));

                //Fill Pattern Data Extract
                pattern.GetFillPatternData(out FillPatternMethodConstants Fill_Pattern_Method, out double xDir_Spacing, out double yDir_Spacing,
                     out double linear_Offset, out double orient_Vector_Angle, out bool center_Orient_on, out double center_Orient_Angle, out double region_Offset);
                patternElements.Add(new XElement("fill_pattern", new XElement("method", Fill_Pattern_Method),
                                                                        new XElement("x_Direction_Spacing", xDir_Spacing),
                                                                        new XElement("y_Direction_Spacing", yDir_Spacing),
                                                                        new XElement("linear_offset", linear_Offset),
                                                                        new XElement("vector_angle", orient_Vector_Angle),
                                                                        new XElement("center_orientation", center_Orient_on),
                                                                        new XElement("center_orientation_angle", center_Orient_Angle),
                                                                        new XElement("region_offset", region_Offset)));

                

                Marshal.ReleaseComObject(Rect_Pattern_Method);
                Marshal.ReleaseComObject(Circ_Pattern_Method);
                Marshal.ReleaseComObject(Fill_Pattern_Method);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Pattern: Error Message:{ex.Message}");
                return new XElement("Pattern");
            }

            finally
            {
                if (pattern != null)
                {
                    Marshal.ReleaseComObject(pattern);
                    pattern = null;
                }
            }


            return patternElements;
        }
    }
}
