using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE14_helical_curve_extractor
    {
        public static XElement HelicalCurve_Extract(HelicalCurve helicalCurve)
        {
            XElement helicalCurveElements = new XElement("HelicalCurve");

            try
            {
                helicalCurveElements.Add(new XElement("name", helicalCurve.Name));
                helicalCurveElements.Add(new XElement("type", helicalCurve.Type));
                helicalCurveElements.Add(new XElement("modelingModeType", helicalCurve.ModelingModeType));
                helicalCurveElements.Add(new XElement("suppress", helicalCurve.Suppress));

                try
                {
                    dynamic dynCurve = helicalCurve;
                    helicalCurveElements.Add(new XElement("status", dynCurve.Status));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelicalCurve GetStatus: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    helicalCurve.GetDimensions(out int numDims, ref dims);
                    helicalCurveElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelicalCurve GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    helicalCurve.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    helicalCurveElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelicalCurve GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // The actual pitch/turns/height data - the reason this extractor exists.
                // NumOfSections is an INPUT, not something we read back - passing a generous
                // upper bound since we don't know a curve's real section count in advance.
                // Untested assumption: verify against a real multi-section helix if you have one.
                try
                {
                    int numOfSections = 10;
                    helicalCurve.getHelicalCurveData(numOfSections,
                        out HelicalCurveMethodType methodType,
                        out HelicalCurveTaperByType taperType,
                        out bool rightHanded,
                        out bool directionStartToEnd,
                        out Array height,
                        out Array pitch,
                        out Array turns,
                        out Array diameter);

                    var dataElement = new XElement("HelicalCurveData",
                        new XAttribute("MethodType", methodType),
                        new XAttribute("TaperType", taperType),
                        new XAttribute("RightHanded", rightHanded),
                        new XAttribute("DirectionStartToEnd", directionStartToEnd));

                    int sectionCount = height?.Length ?? 0;
                    for (int s = 0; s < sectionCount; s++)
                    {
                        var sectionElement = new XElement("Section", new XAttribute("Index", s));
                        if (s < height.Length) sectionElement.Add(new XAttribute("Height", height.GetValue(s)));
                        if (pitch != null && s < pitch.Length) sectionElement.Add(new XAttribute("Pitch", pitch.GetValue(s)));
                        if (turns != null && s < turns.Length) sectionElement.Add(new XAttribute("Turns", turns.GetValue(s)));
                        if (diameter != null && s < diameter.Length) sectionElement.Add(new XAttribute("Diameter", diameter.GetValue(s)));
                        dataElement.Add(sectionElement);
                    }

                    helicalCurveElements.Add(dataElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelicalCurve getHelicalCurveData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"HelicalCurve: Error Message:{ex.Message}");
                return new XElement("HelicalCurve", "Error");
            }
            finally
            {
                if (helicalCurve != null)
                {
                    Marshal.ReleaseComObject(helicalCurve);
                    helicalCurve = null;
                }
            }

            Console.WriteLine("Created HelicalCurve XML list");
            return helicalCurveElements;
        }
    }
}