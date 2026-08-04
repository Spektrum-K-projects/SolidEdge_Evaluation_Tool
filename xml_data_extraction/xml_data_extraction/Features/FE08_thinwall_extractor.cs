using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE08_thinwall_extractor
    {
        public static XElement ThinWall(Thinwall thinwall)
        {
            XElement thinwallElements = new XElement("Thin_Wall", new XAttribute("Type", 462094734));

            try
            {
                thinwallElements.Add(new XElement("name", thinwall.Name));
                thinwallElements.Add(new XElement("type", thinwall.Type));
                thinwallElements.Add(new XElement("thickness", thinwall.Thickness));
                thinwallElements.Add(new XElement("thicknessside", thinwall.ThicknessSide));
                thinwallElements.Add(new XElement("modelingModeType", thinwall.ModelingModeType));

                try
                {
                    dynamic dynThinwall = thinwall;
                    thinwallElements.Add(new XElement("status", dynThinwall.Status));
                    thinwallElements.Add(new XElement("suppress", dynThinwall.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = thinwall.GetStatusEx(out object description);
                    thinwallElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    thinwall.GetDimensions(out int numDims, ref dims);
                    thinwallElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    thinwall.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    thinwallElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array uniqueThicknesses = Array.CreateInstance(typeof(double), 0);
                    thinwall.GetUniqueThicknesses(out int uniqueCount, ref uniqueThicknesses);

                    if (uniqueThicknesses.Length > 0)
                    {
                        var values = new double[uniqueThicknesses.Length];
                        for (int i = 0; i < uniqueThicknesses.Length; i++)
                            values[i] = Convert.ToDouble(uniqueThicknesses.GetValue(i));

                        thinwallElements.Add(new XElement("UniqueThicknesses",
                            new XAttribute("Count", uniqueCount),
                            new XAttribute("values", string.Join(" ", values))));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinWall GetUniqueThicknesses: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ThinWall: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Thin_Wall", "Error");
            }
            finally
            {
                if (thinwall != null)
                {
                    Marshal.ReleaseComObject(thinwall);
                    thinwall = null;
                }
            }

            Console.WriteLine("Created Thin Wall XML list");
            return thinwallElements;
        }
    }
}