using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE09_thinregion_extractor
    {
        public static XElement ThinRegion(Thin thinRegion)
        {
            XElement ThinRegionElements = new XElement("Thin_Region", new XAttribute("Type", 438630050));

            try
            {
                ThinRegionElements.Add(new XElement("name", thinRegion.Name));
                ThinRegionElements.Add(new XElement("type", thinRegion.Type));
                ThinRegionElements.Add(new XElement("modelingModeType", thinRegion.ModelingModeType));

                try
                {
                    dynamic dynThinRegion = thinRegion;
                    ThinRegionElements.Add(new XElement("status", dynThinRegion.Status));
                    ThinRegionElements.Add(new XElement("suppress", dynThinRegion.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = thinRegion.GetStatusEx(out object description);
                    ThinRegionElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    thinRegion.GetDimensions(out int numDims, ref dims);
                    ThinRegionElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    thinRegion.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    ThinRegionElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ThinRegion GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ThinRegion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Thin_Region", "Error");
            }
            finally
            {
                if (thinRegion != null)
                {
                    Marshal.ReleaseComObject(thinRegion);
                    thinRegion = null;
                }
            }

            Console.WriteLine("Created Thin Region XML list");
            return ThinRegionElements;
        }
    }
}