using SolidEdgePart;
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
                var name = thinRegion.Name;
                ThinRegionElements.Add(new XElement("name", name));

                var type = thinRegion.Type;
                ThinRegionElements.Add(new XElement("type", type));

                var modelingModeType = thinRegion.ModelingModeType;
                ThinRegionElements.Add(new XElement("modelingModeType", modelingModeType));
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Thin Region: Error Message:{ex.Message}");
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

            Console.WriteLine($"Created Thin Region XML list");
            return ThinRegionElements;
        }
    }
}
