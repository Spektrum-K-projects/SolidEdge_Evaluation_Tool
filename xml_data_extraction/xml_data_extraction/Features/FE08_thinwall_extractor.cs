using SolidEdgePart;
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
                var name = thinwall.Name;
                thinwallElements.Add(new XElement("name", name));

                var type = thinwall.Type;
                thinwallElements.Add(new XElement("type", type));

                var thickness = thinwall.Thickness;
                thinwallElements.Add(new XElement("thickness", thickness));

                var thicknessSide = thinwall.ThicknessSide;
                thinwallElements.Add(new XElement("thicknessside", thicknessSide));

                var modelingModeType = thinwall.ModelingModeType;
                thinwallElements.Add(new XElement("modelingModeType", modelingModeType));
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Thin Wall: Error Message:{ex.Message}");
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

            Console.WriteLine($"Created Thin Wall XML list");
            return thinwallElements;
        }
    }
}
