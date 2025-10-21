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
    internal class FE10_mirror_extractor
    {
        public static XElement MirrorPart(MirrorPart mirrorPart)
        {
            XElement mirrorPartElements = new XElement("Mirror_Part", new XAttribute("Type", 1908287958));

            try
            {
                var name = mirrorPart.Name;
                mirrorPartElements.Add(new XElement("name", name));

                var type = mirrorPart.Type;
                mirrorPartElements.Add(new XElement("type", type));

                var mirrorPlane = mirrorPart.MirrorPlane;
                mirrorPartElements.Add(new XElement("mirrorPlane", mirrorPlane));
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Mirror Part: Error Message:{ex.Message}");
                return new XElement("Miror_Part", "Error");
            }

            finally
            {
                if (mirrorPart != null)
                {
                    Marshal.ReleaseComObject(mirrorPart);
                    mirrorPart = null;
                }
            }

            Console.WriteLine($"Created Mirror Part XML list");
            return mirrorPartElements;
        }
    }
}
