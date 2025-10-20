using System;
using System.Runtime.InteropServices;
using SolidEdgePart;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE03_cutout_extractor
    {
        public static XElement Cutout(ExtrudedCutout extrudedCutout)
        {
            XElement cutoutElements = new XElement("Cutout", new XAttribute("Type", 462094714));

            try
            {
                var depth = extrudedCutout.Depth;
                cutoutElements.Add(new XElement("depth", depth));

                var extentSide = extrudedCutout.ExtentSide;
                cutoutElements.Add(new XElement("extent_side", extentSide));

                var extrudeType = extrudedCutout.ExtentType;
                cutoutElements.Add(new XElement("extrude_type", extrudeType));

                var profile = extrudedCutout.Profile;
                XElement profileElement = new XElement("Profiles");
                
                profileElement.Add(new XElement("profile_name", profile.Name));
                profileElement.Add(new XElement("profile_type", profile.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                profileElement.Add(dim_extract); // Add dimensions to profile

                cutoutElements.Add(profileElement);  // Add profile to extrusion

                Console.WriteLine($"CUT.Depth []: {depth}");
                Console.WriteLine($"CUT.Extent Direction []: {extentSide}");
                Console.WriteLine($"CUT.Extent Type []: {extrudeType}");
                Console.WriteLine($"CUT.plane []: {profile.Name}");
                Console.WriteLine($"CUT.type []: {profile.type}");

                Marshal.ReleaseComObject(profile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cutout: Error Message:{ex.Message}");
            }
            finally
            {
                if (extrudedCutout != null)
                {
                    Marshal.ReleaseComObject(extrudedCutout);
                    extrudedCutout = null;
                }
            }

            return cutoutElements;
        }
    }
}
