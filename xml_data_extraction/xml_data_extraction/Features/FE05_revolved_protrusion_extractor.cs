using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgePart;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE05_revolved_protrusion_extractor
    {
        public static XElement Revolve(RevolvedProtrusion revolve)
        {
            XElement revolveElements = new XElement("Revolve", new XAttribute("Type", 462094710));

            try
            {
                var extentSide = revolve.ExtentSide;
                revolveElements.Add(new XElement("extent_side", extentSide));
                Console.WriteLine($"REV.Extent Direction: {extentSide}");

                var extrudeType = revolve.ExtentType;
                revolveElements.Add(new XElement("extrude_type", extrudeType));
                Console.WriteLine($"REV.Extent Type: {extrudeType}");

                var profile = revolve.Profile;
                XElement profileElement = new XElement("Profiles");
                
                profileElement.Add(new XElement("profile_name", profile.Name));
                profileElement.Add(new XElement("profile_type", profile.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                profileElement.Add(dim_extract); // Add dimensions to profile

                revolveElements.Add(profileElement);  // Add profile to extrusion
                Console.WriteLine($"REV.plane: {profile.Name}");

                var modeling_mode_type = revolve.ModelingModeType;
                revolveElements.Add(new XElement("modeling_type", modeling_mode_type));
                Console.WriteLine($"REV. Modeling Type: {modeling_mode_type}");

                Marshal.ReleaseComObject(profile);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"REV: Error Message:{ex.Message}");
            }
            finally
            {
                if (revolve != null)
                {
                    Marshal.ReleaseComObject(revolve);
                    revolve = null;
                }
            }

            return revolveElements;
        }
    }
}
