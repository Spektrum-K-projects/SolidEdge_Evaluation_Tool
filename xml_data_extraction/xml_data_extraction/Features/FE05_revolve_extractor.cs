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
    internal class FE05_revolve_extractor
    {
        public static XElement Revolved_protrusion(RevolvedProtrusion revolve)
        {
            XElement revolvedProtrusionElements = new XElement("RevolvedProtrusion", new XAttribute("Type", 462094710));

            try
            {
                var extentSide = revolve.ExtentSide;
                revolvedProtrusionElements.Add(new XElement("extent_side", extentSide));
                Console.WriteLine($"REV.Extent Direction: {extentSide}");

                var extrudeType = revolve.ExtentType;
                revolvedProtrusionElements.Add(new XElement("extrude_type", extrudeType));
                Console.WriteLine($"REV.Extent Type: {extrudeType}");

                var profile = revolve.Profile;
                XElement profileElement = new XElement("Profiles");
                
                profileElement.Add(new XElement("profile_name", profile.Name));
                profileElement.Add(new XElement("profile_type", profile.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                profileElement.Add(dim_extract); // Add dimensions to profile

                revolvedProtrusionElements.Add(profileElement);  // Add profile to extrusion
                Console.WriteLine($"REV.plane: {profile.Name}");

                var modeling_mode_type = revolve.ModelingModeType;
                revolvedProtrusionElements.Add(new XElement("modeling_type", modeling_mode_type));
                Console.WriteLine($"REV. Modeling Type: {modeling_mode_type}");

                Marshal.ReleaseComObject(profile);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Revolved Extrusion: Error Message:{ex.Message}");
            }
            finally
            {
                if (revolve != null)
                {
                    Marshal.ReleaseComObject(revolve);
                    revolve = null;
                }
            }

            return revolvedProtrusionElements;
        }

        public static XElement Revolved_Cutout(RevolvedCutout revolve)
        {
            XElement revolvedCutoutElements = new XElement("RevolvedCutout", new XAttribute("Type", 462094718));

            try
            {
                var angle = revolve.Angle;
                revolvedCutoutElements.Add(new XElement("angle", angle));
                Console.WriteLine($"Rev. Cutout Angle: {angle}");

                var axis = revolve.Axis;
                revolvedCutoutElements.Add(new XElement("axis", axis));
                Console.WriteLine($"Rev. Cutout Axis: {axis}");

                var extentSide = revolve.ExtentSide;
                revolvedCutoutElements.Add(new XElement("extent_side", extentSide));
                Console.WriteLine($"Rev. Cutout Extent Direction: {extentSide}");

                var extrudeType = revolve.ExtentType;
                revolvedCutoutElements.Add(new XElement("extrude_type", extrudeType));
                Console.WriteLine($"Rev. Cutout Extent Type: {extrudeType}");

                var profile = revolve.Profile;
                XElement profileElement = new XElement("Profiles");

                profileElement.Add(new XElement("profile_name", profile.Name));
                profileElement.Add(new XElement("profile_type", profile.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                profileElement.Add(dim_extract); // Add dimensions to profile

                revolvedCutoutElements.Add(profileElement);  // Add profile to extrusion
                Console.WriteLine($"Rev. Cutout plane: {profile.Name}");

                var profileSide = revolve.ProfileSide;
                revolvedCutoutElements.Add(new XElement("profile_side", profileSide));
                Console.WriteLine($"Rev. Cutout Angle: {profileSide}");

                var modeling_mode_type = revolve.ModelingModeType;
                revolvedCutoutElements.Add(new XElement("modeling_type", modeling_mode_type));
                Console.WriteLine($"Rev. Cutout Modeling Type: {modeling_mode_type}");

                Marshal.ReleaseComObject(profile);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Revolved Cutout: Error Message:{ex.Message}");
            }

            finally
            {
                if (revolve != null)
                {
                    Marshal.ReleaseComObject(revolve);
                    revolve = null;
                }
            }

            return revolvedCutoutElements;
        }
    }
}
