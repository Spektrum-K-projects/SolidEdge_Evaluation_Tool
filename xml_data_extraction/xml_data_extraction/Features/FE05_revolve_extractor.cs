using System.Runtime.InteropServices;
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
                revolvedProtrusionElements.Add(new XElement("name", revolve.Name));

                revolvedProtrusionElements.Add(new XElement("type", revolve.Type));

                revolvedProtrusionElements.Add(new XElement("angle", revolve.Angle));

                var extentSide = revolve.ExtentSide;
                revolvedProtrusionElements.Add(new XElement("extent_side", extentSide));

                var extrudeType = revolve.ExtentType;
                revolvedProtrusionElements.Add(new XElement("extrude_type", extrudeType));

                var modeling_mode_type = revolve.ModelingModeType;
                revolvedProtrusionElements.Add(new XElement("modeling_type", modeling_mode_type));

                revolvedProtrusionElements.Add(new XElement("convertToCutoutAllowed", revolve.ConvertToCutoutAllowed));

                revolve.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side,
                                        out double angle1);
                revolvedProtrusionElements.Add(new XElement("Direction1Extent",
                                            new XElement("extent_type", extent1Type.ToString()),
                                            new XElement("extent_side", extent1Side.ToString()),
                                            new XElement("angle", angle1)));

                revolve.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side,
                                                        out double angle2);
                revolvedProtrusionElements.Add(new XElement("Direction2Extent",
                                            new XElement("extent_type", extent2Type.ToString()),
                                            new XElement("extent_side", extent2Side.ToString()),
                                            new XElement("angle", angle2)));

                revolvedProtrusionElements.Add(new XElement("profile_side", revolve.ProfileSide));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(revolve);
                revolvedProtrusionElements.Add(profile_extract);
                                
                //var profile = revolve.Profile;
                //XElement profileElement = new XElement("Profiles");

                //profileElement.Add(new XElement("profile_name", profile.Name));
                //profileElement.Add(new XElement("profile_type", profile.Type));

                //var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                //profileElement.Add(dim_extract); // Add dimensions to profile

                //revolvedProtrusionElements.Add(profileElement);  // Add profile to extrusion
                //Console.WriteLine($"REV.plane: {profile.Name}");

                //Marshal.ReleaseComObject(profile);

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

                revolvedCutoutElements.Add(new XElement("name", revolve.Name));

                revolvedCutoutElements.Add(new XElement("type", revolve.Type));

                var angle = revolve.Angle;
                revolvedCutoutElements.Add(new XElement("angle", angle));
                //Console.WriteLine($"Rev. Cutout Angle: {angle}");

                var axis = revolve.Axis;
                //revolvedCutoutElements.Add(new XElement("axis", axis));
                //Console.WriteLine($"Rev. Cutout Axis: {axis}");

                var extentSide = revolve.ExtentSide;
                revolvedCutoutElements.Add(new XElement("extent_side", extentSide));
                //Console.WriteLine($"Rev. Cutout Extent Direction: {extentSide}");

                var extrudeType = revolve.ExtentType;
                revolvedCutoutElements.Add(new XElement("extrude_type", extrudeType));
                //Console.WriteLine($"Rev. Cutout Extent Type: {extrudeType}");

                revolvedCutoutElements.Add(new XElement("modelingModeType", revolve.ModelingModeType));

                revolve.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side,
                                        out double angle1);
                revolvedCutoutElements.Add(new XElement("Direction1Extent",
                                            new XElement("extent_type", extent1Type.ToString()),
                                            new XElement("extent_side", extent1Side.ToString()),
                                            new XElement("angle", angle1)));

                revolve.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side,
                                                        out double angle2);
                revolvedCutoutElements.Add(new XElement("Direction2Extent",
                                            new XElement("extent_type", extent2Type.ToString()),
                                            new XElement("extent_side", extent2Side.ToString()),
                                            new XElement("angle", angle2)));

                revolvedCutoutElements.Add(new XElement("profileSide", revolve.ProfileSide));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(revolve);
                revolvedCutoutElements.Add(profile_extract);

                //var profile = revolve.Profile;
                //XElement profileElement = new XElement("Profiles");

                //profileElement.Add(new XElement("profile_name", profile.Name));
                //profileElement.Add(new XElement("profile_type", profile.Type));

                //var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                //profileElement.Add(dim_extract); // Add dimensions to profile

                //revolvedCutoutElements.Add(profileElement);  // Add profile to extrusion
                //Console.WriteLine($"Rev. Cutout plane: {profile.Name}");

                var profileSide = revolve.ProfileSide;
                revolvedCutoutElements.Add(new XElement("profile_side", profileSide));
                Console.WriteLine($"Rev. Cutout Angle: {profileSide}");

                var modeling_mode_type = revolve.ModelingModeType;
                revolvedCutoutElements.Add(new XElement("modeling_type", modeling_mode_type));
                Console.WriteLine($"Rev. Cutout Modeling Type: {modeling_mode_type}");

                //Marshal.ReleaseComObject(profile);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Revolved Cutout: Error Message:{ex.Message}");
                return new XElement("RevolvedCutout", "Error");
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
