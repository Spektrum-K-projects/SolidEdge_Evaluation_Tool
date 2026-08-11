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
            XElement revolvedProtrusionElements = new XElement("RevolvedProtrusion",
                                                new XAttribute("Type", 462094710));

            try
            {
                try { revolvedProtrusionElements.Add(new XElement("name", revolve.Name)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedProtrusionElements.Add(new XElement("type", revolve.Type)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    FeatureStatusConstants featureStatus = revolve.GetStatusEx(out object statusDescription);
                    revolvedProtrusionElements.Add(new XElement("status", featureStatus));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Protrusion Status: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { revolvedProtrusionElements.Add(new XElement("angle", revolve.Angle)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion Angle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedProtrusionElements.Add(new XElement("extent_side", revolve.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedProtrusionElements.Add(new XElement("extent_type", revolve.ExtentType)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion ExtentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedProtrusionElements.Add(new XElement("modeling_type", revolve.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedProtrusionElements.Add(new XElement("convertToCutoutAllowed", revolve.ConvertToCutoutAllowed)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion ConvertToCutoutAllowed: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    revolve.GetDirection1Extent(out FeaturePropertyConstants extent1Type,
                                                out FeaturePropertyConstants extent1Side,
                                                out double angle1);

                    revolvedProtrusionElements.Add(
                        new XElement("Direction1Extent",
                            new XElement("extent_type", extent1Type),
                            new XElement("extent_side", extent1Side),
                            new XElement("angle", angle1)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Protrusion Direction1Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    revolve.GetDirection2Extent(out FeaturePropertyConstants extent2Type,
                                                out FeaturePropertyConstants extent2Side,
                                                out double angle2);

                    revolvedProtrusionElements.Add(
                        new XElement("Direction2Extent",
                            new XElement("extent_type", extent2Type),
                            new XElement("extent_side", extent2Side),
                            new XElement("angle", angle2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Protrusion Direction2Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { revolvedProtrusionElements.Add(new XElement("profile_side", revolve.ProfileSide)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Protrusion ProfileSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    revolve.GetDimensions(out int numDims, ref dims);
                    revolvedProtrusionElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    XElement profileExtract = GE04_getProfiles_extractor.getProfile_extract(revolve);
                    revolvedProtrusionElements.Add(profileExtract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Protrusion Profiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                Console.WriteLine("Created Revolved Protrusion XML");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Revolved Protrusion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("RevolvedProtrusion", "Error");
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
            XElement revolvedCutoutElements = new XElement("RevolvedCutout",
                                                new XAttribute("Type", 462094718));

            try
            {
                try { revolvedCutoutElements.Add(new XElement("name", revolve.Name)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedCutoutElements.Add(new XElement("type", revolve.Type)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    FeatureStatusConstants featureStatus = revolve.GetStatusEx(out object statusDescription);
                    revolvedCutoutElements.Add(new XElement("status", featureStatus));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Cutout Status: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { revolvedCutoutElements.Add(new XElement("angle", revolve.Angle)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout Angle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedCutoutElements.Add(new XElement("extent_side", revolve.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedCutoutElements.Add(new XElement("extent_type", revolve.ExtentType)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout ExtentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { revolvedCutoutElements.Add(new XElement("modeling_type", revolve.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    revolve.GetDirection1Extent(out FeaturePropertyConstants extent1Type,
                                                out FeaturePropertyConstants extent1Side,
                                                out double angle1);

                    revolvedCutoutElements.Add(
                        new XElement("Direction1Extent",
                            new XElement("extent_type", extent1Type),
                            new XElement("extent_side", extent1Side),
                            new XElement("angle", angle1)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Cutout Direction1Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    revolve.GetDirection2Extent(out FeaturePropertyConstants extent2Type,
                                                out FeaturePropertyConstants extent2Side,
                                                out double angle2);

                    revolvedCutoutElements.Add(
                        new XElement("Direction2Extent",
                            new XElement("extent_type", extent2Type),
                            new XElement("extent_side", extent2Side),
                            new XElement("angle", angle2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Cutout Direction2Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { revolvedCutoutElements.Add(new XElement("profile_side", revolve.ProfileSide)); }
                catch (Exception ex) { Console.WriteLine($"Revolved Cutout ProfileSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    revolve.GetDimensions(out int numDims, ref dims);

                    revolvedCutoutElements.Add(
                        GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Cutout GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    XElement profileExtract = GE04_getProfiles_extractor.getProfile_extract(revolve);
                    revolvedCutoutElements.Add(profileExtract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Revolved Cutout Profiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                Console.WriteLine($"Created Revolved Cutout XML");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Revolved Cutout: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
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
