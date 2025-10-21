using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgePart;
using xml_data_extraction.Geometries;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace xml_data_extraction.Features
{
    internal class FE01_extruded_protrusion_extractor
    {
        public static XElement Protrusion(ExtrudedProtrusion extrudedProtrusion)
        {
            //SolidEdgePart.ExtrudedProtrusions extrudedProtrusions = null;
            //SolidEdgePart.ExtrudedProtrusion extrudedProtrusion = null;

            //var protrusionElements = new Dictionary<string, object>();

            XElement protrusionElements = new XElement("Extrusion", new XAttribute("Type", 462094706));

            try
            {
                //extrudedProtrusions = (SolidEdgePart.ExtrudedProtrusions)model.ExtrudedProtrusions;
                //for (int j = 1; j <= extrudedProtrusions.Count; j++)
                //{
                    //extrudedProtrusion = (SolidEdgePart.ExtrudedProtrusion)extrudedProtrusions.Item(j);
                var depth = extrudedProtrusion.Depth;
                protrusionElements.Add(new XElement("depth", depth));

                var extentSide = extrudedProtrusion.ExtentSide;
                protrusionElements.Add(new XElement("extent_side", extentSide));

                var extrudeType = extrudedProtrusion.ExtentType;
                protrusionElements.Add(new XElement("extrude_type", extrudeType));

                Profile profile = extrudedProtrusion.Profile;

                XElement profileElement = new XElement("Profiles");
                //new XElement("Profile",  //create profile element
                profileElement.Add(new XElement("profile_name", profile.Name));
                profileElement.Add(new XElement("profile_type", profile.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                profileElement.Add(dim_extract); // Add dimensions to profile

                var relations2d_extract = GE02_relations_extractor.Relations2d_extract(profile);
                profileElement.Add(relations2d_extract);

                var line_extract = GE03_2d_geometries_extractor.Line2d_extract(profile);
                profileElement.Add(line_extract);

                var circle_extract = GE03_2d_geometries_extractor.Circle2d_extract(profile);
                profileElement.Add(circle_extract);

                var arc_extract = GE03_2d_geometries_extractor.Arc2d_extract(profile);
                profileElement.Add(arc_extract);

                protrusionElements.Add(profileElement);  // Add profile to extrusion

                //var dim_extract = Dimensions_extract.Dimension_extract(profile);

                //protrusionElements.Add(new XElement("Profile", 
                //                        (new XElement("Profile Name", profile.Name.ToString())), (new XElement("Profile Type", profile.Type.ToString())),
                //                        (new XElement(dim_extract))));


                //Console.WriteLine($"PRO.Depth []: {depth}");
                //Console.WriteLine($"PRO.Extent Direction []: {extentSide}");
                //Console.WriteLine($"PRO.Extent Type []: {extrudeType}");
                //Console.WriteLine($"PRO.plane []: {profile.Name}");
                //Console.WriteLine($"PRO.type []: {profile.type}");

                Marshal.ReleaseComObject(profile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Protrusion: Error Message:{ex.Message}");
                return new XElement("Extrusion");
            }
            finally
            {
                if (extrudedProtrusion != null)
                {
                    Marshal.ReleaseComObject(extrudedProtrusion);
                    extrudedProtrusion = null;
                }
                //if (extrudedProtrusions != null)
                //{
                //    Marshal.ReleaseComObject(extrudedProtrusions);
                //    extrudedProtrusions = null;
                //}
            }

            //XElement xmlProtrusion = new XElement("Protrusion", protrusionElements.Select(kv => new XElement(
            //                              kv.Key, kv.Value?.ToString() ?? System.String.Empty)));

            Console.WriteLine($"Created Protrusion XML list");
            //return xmlProtrusion;

            return protrusionElements;
;

        }
    }
}
