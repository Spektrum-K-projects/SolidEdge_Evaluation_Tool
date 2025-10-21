using SolidEdgePart;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE02_hole_extractor
    {
        public static XElement Holes(Holes holes)
        {
            XElement holesElements = new XElement("Holes", new XAttribute("Type", 462094722));

            try
            {
                for (int i = 1; i <= holes.Count; i++)
                {
                    var holesFeat = holes.Item(i);
                    holesElements.Add(new XElement(FE02_hole_extractor.Hole((Hole)holesFeat)));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Holes: Error Message:{ex.Message}");
            }
            finally
            {
                if (holes != null)
                {
                    Marshal.ReleaseComObject(holes);
                    holes = null;
                }
            }

            return holesElements;
        }

        public static XElement Hole(Hole hole)
        {
            //SolidEdgePart.Holes holes = null;
            //SolidEdgePart.Hole hole = null;

            XElement holeElements = new XElement("Hole");

            try
            {
                //holes = (SolidEdgePart.Holes)model.Holes;

                //for (int i = 1; i <= holes.Count; i++)
                //{
                    //hole = (SolidEdgePart.Hole)holes.Item(i);
                var depth_hole = hole.Depth;
                holeElements.Add(new XElement("depth", depth_hole));

                var extentSide_hole = hole.ExtentSide;
                holeElements.Add(new XElement("extent_side", extentSide_hole));

                var extentType_hole = hole.ExtentType;
                holeElements.Add(new XElement("extent_type", extentType_hole));

                var profile_hole = hole.Profile;
                XElement profileElement = new XElement("Profiles");

                profileElement.Add(new XElement("profile_name", profile_hole.Name));
                profileElement.Add(new XElement("profile_type", profile_hole.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile_hole);
                profileElement.Add(dim_extract); // Add dimensions to profile

                holeElements.Add(profileElement);  // Add profile to extrusion

                // Console.WriteLine($"HOLE.Depth []: {depth_hole}");
                    //Console.WriteLine($"HOLE.Extent Side []: {extentSide_hole}");
                    //Console.WriteLine($"HOLE.Extent Type []: {extentType_hole}");

                GE01_dimensions_extractor.Dimension_extract(profile_hole);
                Marshal.ReleaseComObject(profile_hole);
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hole: Error Message:{ex.Message}");
            }
            finally
            {
                if (hole != null)
                {
                    Marshal.ReleaseComObject(hole);
                    hole = null;
                }
            }

            Console.WriteLine($"Created Hole XML list");
            return holeElements;
        }
    }
}
