using SolidEdgePart;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using xml_data_extraction.Properties;

namespace xml_data_extraction.Features
{
    internal class FE02_hole_extractor
    {
        public static XElement Hole(Hole hole)
        {
            //SolidEdgePart.Holes holes = null;
            //SolidEdgePart.Hole hole = null;

            XElement holeElements = new XElement("Hole", new XAttribute("Type", 462094722));

            try
            {
                //holes = (SolidEdgePart.Holes)model.Holes;

                //for (int i = 1; i <= holes.Count; i++)
                //{
                    //hole = (SolidEdgePart.Hole)holes.Item(i);
                //var depth_hole = hole.Depth;
                //holeElements.Add(new XElement("depth", depth_hole));

                var extentSide_hole = hole.ExtentSide;
                holeElements.Add(new XElement("extent_side", extentSide_hole));

                var extentType_hole = hole.ExtentType;
                holeElements.Add(new XElement("extent_type", extentType_hole));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(hole);
                holeElements.Add(profile_extract);

                holeElements.Add(PR03_hole_data_extractor.Hole_Data(hole));

                //int numProfiles = 0;
                ////Array profilesArray = null;
                //Array profilesArray = Array.CreateInstance(typeof(object), 0);

                //try
                //{
                //    hole.GetProfiles(out numProfiles, ref profilesArray);
                //    Console.WriteLine($"Number of profiles found: {numProfiles}");
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine($"Error retrieving profiles: {ex.Message}");
                //    return holeElements;
                //}

                //if (profilesArray == null || numProfiles == 0)
                //{
                //    Console.WriteLine("No valid profiles found for this extrusion.");
                //    return holeElements;
                //}

                //XElement profilesRoot = new XElement("Profiles");

                //foreach (object profileObj in profilesArray)
                //{
                //    if (profileObj == null)
                //    {
                //        Console.WriteLine("Warning: Null profile object encountered, skipping.");
                //        continue;
                //    }

                //    Profile profile = profileObj as Profile;
                //    if (profile == null)
                //    {
                //        Console.WriteLine("Warning: Unable to cast object to Profile, skipping.");
                //        continue;
                //    }

                //    XElement profileElement = new XElement("Profile");

                //    try
                //    {
                //        profileElement.Add(new XElement("profile_name", profile.Name ?? "Unnamed"));
                //        profileElement.Add(new XElement("profile_type", profile.Type.ToString()));

                //        // --- Extract geometry and dimension data ---
                //        var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                //        if (dim_extract != null)
                //            profileElement.Add(dim_extract);
                //        else
                //            Console.WriteLine($"Warning: No dimensions extracted for {profile.Name}.");

                //        var relations2d_extract = GE02_relations_extractor.Relations2d_extract(profile);
                //        if (relations2d_extract != null)
                //            profileElement.Add(relations2d_extract);

                //        var line_extract = GE03_2d_geometries_extractor.Line2d_extract(profile);
                //        if (line_extract != null)
                //            profileElement.Add(line_extract);

                //        var circle_extract = GE03_2d_geometries_extractor.Circle2d_extract(profile);
                //        if (circle_extract != null)
                //            profileElement.Add(circle_extract);

                //        var arc_extract = GE03_2d_geometries_extractor.Arc2d_extract(profile);
                //        if (arc_extract != null)
                //            profileElement.Add(arc_extract);

                //        // --- Add to profiles list ---
                //        profilesRoot.Add(profileElement);
                //    }
                //    catch (Exception innerEx)
                //    {
                //        Console.WriteLine($"Error processing profile '{profile?.Name ?? "Unknown"}': {innerEx.Message}");
                //    }
                //    finally
                //    {
                //        Marshal.ReleaseComObject(profile);
                //        profile = null;
                //    }
                //}

                //holeElements.Add(profilesRoot);

                //-------------Need to check the following the if it needs to be shelved---------
                //var profile_hole = hole.Profile;
                //XElement profileElement = new XElement("Profiles");

                //profileElement.Add(new XElement("profile_name", profile_hole.Name));
                //profileElement.Add(new XElement("profile_type", profile_hole.Type));

                //var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile_hole);
                //profileElement.Add(dim_extract); // Add dimensions to profile

                //holeElements.Add(profileElement);  // Add profile to extrusion

                ////Console.WriteLine($"HOLE.Depth []: {depth_hole}");
                ////    Console.WriteLine($"HOLE.Extent Side []: {extentSide_hole}");
                ////    Console.WriteLine($"HOLE.Extent Type []: {extentType_hole}");

                ////GE01_dimensions_extractor.Dimension_extract(profile_hole);
                ////Marshal.ReleaseComObject(profile_hole);
                //////}
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hole: Error Message:{ex.Message}");
                return new XElement("Hole", "Error");
            }
            finally
            {
                if (hole != null)
                {
                    Marshal.ReleaseComObject(hole);
                    hole = null;
                }
                //if (holes != null)
                //{
                //    Marshal.ReleaseComObject(holes);
                //    holes = null;
                //}
            }

            Console.WriteLine($"Created Hole XML list");
            return holeElements;
        }
    }
}
