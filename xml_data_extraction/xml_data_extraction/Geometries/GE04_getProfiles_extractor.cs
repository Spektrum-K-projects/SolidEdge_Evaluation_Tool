using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace xml_data_extraction.Geometries
{
    internal class GE04_getProfiles_extractor
    {
        public static XElement Profile_extract(object feature)
        {
            //XElement profileElements = new XElement("Profile");

            XElement profilesRoot = new XElement("Profiles");

            int numProfiles = 0;
            //Array profilesArray = null;
            Array profilesArray = Array.CreateInstance(typeof(object), 0);

            try
            {
                // Check which feature type we have
                if (feature is ExtrudedProtrusion extrudedProtrusion)
                {
                    extrudedProtrusion.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is ExtrudedCutout extrudedCutout)
                {
                    extrudedCutout.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is RevolvedProtrusion revolvedProtrusion)
                {
                    revolvedProtrusion.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is RevolvedCutout revolvedCutout)
                {
                    revolvedCutout.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is Hole hole)
                {
                    hole.GetProfiles(out numProfiles, ref profilesArray);
                }
                else
                {
                    profilesRoot.Add(new XElement("Error", "Unsupported feature type for GetProfiles()"));
                    return profilesRoot;
                }

                Console.WriteLine($"Number of profiles found: {numProfiles}");
            }
            catch (Exception ex)
            {
                profilesRoot.Add(new XElement("Error", $"GetProfiles failed: {ex.Message}"));
                return profilesRoot;
            }

            //try
            //{
            //    profiles.GetProfiles(out numProfiles, ref profilesArray);
            //    Console.WriteLine($"Number of profiles found: {numProfiles}");
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"Error retrieving profiles: {ex.Message}");
            //    profilesRoot.Add(new XElement("Error", $"Failed to retrieve profiles: {ex.Message}"));
            //    return profilesRoot;
            //}

            if (profilesArray == null || numProfiles == 0)
            {
                Console.WriteLine("No valid profiles found for this extrusion.");
                profilesRoot.Add(new XElement("Info", "No profiles found."));
                return profilesRoot;
            }

            foreach (object profileObj in profilesArray)
            {
                if (profileObj == null)
                {
                    Console.WriteLine("Warning: Null profile object encountered, skipping.");
                    continue;
                }

                Profile profile = profileObj as Profile;
                if (profile == null)
                {
                    Console.WriteLine("Warning: Unable to cast object to Profile, skipping.");
                    continue;
                }

                XElement profileElement = new XElement("Profile");

                try
                {
                    profileElement.Add(new XElement("profile_name", profile.Name ?? "Unnamed"));
                    profileElement.Add(new XElement("profile_type", profile.Type.ToString()));

                    // --- Extract geometry and dimension data ---
                    var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                    if (dim_extract != null)
                        profileElement.Add(dim_extract);
                    else
                        Console.WriteLine($"Warning: No dimensions extracted for {profile.Name}.");

                    var relations2d_extract = GE02_relations_extractor.Relations2d_extract(profile);
                    if (relations2d_extract != null)
                        profileElement.Add(relations2d_extract);

                    var points_extract = GE03_2d_geometries_extractor.Point2d_extract(profile);
                    if (points_extract != null)
                        profileElement.Add(points_extract);
                    else
                        Console.WriteLine($"Warning: No Points extracted for {profile.Name}.");

                    var line_extract = GE03_2d_geometries_extractor.Line2d_extract(profile);
                    if (line_extract != null)
                        profileElement.Add(line_extract);
                    else
                        Console.WriteLine($"Warning: No lines extracted for {profile.Name}.");

                    var circle_extract = GE03_2d_geometries_extractor.Circle2d_extract(profile);
                    if (circle_extract != null)
                        profileElement.Add(circle_extract);
                    else
                        Console.WriteLine($"Warning: No circles extracted for {profile.Name}.");

                    var arc_extract = GE03_2d_geometries_extractor.Arc2d_extract(profile);
                    if (arc_extract != null)
                        profileElement.Add(arc_extract);
                    else
                        Console.WriteLine($"Warning: No arcs extracted for {profile.Name}.");

                    var ellipse_extract = GE03_2d_geometries_extractor.Ellipse2d_extract(profile);
                    if (ellipse_extract != null)
                        profileElement.Add(ellipse_extract);
                    else
                        Console.WriteLine($"Warning: No ellipses extracted for {profile.Name}.");

                    var ellipticalArc_extract = GE03_2d_geometries_extractor.EllipticalArc2d_extract(profile);
                    if (ellipticalArc_extract != null)
                        profileElement.Add(ellipticalArc_extract);
                    else
                        Console.WriteLine($"Warning: No Elliptical Arcs extracted for {profile.Name}.");

                    var bSpline_extract = GE03_2d_geometries_extractor.BSplineCurve2d_extract(profile);
                    if (bSpline_extract != null)
                        profileElement.Add(bSpline_extract);
                    else
                        Console.WriteLine($"Warning: No B-Splines extracted for {profile.Name}.");

                    var conics_extract = GE03_2d_geometries_extractor.Conic2d_extract(profile);
                    if (conics_extract != null)
                        profileElement.Add(conics_extract);
                    else
                        Console.WriteLine($"Warning: No Conics extracted for {profile.Name}.");

                    // --- Add additional Geometries here ---

                    profilesRoot.Add(profileElement);   // Annexing the data to the Profiles list
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Error processing profile '{profile?.Name ?? "Unknown"}': {innerEx.Message}");
                }
                finally
                {
                    Marshal.ReleaseComObject(profile);
                    profile = null;
                }
            }

            Console.WriteLine($"Created Profiles XML list");
            return profilesRoot;
        }
    }
}
