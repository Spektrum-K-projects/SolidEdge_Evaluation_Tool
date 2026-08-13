using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Geometries
{
    internal class GE04_getProfiles_extractor
    {
        public static XElement getProfile_extract(object feature)
        {
            XElement profilesRoot = new XElement("Profiles");

            int numProfiles = 0;
            Array profilesArray = Array.CreateInstance(typeof(object), 0);

            try
            {
                // Feature Type Check and GetProfiles
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
                else if (feature is Rib rib)
                {
                    rib.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is HelixProtrusion helixProtrusion)
                {
                    helixProtrusion.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is HelixCutout helixCutout)
                {
                    helixCutout.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is WebNetwork webNetworkFeature)
                {
                    webNetworkFeature.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is Slot slot)
                {
                    slot.GetProfiles(out numProfiles, ref profilesArray);
                }
                else if (feature is NormalCutout normalCutout)
                {
                    normalCutout.GetProfiles(out numProfiles, ref profilesArray);
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

            if (profilesArray == null || numProfiles == 0)
            {
                Console.WriteLine("No valid profiles found for this feature.");
                profilesRoot.Add(new XElement("Info", "No profiles found."));
                return profilesRoot;
            }

            profilesRoot.Add(new XAttribute("Count", numProfiles));

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
                    if (Marshal.IsComObject(profileObj))
                    {
                        Marshal.ReleaseComObject(profileObj);
                    }
                    continue;
                }

                profilesRoot.Add(Profile_Data(profile));
            }

            Console.WriteLine($"Created Profiles XML list");
            return profilesRoot;
        }

        public static XElement Profile_Data(Profile profile)
        {
            XElement profileElement = new XElement("Profile");

            if (profile == null)
            {
                profileElement.Add(new XAttribute("Error", "Null profile"));
                return profileElement;
            }

            try
            {
                try { profileElement.Add(new XElement("profile_name", profile.Name ?? "Unnamed")); }
                catch (Exception ex) { Console.WriteLine($"Profile Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { profileElement.Add(new XElement("profile_type", profile.Type.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Profile Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    Array matrix = Array.CreateInstance(typeof(double), 0);
                    profile.GetMatrix(ref matrix);
                    double[] flat = (double[])matrix;

                    var matrixElement = new XElement("TransformMatrix");

                    if (flat.Length == 16)
                    {
                        matrixElement.Add(new XAttribute("Rows", 4), new XAttribute("Columns", 4));
                        for (int r = 0; r < 4; r++)
                        {
                            var rowElement = new XElement($"Row{r}");
                            for (int c = 0; c < 4; c++)
                            {
                                rowElement.Add(new XAttribute($"C{c}", flat[(r * 4) + c]));
                            }
                            matrixElement.Add(rowElement);
                        }
                    }
                    else
                    {
                        matrixElement.Add(new XAttribute("Length", flat.Length));
                        matrixElement.Add(new XAttribute("Values", string.Join(" ", flat)));
                    }

                    profileElement.Add(matrixElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Profile GetMatrix: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // --- Extract geometry and dimension data ---

                try
                {
                    var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                    if (dim_extract != null)
                        profileElement.Add(dim_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No dimensions extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Dimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var relations2d_extract = GE02_relations_extractor.Relations2d_extract(profile);
                    if (relations2d_extract != null)
                        profileElement.Add(relations2d_extract);
                }
                catch (Exception ex) { Console.WriteLine($"Profile Relations2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var points_extract = GE03_2d_geometries_extractor.Point2d_extract(profile);
                    if (points_extract != null)
                        profileElement.Add(points_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No Points extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Points2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var line_extract = GE03_2d_geometries_extractor.Line2d_extract(profile);
                    if (line_extract != null)
                        profileElement.Add(line_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No lines extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Line2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var circle_extract = GE03_2d_geometries_extractor.Circle2d_extract(profile);
                    if (circle_extract != null)
                        profileElement.Add(circle_extract);
                    else
                        Console.WriteLine($"\t\tWarning: No circles extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Circle2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var arc_extract = GE03_2d_geometries_extractor.Arc2d_extract(profile);
                    if (arc_extract != null)
                        profileElement.Add(arc_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No arcs extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Arc2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var ellipse_extract = GE03_2d_geometries_extractor.Ellipse2d_extract(profile);
                    if (ellipse_extract != null)
                        profileElement.Add(ellipse_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No ellipses extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Ellipse2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var ellipticalArc_extract = GE03_2d_geometries_extractor.EllipticalArc2d_extract(profile);
                    if (ellipticalArc_extract != null)
                        profileElement.Add(ellipticalArc_extract);
                    else
                        Console.WriteLine($"\t\tWarning: No Elliptical Arcs extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile EllipticalArc2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var bSpline_extract = GE03_2d_geometries_extractor.BSplineCurve2d_extract(profile);
                    if (bSpline_extract != null)
                        profileElement.Add(bSpline_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No B-Splines extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile BSplineCurve2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var conics_extract = GE03_2d_geometries_extractor.Conic2d_extract(profile);
                    if (conics_extract != null)
                        profileElement.Add(conics_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No Conics extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Conic2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var holes2d_extract = GE03_2d_geometries_extractor.Hole2d_extract(profile);
                    if (holes2d_extract != null)
                        profileElement.Add(holes2d_extract);
                    else
                        Console.WriteLine($"\t\t Warning: No Holes2d extracted for {profile.Name}.");
                }
                catch (Exception ex) { Console.WriteLine($"Profile Hole2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var boundaries_extract = GE03_2d_geometries_extractor.Boundaries2d_extract(profile);
                    if (boundaries_extract != null)
                        profileElement.Add(boundaries_extract);
                }
                catch (Exception ex) { Console.WriteLine($"Profile Boundaries2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var rectPatterns_extract = GE03_2d_geometries_extractor.RectangularPatterns2d_extract(profile);
                    if (rectPatterns_extract != null)
                        profileElement.Add(rectPatterns_extract);
                }
                catch (Exception ex) { Console.WriteLine($"Profile RectangularPatterns2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var circPatterns_extract = GE03_2d_geometries_extractor.CircularPatterns2d_extract(profile);
                    if (circPatterns_extract != null)
                        profileElement.Add(circPatterns_extract);
                }
                catch (Exception ex) { Console.WriteLine($"Profile CircularPatterns2d: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                // --- Add additional Geometries here ---
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Profile: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (profile != null)
                {
                    Marshal.ReleaseComObject(profile);
                    profile = null;
                }
            }

            Console.WriteLine($"\t Created Profile XML list");
            return profileElement;
        }
    }
}