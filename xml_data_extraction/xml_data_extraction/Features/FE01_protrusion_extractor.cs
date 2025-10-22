using SolidEdgeGeometry;
using SolidEdgePart;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE01_protrusion_extractor
    {
        //Extruded Protrusion Data Extraction
        public static XElement Protrusion_Extrude(ExtrudedProtrusion extrudedProtrusion)
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

                protrusionElements.Add(new XElement("name", extrudedProtrusion.Name));

                protrusionElements.Add(new XElement("type", extrudedProtrusion.Type));

                var depth = extrudedProtrusion.Depth;
                protrusionElements.Add(new XElement("depth", depth));

                var extrudeSide = extrudedProtrusion.ExtentSide;
                protrusionElements.Add(new XElement("extent_side", extrudeSide));

                var extrudeType = extrudedProtrusion.ExtentType;
                protrusionElements.Add(new XElement("extrude_type", extrudeType));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(extrudedProtrusion);
                protrusionElements.Add(profile_extract);

                extrudedProtrusion.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side, 
                                                        out double finiteDepth1);
                protrusionElements.Add(new XElement("Direction1Extent",
                                            new XElement("extent_type", extent1Type.ToString()),
                                            new XElement("extent_side", extent1Side.ToString()),
                                            new XElement("finite_depth", finiteDepth1)));

                extrudedProtrusion.GetDirection1Treatment(out TreatmentTypeConstants treatment1Type, out DraftSideConstants draft1Side, out double treatmentDraftAngle1,
                                                            out TreatmentCrownTypeConstants treatment1CrownType, out TreatmentCrownSideConstants treatment1CrownSide,
                                                            out TreatmentCrownCurvatureSideConstants treatment1CrownCurvatureSide, out double treatment1CrownRadiusOrOffset,
                                                            out double treatment1CrownTakeOffAngle);
                protrusionElements.Add(new XElement("Direction1Treatment",
                                            new XElement("treatment_type", treatment1Type.ToString()),
                                            new XElement("draft_side", draft1Side.ToString()),
                                            new XElement("treatment_draft_angle", treatmentDraftAngle1),
                                            new XElement("treatment_crown_type", treatment1CrownType.ToString()),
                                            new XElement("treatment_crown_side", treatment1CrownSide.ToString()),
                                            new XElement("treatment_crown_curvature_side", treatment1CrownCurvatureSide.ToString()),
                                            new XElement("treatment_crown_radius_or_offset", treatment1CrownRadiusOrOffset),
                                            new XElement("treatment_crown_take_off_angle", treatment1CrownTakeOffAngle)));

                extrudedProtrusion.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side,
                                                        out double finiteDepth2);
                protrusionElements.Add(new XElement("Direction2Extent",
                                            new XElement("extent_type", extent2Type.ToString()),
                                            new XElement("extent_side", extent2Side.ToString()),
                                            new XElement("finite_depth", finiteDepth2)));

                extrudedProtrusion.GetDirection2Treatment(out TreatmentTypeConstants treatment2Type, out DraftSideConstants draft2Side, out double treatmentDraftAngle2,
                                                            out TreatmentCrownTypeConstants treatment2CrownType, out TreatmentCrownSideConstants treatment2CrownSide,
                                                            out TreatmentCrownCurvatureSideConstants treatment2CrownCurvatureSide, out double treatment2CrownRadiusOrOffset,
                                                            out double treatment2CrownTakeOffAngle);
                protrusionElements.Add(new XElement("Direction2Treatment",
                                            new XElement("treatment_type", treatment2Type.ToString()),
                                            new XElement("draft_side", draft2Side.ToString()),
                                            new XElement("treatment_draft_angle", treatmentDraftAngle2),
                                            new XElement("treatment_crown_type", treatment2CrownType.ToString()),
                                            new XElement("treatment_crown_side", treatment2CrownSide.ToString()),
                                            new XElement("treatment_crown_curvature_side", treatment2CrownCurvatureSide.ToString()),
                                            new XElement("treatment_crown_radius_or_offset", treatment2CrownRadiusOrOffset),
                                            new XElement("treatment_crown_take_off_angle", treatment2CrownTakeOffAngle)));

                extrudedProtrusion.GetFromFaceOffsetData(out object fromFaceorPlane, out OffsetSideConstants fromFaceOffsetSide,out double fromFaceOffsetDistance);
                protrusionElements.Add(new XElement("FromFaceOffsetData",
                                            //new XElement("from_face_or_plane", fromFaceorPlane != null ? fromFaceorPlane.ToString() : "null"),
                                            new XElement("from_face_offset_side", fromFaceOffsetSide.ToString()),
                                            new XElement("from_face_offset_distance", fromFaceOffsetDistance)));

                extrudedProtrusion.GetToFaceOffsetData(out object toFaceorPlane, out OffsetSideConstants toFaceOffsetSide, out double toFaceOffsetDistance);
                protrusionElements.Add(new XElement("ToFaceOffsetData",
                                            //new XElement("to_face_or_plane", toFaceorPlane != null ? toFaceorPlane.ToString() : "null"),
                                            new XElement("to_face_offset_side", toFaceOffsetSide.ToString()),
                                            new XElement("to_face_offset_distance", toFaceOffsetDistance)));

                FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                var edgeTyp = extrudedProtrusion.Edges[EdgeTyp];
                Console.WriteLine($"Extrude.Edges []: {edgeTyp.Count}");

                for (int e = 1; e <= edgeTyp.Count; e++)
                {
                    var edge = (Edge)edgeTyp.Item(e);
                    var edgeElement = edge.Type.ToString();
                    protrusionElements.Add(new XElement($"edgeType{e}", edgeElement));
                }

                Marshal.ReleaseComObject(edgeTyp);



                //int numProfiles = 0;
                ////Array profilesArray = null;
                //Array profilesArray = Array.CreateInstance(typeof(object), 0);

                //try
                //{
                //    extrudedProtrusion.GetProfiles(out numProfiles, ref profilesArray);
                //    Console.WriteLine($"Number of profiles found: {numProfiles}");
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine($"Error retrieving profiles: {ex.Message}");
                //    return protrusionElements;
                //}

                //if (profilesArray == null || numProfiles == 0)
                //{
                //    Console.WriteLine("No valid profiles found for this extrusion.");
                //    return protrusionElements;
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

                //protrusionElements.Add(profilesRoot);


                //Profile profile = extrudedProtrusion.Profile;

                //XElement profileElement = new XElement("Profiles");
                ////new XElement("Profile",  //create profile element
                //profileElement.Add(new XElement("profile_name", profile.Name));
                //profileElement.Add(new XElement("profile_type", profile.Type));

                //var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                //profileElement.Add(dim_extract); // Add dimensions to profile

                //var relations2d_extract = GE02_relations_extractor.Relations2d_extract(profile);
                //profileElement.Add(relations2d_extract);

                //var line_extract = GE03_2d_geometries_extractor.Line2d_extract(profile);
                //profileElement.Add(line_extract);

                //var circle_extract = GE03_2d_geometries_extractor.Circle2d_extract(profile);
                //profileElement.Add(circle_extract);

                //var arc_extract = GE03_2d_geometries_extractor.Arc2d_extract(profile);
                //profileElement.Add(arc_extract);

                //protrusionElements.Add(profileElement);  // Add profile to extrusion

                ////var dim_extract = Dimensions_extract.Dimension_extract(profile);

                ////protrusionElements.Add(new XElement("Profile", 
                ////                        (new XElement("Profile Name", profile.Name.ToString())), (new XElement("Profile Type", profile.Type.ToString())),
                ////                        (new XElement(dim_extract))));


                ////Console.WriteLine($"PRO.Depth []: {depth}");
                ////Console.WriteLine($"PRO.Extent Direction []: {extentSide}");
                ////Console.WriteLine($"PRO.Extent Type []: {extrudeType}");
                ////Console.WriteLine($"PRO.plane []: {profile.Name}");
                ////Console.WriteLine($"PRO.type []: {profile.type}");

                //Marshal.ReleaseComObject(profile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Protrusion: Error Message:{ex.Message}");
                return new XElement("Extrusion", "Error");
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

            Console.WriteLine($"Created Protrusion XML list");
            return protrusionElements;
        }

        //Helix Protrusion Data Extraction
        public static XElement Helix_Protrusion_Extrude(HelixProtrusion helixProtrusion)
        {
            XElement helixProtrusionElements = new XElement("helixExtrusion", new XAttribute("Type", -1204891230));

            try
            {
                helixProtrusionElements.Add(new XElement("name", helixProtrusion.Name));

                helixProtrusionElements.Add(new XElement("type", helixProtrusion.Type));

                helixProtrusionElements.Add(new XElement("modelingModeType", helixProtrusion.ModelingModeType));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(helixProtrusion);
                helixProtrusionElements.Add(profile_extract);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Helix Protrusion: Error Message:{ex.Message}");
                return new XElement("HelixExtrusion", "Error");
            }
            finally
            {
                if (helixProtrusion != null)
                {
                    Marshal.ReleaseComObject(helixProtrusion);
                    helixProtrusion = null;
                }
            }

            return helixProtrusionElements;
        }
    }
}
