using System.Runtime.InteropServices;
using SolidEdgePart;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using SolidEdgeGeometry;

namespace xml_data_extraction.Features
{
    internal class FE03_cutout_extractor
    {
        public static XElement Cutout_Extrude(ExtrudedCutout extrudedCutout)
        {
            XElement cutoutElements = new XElement("Cutout", new XAttribute("Type", 462094714));

            try
            {
                cutoutElements.Add(new XElement("name",extrudedCutout.Name));

                cutoutElements.Add(new XElement("type", extrudedCutout.Type));

                //var depth = extrudedCutout.Depth;
                //cutoutElements.Add(new XElement("depth", depth));

                //var extentSide = extrudedCutout.ExtentSide;
                //cutoutElements.Add(new XElement("extent_side", extentSide));

                //var extrudeType = extrudedCutout.ExtentType;
                //cutoutElements.Add(new XElement("extrude_type", extrudeType));

                cutoutElements.Add(new XElement("modelingModeType", extrudedCutout.ModelingModeType));

                var profileSide = extrudedCutout.ProfileSide;
                cutoutElements.Add(new XElement("profile_side", profileSide));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(extrudedCutout);
                cutoutElements.Add(profile_extract);

                //var profile = extrudedCutout.Profile;
                //XElement profileElement = new XElement("Profiles");

                //profileElement.Add(new XElement("profile_name", profile.Name));
                //profileElement.Add(new XElement("profile_type", profile.Type));

                //var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                //profileElement.Add(dim_extract); // Add dimensions to profile

                //cutoutElements.Add(profileElement);  // Add profile to extrusion

                //Console.WriteLine($"CUT.Depth []: {depth}");
                //Console.WriteLine($"CUT.Extent Direction []: {extentSide}");
                //Console.WriteLine($"CUT.Extent Type []: {extrudeType}");

                extrudedCutout.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side,
                                                        out double finiteDepth1);
                cutoutElements.Add(new XElement("Direction1Extent",
                                            new XElement("extent_type", extent1Type.ToString()),
                                            new XElement("extent_side", extent1Side.ToString()),
                                            new XElement("finite_depth", finiteDepth1)));

                extrudedCutout.GetDirection1Treatment(out TreatmentTypeConstants treatment1Type, out DraftSideConstants draft1Side, out double treatmentDraftAngle1,
                                                            out TreatmentCrownTypeConstants treatment1CrownType, out TreatmentCrownSideConstants treatment1CrownSide,
                                                            out TreatmentCrownCurvatureSideConstants treatment1CrownCurvatureSide, out double treatment1CrownRadiusOrOffset,
                                                            out double treatment1CrownTakeOffAngle);
                cutoutElements.Add(new XElement("Direction1Treatment",
                                            new XElement("treatment_type", treatment1Type.ToString()),
                                            new XElement("draft_side", draft1Side.ToString()),
                                            new XElement("treatment_draft_angle", treatmentDraftAngle1),
                                            new XElement("treatment_crown_type", treatment1CrownType.ToString()),
                                            new XElement("treatment_crown_side", treatment1CrownSide.ToString()),
                                            new XElement("treatment_crown_curvature_side", treatment1CrownCurvatureSide.ToString()),
                                            new XElement("treatment_crown_radius_or_offset", treatment1CrownRadiusOrOffset),
                                            new XElement("treatment_crown_take_off_angle", treatment1CrownTakeOffAngle)));

                extrudedCutout.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side,
                                                        out double finiteDepth2);
                cutoutElements.Add(new XElement("Direction2Extent",
                                            new XElement("extent_type", extent2Type.ToString()),
                                            new XElement("extent_side", extent2Side.ToString()),
                                            new XElement("finite_depth", finiteDepth2)));

                extrudedCutout.GetDirection2Treatment(out TreatmentTypeConstants treatment2Type, out DraftSideConstants draft2Side, out double treatmentDraftAngle2,
                                                            out TreatmentCrownTypeConstants treatment2CrownType, out TreatmentCrownSideConstants treatment2CrownSide,
                                                            out TreatmentCrownCurvatureSideConstants treatment2CrownCurvatureSide, out double treatment2CrownRadiusOrOffset,
                                                            out double treatment2CrownTakeOffAngle);
                cutoutElements.Add(new XElement("Direction1Treatment",
                                            new XElement("treatment_type", treatment2Type.ToString()),
                                            new XElement("draft_side", draft2Side.ToString()),
                                            new XElement("treatment_draft_angle", treatmentDraftAngle2),
                                            new XElement("treatment_crown_type", treatment2CrownType.ToString()),
                                            new XElement("treatment_crown_side", treatment2CrownSide.ToString()),
                                            new XElement("treatment_crown_curvature_side", treatment2CrownCurvatureSide.ToString()),
                                            new XElement("treatment_crown_radius_or_offset", treatment2CrownRadiusOrOffset),
                                            new XElement("treatment_crown_take_off_angle", treatment2CrownTakeOffAngle)));

                extrudedCutout.GetFromFaceOffsetData(out object fromFaceorPlane, out OffsetSideConstants fromFaceOffsetSide, out double fromFaceOffsetDistance);
                cutoutElements.Add(new XElement("FromFaceOffsetData",
                                            //new XElement("from_face_or_plane", fromFaceorPlane != null ? fromFaceorPlane.ToString() : "null"),
                                            new XElement("from_face_offset_side", fromFaceOffsetSide.ToString()),
                                            new XElement("from_face_offset_distance", fromFaceOffsetDistance)));

                extrudedCutout.GetToFaceOffsetData(out object toFaceorPlane, out OffsetSideConstants toFaceOffsetSide, out double toFaceOffsetDistance);
                cutoutElements.Add(new XElement("ToFaceOffsetData",
                                            //new XElement("to_face_or_plane", toFaceorPlane != null ? toFaceorPlane.ToString() : "null"),
                                            new XElement("to_face_offset_side", toFaceOffsetSide.ToString()),
                                            new XElement("to_face_offset_distance", toFaceOffsetDistance)));

                //FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                //var edgeTyp = extrudedCutout.Edges[EdgeTyp];
                //Console.WriteLine($"CUT.Edges []: {edgeTyp.Count}");

                //for (int e = 1; e <= edgeTyp.Count; e++)
                //{
                //    var edge = (Edge)edgeTyp.Item(e);
                //    var edgeElement = edge.Type.ToString();
                //    cutoutElements.Add(new XElement($"edgeType{e}", edgeElement));
                //}

            }

            catch (Exception ex)
            {
                Console.WriteLine($"Extruded Cutout: Error Message:{ex.Message}");
                return new XElement("ExtrudedCutout", "Error");
            }

            finally
            {
                if (extrudedCutout != null)
                {
                    Marshal.ReleaseComObject(extrudedCutout);
                    extrudedCutout = null;
                }
            }

            Console.WriteLine($"Created Extruded Cutout XML list");
            return cutoutElements;
        }

        //Helix Cutout Data Extraction
        public static XElement Helix_Cutout_Extrude(HelixCutout helixCutout)
        {
            XElement helixCutoutElements = new XElement("helixCutout", new XAttribute("Type", 1197717883));

            try
            {
                helixCutoutElements.Add(new XElement("name", helixCutout.Name));

                helixCutoutElements.Add(new XElement("type", helixCutout.Type));

                helixCutoutElements.Add(new XElement("modelingModeType", helixCutout.ModelingModeType));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(helixCutout);
                helixCutoutElements.Add(profile_extract);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"\tHelix Cutout: Error Message:{ex.Message}");
                return new XElement("HelixCutout", "Error");
            }

            finally
            {
                if (helixCutout != null)
                {
                    Marshal.ReleaseComObject(helixCutout);
                    helixCutout = null;
                }
            }
            return helixCutoutElements;
        }
    }
}
