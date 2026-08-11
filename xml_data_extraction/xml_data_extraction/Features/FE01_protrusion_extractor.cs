using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    internal class FE01_protrusion_extractor
    {
        //Extruded Protrusion Data Extraction
        public static XElement Protrusion_Extrude(ExtrudedProtrusion extrudedProtrusion)
        {
            XElement protrusionElements = new XElement("Extrusion", new XAttribute("Type", 462094706));

            try
            {
                try { protrusionElements.Add(new XElement("name", extrudedProtrusion.Name)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("type", extrudedProtrusion.Type)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("modelingModeType", extrudedProtrusion.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("showDimensions", extrudedProtrusion.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("visible", extrudedProtrusion.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("depth", extrudedProtrusion.Depth)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion Depth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("extent_side", extrudedProtrusion.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("extrude_type", extrudedProtrusion.ExtentType)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion ExtentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(new XElement("profile_side", extrudedProtrusion.ProfileSide)); }
                catch (Exception ex) { Console.WriteLine($"Extrusion ProfileSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(extrudedProtrusion);
                    protrusionElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedProtrusion.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side, out double finiteDepth1);
                    protrusionElements.Add(new XElement("Direction1Extent",
                        new XElement("extent_type", extent1Type.ToString()),
                        new XElement("extent_side", extent1Side.ToString()),
                        new XElement("finite_depth", finiteDepth1)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetDirection1Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
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
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetDirection1Treatment: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedProtrusion.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side, out double finiteDepth2);
                    protrusionElements.Add(new XElement("Direction2Extent",
                        new XElement("extent_type", extent2Type.ToString()),
                        new XElement("extent_side", extent2Side.ToString()),
                        new XElement("finite_depth", finiteDepth2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetDirection2Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedProtrusion.GetDirection2Treatment(out TreatmentTypeConstants treatment2Type, out DraftSideConstants draft2Side, out double treatmentDraftAngle2,
                        out TreatmentCrownTypeConstants treatment2CrownType, out TreatmentCrownSideConstants treatment2CrownSide,
                        out TreatmentCrownCurvatureSideConstants treatment2CrownCurvatureSide, out double treatment2CrownRadiusOrOffset,
                        out double treatment2CrownTakeOffAngle);
                    protrusionElements.Add(new XElement("Direction2Treatment",
                        new XElement("treatment_type", treatment2Type.ToString()),
                        new XElement("draft_side", draft2Side.ToString()),
                        new XElement("treatment_crown_type", treatment2CrownType.ToString()),
                        new XElement("treatment_crown_side", treatment2CrownSide.ToString()),
                        new XElement("treatment_crown_curvature_side", treatment2CrownCurvatureSide.ToString()),
                        new XElement("treatment_crown_radius_or_offset", treatment2CrownRadiusOrOffset),
                        new XElement("treatment_crown_take_off_angle", treatment2CrownTakeOffAngle),
                        new XElement("treatment_draft_angle", treatmentDraftAngle2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetDirection2Treatment: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedProtrusion.GetFromFaceOffsetData(out object fromFaceorPlane, out OffsetSideConstants fromFaceOffsetSide, out double fromFaceOffsetDistance);
                    protrusionElements.Add(new XElement("FromFaceOffsetData",
                        new XElement("from_face_offset_side", fromFaceOffsetSide.ToString()),
                        new XElement("from_face_offset_distance", fromFaceOffsetDistance)));
                    protrusionElements.Add(MM01_geometry_methods.GetPlaneData(fromFaceorPlane, "FromFaceOrPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetFromFaceOffsetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedProtrusion.GetToFaceOffsetData(out object toFaceorPlane, out OffsetSideConstants toFaceOffsetSide, out double toFaceOffsetDistance);
                    protrusionElements.Add(new XElement("ToFaceOffsetData",
                        new XElement("to_face_offset_side", toFaceOffsetSide.ToString()),
                        new XElement("to_face_offset_distance", toFaceOffsetDistance)));
                    protrusionElements.Add(MM01_geometry_methods.GetPlaneData(toFaceorPlane, "ToFaceOrPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetToFaceOffsetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { protrusionElements.Add(MM01_geometry_methods.GetPlaneData(extrudedProtrusion.TopCap, "TopCap")); }
                catch (Exception ex) { Console.WriteLine($"Extrusion GetTopCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(MM01_geometry_methods.GetPlaneData(extrudedProtrusion.BottomCap, "BottomCap")); }
                catch (Exception ex) { Console.WriteLine($"Extrusion GetBottomCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(MM01_geometry_methods.GetPlaneData(extrudedProtrusion.SideFaces, "SideFaces")); }
                catch (Exception ex) { Console.WriteLine($"Extrusion GetSideFaces: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(MM01_geometry_methods.GetPlaneData(extrudedProtrusion.TopCaps, "TopCaps")); }
                catch (Exception ex) { Console.WriteLine($"Extrusion GetTopCaps: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { protrusionElements.Add(MM01_geometry_methods.GetPlaneData(extrudedProtrusion.BottomCaps, "BottomCaps")); }
                catch (Exception ex) { Console.WriteLine($"Extrusion GetBottomCaps: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = extrudedProtrusion.Edges[EdgeTyp];

                    XElement edgeElements = new XElement("edges", new XAttribute("count", edges.Count));

                    for (int e = 1; e <= edges.Count; e++)
                    {
                        var edge = (Edge)edges.Item(e);
                        edgeElements.Add(new XElement($"type{e}", edge.Type.ToString()));

                        edge.GetEndPoints(ref startPoint, ref endPoint);
                        edgeElements.Add(new XElement($"endPoints{e}",
                            new XAttribute("startpoint", string.Join(" ", (double[])startPoint)),
                            new XAttribute("endPoint", string.Join(" ", (double[])endPoint))));

                        Marshal.ReleaseComObject(edge);
                    }

                    protrusionElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = extrudedProtrusion.Faces[faceTyp];
                    protrusionElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynExtrusion = extrudedProtrusion;
                    protrusionElements.Add(new XElement("status", dynExtrusion.Status));
                    protrusionElements.Add(new XElement("suppress", dynExtrusion.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    protrusionElements.Add(new XElement("convertToCutoutAllowed", extrudedProtrusion.ConvertToCutoutAllowed));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion ConvertToCutoutAllowed: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    extrudedProtrusion.GetDimensions(out int numDims, ref dims);
                    protrusionElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedProtrusion.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    protrusionElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Extrusion GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Protrusion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (extrudedProtrusion != null)
                {
                    Marshal.ReleaseComObject(extrudedProtrusion);
                    extrudedProtrusion = null;
                }
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
                try { helixProtrusionElements.Add(new XElement("name", helixProtrusion.Name)); }
                catch (Exception ex) { Console.WriteLine($"HelixProtrusion Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixProtrusionElements.Add(new XElement("type", helixProtrusion.Type)); }
                catch (Exception ex) { Console.WriteLine($"HelixProtrusion Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixProtrusionElements.Add(new XElement("modelingModeType", helixProtrusion.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"HelixProtrusion ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixProtrusionElements.Add(new XElement("showDimensions", helixProtrusion.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"HelixProtrusion ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixProtrusionElements.Add(new XElement("visible", helixProtrusion.Visible)); }
                catch (Exception ex) { Console.WriteLine($"HelixProtrusion Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(helixProtrusion);
                    helixProtrusionElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    helixProtrusion.GetDimensions(out int numDims, ref dims);
                    helixProtrusionElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynHelix = helixProtrusion;
                    helixProtrusionElements.Add(new XElement("status", dynHelix.Status));
                    helixProtrusionElements.Add(new XElement("suppress", dynHelix.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = helixProtrusion.GetStatusEx(out object description);
                    helixProtrusionElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = helixProtrusion.Edges[edgeTyp];

                    XElement edgeElements = new XElement("edges", new XAttribute("count", edges.Count));

                    for (int e = 1; e <= edges.Count; e++)
                    {
                        var edge = (Edge)edges.Item(e);
                        edgeElements.Add(new XElement($"type{e}", edge.Type.ToString()));

                        edge.GetEndPoints(ref startPoint, ref endPoint);
                        edgeElements.Add(new XElement($"endPoints{e}",
                            new XAttribute("startpoint", string.Join(" ", (double[])startPoint)),
                            new XAttribute("endPoint", string.Join(" ", (double[])endPoint))));

                        Marshal.ReleaseComObject(edge);
                    }

                    helixProtrusionElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = helixProtrusion.Faces[faceTyp];
                    helixProtrusionElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    helixProtrusion.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    helixProtrusionElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixProtrusion GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Helix Protrusion: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (helixProtrusion != null)
                {
                    Marshal.ReleaseComObject(helixProtrusion);
                    helixProtrusion = null;
                }
            }

            Console.WriteLine("Created Helix Protrusion XML list");
            return helixProtrusionElements;
        }
    }
}