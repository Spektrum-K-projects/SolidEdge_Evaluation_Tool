using System.Runtime.InteropServices;
using SolidEdgePart;
using System.Xml.Linq;
using xml_data_extraction.Geometries;
using SolidEdgeGeometry;
using xml_data_extraction.Miscellaneous_Methods;

namespace xml_data_extraction.Features
{
    internal class FE03_cutout_extractor
    {
        public static XElement Cutout_Extrude(ExtrudedCutout extrudedCutout)
        {
            XElement cutoutElements = new XElement("Cutout", new XAttribute("Type", 462094714));

            try
            {
                try { cutoutElements.Add(new XElement("name", extrudedCutout.Name)); }
                catch (Exception ex) { Console.WriteLine($"Cutout Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("type", extrudedCutout.Type)); }
                catch (Exception ex) { Console.WriteLine($"Cutout Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("modelingModeType", extrudedCutout.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Cutout ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("showDimensions", extrudedCutout.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Cutout ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("visible", extrudedCutout.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Cutout Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("depth", extrudedCutout.Depth)); }
                catch (Exception ex) { Console.WriteLine($"Cutout Depth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("extent_side", extrudedCutout.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Cutout ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("extrude_type", extrudedCutout.ExtentType)); }
                catch (Exception ex) { Console.WriteLine($"Cutout ExtentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(new XElement("profile_side", extrudedCutout.ProfileSide)); }
                catch (Exception ex) { Console.WriteLine($"Cutout ProfileSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(extrudedCutout);
                    cutoutElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedCutout.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side, out double finiteDepth1);
                    cutoutElements.Add(new XElement("Direction1Extent",
                        new XElement("extent_type", extent1Type.ToString()),
                        new XElement("extent_side", extent1Side.ToString()),
                        new XElement("finite_depth", finiteDepth1)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetDirection1Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
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
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetDirection1Treatment: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedCutout.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side, out double finiteDepth2);
                    cutoutElements.Add(new XElement("Direction2Extent",
                        new XElement("extent_type", extent2Type.ToString()),
                        new XElement("extent_side", extent2Side.ToString()),
                        new XElement("finite_depth", finiteDepth2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetDirection2Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedCutout.GetDirection2Treatment(out TreatmentTypeConstants treatment2Type, out DraftSideConstants draft2Side, out double treatmentDraftAngle2,
                        out TreatmentCrownTypeConstants treatment2CrownType, out TreatmentCrownSideConstants treatment2CrownSide,
                        out TreatmentCrownCurvatureSideConstants treatment2CrownCurvatureSide, out double treatment2CrownRadiusOrOffset,
                        out double treatment2CrownTakeOffAngle);
                    cutoutElements.Add(new XElement("Direction2Treatment",
                        new XElement("treatment_type", treatment2Type.ToString()),
                        new XElement("draft_side", draft2Side.ToString()),
                        new XElement("treatment_draft_angle", treatmentDraftAngle2),
                        new XElement("treatment_crown_type", treatment2CrownType.ToString()),
                        new XElement("treatment_crown_side", treatment2CrownSide.ToString()),
                        new XElement("treatment_crown_curvature_side", treatment2CrownCurvatureSide.ToString()),
                        new XElement("treatment_crown_radius_or_offset", treatment2CrownRadiusOrOffset),
                        new XElement("treatment_crown_take_off_angle", treatment2CrownTakeOffAngle)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetDirection2Treatment: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedCutout.GetFromFaceOffsetData(out object fromFaceorPlane, out OffsetSideConstants fromFaceOffsetSide, out double fromFaceOffsetDistance);
                    cutoutElements.Add(new XElement("FromFaceOffsetData",
                        new XElement("from_face_offset_side", fromFaceOffsetSide.ToString()),
                        new XElement("from_face_offset_distance", fromFaceOffsetDistance)));
                    cutoutElements.Add(MM01_geometry_methods.GetPlaneData(fromFaceorPlane, "FromFaceOrPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetFromFaceOffsetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedCutout.GetToFaceOffsetData(out object toFaceorPlane, out OffsetSideConstants toFaceOffsetSide, out double toFaceOffsetDistance);
                    cutoutElements.Add(new XElement("ToFaceOffsetData",
                        new XElement("to_face_offset_side", toFaceOffsetSide.ToString()),
                        new XElement("to_face_offset_distance", toFaceOffsetDistance)));
                    cutoutElements.Add(MM01_geometry_methods.GetPlaneData(toFaceorPlane, "ToFaceOrPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetToFaceOffsetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Cap / side faces ----
                try { cutoutElements.Add(MM01_geometry_methods.GetPlaneData(extrudedCutout.TopCap, "TopCap")); }
                catch (Exception ex) { Console.WriteLine($"Cutout GetTopCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(MM01_geometry_methods.GetPlaneData(extrudedCutout.BottomCap, "BottomCap")); }
                catch (Exception ex) { Console.WriteLine($"Cutout GetBottomCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(MM01_geometry_methods.GetPlaneData(extrudedCutout.SideFaces, "SideFaces")); }
                catch (Exception ex) { Console.WriteLine($"Cutout GetSideFaces: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(MM01_geometry_methods.GetPlaneData(extrudedCutout.TopCaps, "TopCaps")); }
                catch (Exception ex) { Console.WriteLine($"Cutout GetTopCaps: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { cutoutElements.Add(MM01_geometry_methods.GetPlaneData(extrudedCutout.BottomCaps, "BottomCaps")); }
                catch (Exception ex) { Console.WriteLine($"Cutout GetBottomCaps: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = extrudedCutout.Edges[EdgeTyp];

                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

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

                    cutoutElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = extrudedCutout.Faces[faceTyp];
                    cutoutElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynCutout = extrudedCutout;
                    cutoutElements.Add(new XElement("status", dynCutout.Status));
                    cutoutElements.Add(new XElement("suppress", dynCutout.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = extrudedCutout.GetStatusEx(out object description);
                    cutoutElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    extrudedCutout.GetDimensions(out int numDims, ref dims);
                    cutoutElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    extrudedCutout.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    cutoutElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array bodyArray = Array.CreateInstance(typeof(object), 0);
                    extrudedCutout.GetBodyArray(out bool multiBodyCut, out int numberOfBodies, out bodyArray);
                    cutoutElements.Add(new XElement("BodyArray",
                        new XAttribute("MultiBodyCut", multiBodyCut),
                        new XAttribute("NumberOfBodies", numberOfBodies)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Cutout GetBodyArray: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Extruded Cutout: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
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
                try { helixCutoutElements.Add(new XElement("name", helixCutout.Name)); }
                catch (Exception ex) { Console.WriteLine($"HelixCutout Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixCutoutElements.Add(new XElement("type", helixCutout.Type)); }
                catch (Exception ex) { Console.WriteLine($"HelixCutout Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixCutoutElements.Add(new XElement("modelingModeType", helixCutout.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"HelixCutout ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixCutoutElements.Add(new XElement("showDimensions", helixCutout.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"HelixCutout ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { helixCutoutElements.Add(new XElement("visible", helixCutout.Visible)); }
                catch (Exception ex) { Console.WriteLine($"HelixCutout Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(helixCutout);
                    helixCutoutElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    helixCutout.GetDimensions(out int numDims, ref dims);
                    helixCutoutElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynHelixCutout = helixCutout;
                    helixCutoutElements.Add(new XElement("status", dynHelixCutout.Status));
                    helixCutoutElements.Add(new XElement("suppress", dynHelixCutout.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = helixCutout.GetStatusEx(out object description);
                    helixCutoutElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

                    FeatureTopologyQueryTypeConstants edgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = helixCutout.Edges[edgeTyp];

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

                    helixCutoutElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = helixCutout.Faces[faceTyp];
                    helixCutoutElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    helixCutout.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    helixCutoutElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HelixCutout GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Helix Cutout: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (helixCutout != null)
                {
                    Marshal.ReleaseComObject(helixCutout);
                    helixCutout = null;
                }
            }

            Console.WriteLine("Created Helix Cutout XML list");
            return helixCutoutElements;
        }

        // Slot Feature Data Extraction
        public static XElement Slot_Extract(Slot slot)
        {
            XElement slotElements = new XElement("Slot", new XAttribute("Type", 1336556513));

            try
            {
                try { slotElements.Add(new XElement("name", slot.Name)); }
                catch (Exception ex) { Console.WriteLine($"Slot Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("type", slot.Type)); }
                catch (Exception ex) { Console.WriteLine($"Slot Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("modelingModeType", slot.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Slot ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("showDimensions", slot.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Slot ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("slot_type", slot.SlotType.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Slot SlotType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("slot_end_condition", slot.SlotEndCondition.ToString())); }
                catch (Exception ex) { Console.WriteLine($"Slot SlotEndCondition: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("slot_width", slot.SlotWidth)); }
                catch (Exception ex) { Console.WriteLine($"Slot SlotWidth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("slot_offset_width", slot.SlotOffsetWidth)); }
                catch (Exception ex) { Console.WriteLine($"Slot SlotOffsetWidth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotElements.Add(new XElement("slot_offset_depth", slot.SlotOffsetDepth)); }
                catch (Exception ex) { Console.WriteLine($"Slot SlotOffsetDepth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    var profile_extract = GE04_getProfiles_extractor.getProfile_extract(slot);
                    slotElements.Add(profile_extract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    slot.GetDirection1Extent(out FeaturePropertyConstants extent1Type, out FeaturePropertyConstants extent1Side, out double finiteDepth1);
                    slotElements.Add(new XElement("Direction1Extent",
                        new XElement("extent_type", extent1Type.ToString()),
                        new XElement("extent_side", extent1Side.ToString()),
                        new XElement("finite_depth", finiteDepth1)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetDirection1Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    slot.GetDirection2Extent(out FeaturePropertyConstants extent2Type, out FeaturePropertyConstants extent2Side, out double finiteDepth2);
                    slotElements.Add(new XElement("Direction2Extent",
                        new XElement("extent_type", extent2Type.ToString()),
                        new XElement("extent_side", extent2Side.ToString()),
                        new XElement("finite_depth", finiteDepth2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetDirection2Extent: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    slot.GetFromFaceOffsetData(out object fromFaceorPlane, out OffsetSideConstants fromFaceOffsetSide, out double fromFaceOffsetDistance);
                    slotElements.Add(new XElement("FromFaceOffsetData",
                        new XElement("from_face_offset_side", fromFaceOffsetSide.ToString()),
                        new XElement("from_face_offset_distance", fromFaceOffsetDistance)));
                    slotElements.Add(MM01_geometry_methods.GetPlaneData(fromFaceorPlane, "FromFaceOrPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetFromFaceOffsetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    slot.GetToFaceOffsetData(out object toFaceorPlane, out OffsetSideConstants toFaceOffsetSide, out double toFaceOffsetDistance);
                    slotElements.Add(new XElement("ToFaceOffsetData",
                        new XElement("to_face_offset_side", toFaceOffsetSide.ToString()),
                        new XElement("to_face_offset_distance", toFaceOffsetDistance)));
                    slotElements.Add(MM01_geometry_methods.GetPlaneData(toFaceorPlane, "ToFaceOrPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetToFaceOffsetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = slot.Edges[EdgeTyp];

                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

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

                    slotElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = slot.Faces[faceTyp];
                    slotElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynSlot = slot;
                    slotElements.Add(new XElement("status", dynSlot.Status));
                    slotElements.Add(new XElement("suppress", dynSlot.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = slot.GetStatusEx(out object description);
                    slotElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    slot.GetDimensions(out int numDims, ref dims);
                    slotElements.Add(GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    slot.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    slotElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    slot.ExactRange(out double ex1, out double ey1, out double ez1, out double ex2, out double ey2, out double ez2);
                    slotElements.Add(new XElement("ExactRange",
                        new XAttribute("X1", ex1), new XAttribute("Y1", ey1), new XAttribute("Z1", ez1),
                        new XAttribute("X2", ex2), new XAttribute("Y2", ey2), new XAttribute("Z2", ez2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetExactRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array bodyArray = Array.CreateInstance(typeof(object), 0);
                    slot.GetBodyArray(out bool multiBodyCut, out int numberOfBodies, out bodyArray);
                    slotElements.Add(new XElement("BodyArray",
                        new XAttribute("MultiBodyCut", multiBodyCut),
                        new XAttribute("NumberOfBodies", numberOfBodies)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Slot GetBodyArray: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Slot: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (slot != null)
                {
                    Marshal.ReleaseComObject(slot);
                    slot = null;
                }
            }

            Console.WriteLine($"Created Slot XML list");
            return slotElements;
        }

        // Slot Group Feature Data Extraction
        public static XElement SlotGroup_Extract(SlotGroup slotGroup)
        {
            XElement slotGroupElements = new XElement("SlotGroup", new XAttribute("Type", 339115113));

            try
            {
                try { slotGroupElements.Add(new XElement("type", slotGroup.Type)); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotGroupElements.Add(new XElement("slot_type", slotGroup.SlotType.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup SlotType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotGroupElements.Add(new XElement("slot_end_condition", slotGroup.SlotEndCondition.ToString())); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup SlotEndCondition: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotGroupElements.Add(new XElement("slot_width", slotGroup.SlotWidth)); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup SlotWidth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotGroupElements.Add(new XElement("slot_offset_width", slotGroup.SlotOffsetWidth)); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup SlotOffsetWidth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotGroupElements.Add(new XElement("slot_offset_depth", slotGroup.SlotOffsetDepth)); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup SlotOffsetDepth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { slotGroupElements.Add(new XElement("occurrence_count", slotGroup.Count)); }
                catch (Exception ex) { Console.WriteLine($"SlotGroup Count: {ex.Message} | Inner: {ex.InnerException?.Message}"); }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SlotGroup: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (slotGroup != null)
                {
                    Marshal.ReleaseComObject(slotGroup);
                    slotGroup = null;
                }
            }

            Console.WriteLine($"Created SlotGroup XML list");
            return slotGroupElements;
        }

        // Normal Cutout Data Extraction
        public static XElement Normal_Cutout(NormalCutout normalCutout)
        {
            XElement normalCutoutElements = new XElement("NormalCutout",
                                            new XAttribute("Type", -292547215));

            try
            {
                try { normalCutoutElements.Add(new XElement("name", normalCutout.Name)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("type", normalCutout.Type)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("modeling_type", normalCutout.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("showDimensions", normalCutout.ShowDimensions)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout ShowDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("visible", normalCutout.Visible)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("method", normalCutout.Method)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout Method: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("depth", normalCutout.Depth)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout Depth: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("extent_side", normalCutout.ExtentSide)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout ExtentSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("extent_type", normalCutout.ExtentType)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout ExtentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(new XElement("profile_side", normalCutout.ProfileSide)); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout ProfileSide: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    FeatureTopologyQueryTypeConstants EdgeTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var edges = normalCutout.Edges[EdgeTyp];

                    Array startPoint = Array.CreateInstance(typeof(double), 0);
                    Array endPoint = Array.CreateInstance(typeof(double), 0);

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

                    normalCutoutElements.Add(edgeElements);
                    Marshal.ReleaseComObject(edges);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout GetEdges: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureTopologyQueryTypeConstants faceTyp = FeatureTopologyQueryTypeConstants.igQueryAll;
                    var faces = normalCutout.Faces[faceTyp];
                    normalCutoutElements.Add(new XElement("Faces", new XAttribute("Count", faces.Count)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout GetFaces: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                // ---- Cap / side faces ----
                try { normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.TopCap, "TopCap")); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout GetTopCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.BottomCap, "BottomCap")); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout GetBottomCap: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.SideFaces, "SideFaces")); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout GetSideFaces: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.TopCaps, "TopCaps")); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout GetTopCaps: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.BottomCaps, "BottomCaps")); }
                catch (Exception ex) { Console.WriteLine($"Normal Cutout GetBottomCaps: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);

                    normalCutout.GetDimensions(out int numDims, ref dims);

                    normalCutoutElements.Add(
                        GE01_dimensions_extractor.Dimensions_extract_fromArray(dims));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    XElement profileExtract =
                        GE04_getProfiles_extractor.getProfile_extract(normalCutout);

                    normalCutoutElements.Add(profileExtract);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout Profiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    dynamic dynNormalCutout = normalCutout;
                    normalCutoutElements.Add(new XElement("status", dynNormalCutout.Status));
                    normalCutoutElements.Add(new XElement("suppress", dynNormalCutout.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = normalCutout.GetStatusEx(out object description);
                    normalCutoutElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    normalCutout.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    normalCutoutElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.FromPlane, "FromPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout FromPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    normalCutoutElements.Add(MM01_geometry_methods.GetPlaneData(normalCutout.ToPlane, "ToPlane"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Normal Cutout ToPlane: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Normal Cutout: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (normalCutout != null)
                {
                    Marshal.ReleaseComObject(normalCutout);
                    normalCutout = null;
                }
            }

            Console.WriteLine($"Created Normal Cutout XML");
            return normalCutoutElements;
        }
    }
}