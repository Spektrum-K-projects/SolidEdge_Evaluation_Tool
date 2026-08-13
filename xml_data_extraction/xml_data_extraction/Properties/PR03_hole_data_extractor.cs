using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Properties
{
    internal class PR03_hole_data_extractor
    {
        public static XElement Hole_Data(Hole holeFeature)
        {
            HoleData holeData = null;

            try
            {
                holeData = holeFeature.HoleData;
                return HoleData_Extract(holeData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hole_Data: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("HoleData", "Error");
            }
        }

        public static XElement HoleData_Extract(HoleData holeData)
        {
            XElement holeDataElements = new XElement("HoleData");

            try
            {
                //General Data
                try { holeDataElements.Add(new XElement("name", holeData.Name.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("holeType", holeData.HoleType)); }
                catch (Exception ex) { Console.WriteLine($"HoleData HoleType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("treatmentType", holeData.TreatmentType.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData TreatmentType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("internalThreadDescription", holeData.InternalThreadDescription)); }
                catch (Exception ex) { Console.WriteLine($"HoleData InternalThreadDescription: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("threadLocation", holeData.ThreadLocation.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData ThreadLocation: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                //Hole Size Data
                try { holeDataElements.Add(new XElement("size", holeData.Size.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData Size: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("standard", holeData.Standard.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData Standard: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("subType", holeData.SubType.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData SubType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("units", holeData.Units)); }
                catch (Exception ex) { Console.WriteLine($"HoleData Units: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("fit", holeData.Fit)); }
                catch (Exception ex) { Console.WriteLine($"HoleData Fit: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                //Checking for the Hole Tolerances
                try { holeDataElements.Add(new XElement("isHoleDiameterToleranceApplied", holeData.IsHoleDiameterToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsHoleDiameterToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isHoleDepthToleranceApplied", holeData.IsHoleDepthToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsHoleDepthToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isHeadClearanceToleranceApplied", holeData.IsHeadClearanceToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsHeadClearanceToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isCounterSinkDiameterToleranceApplied", holeData.IsCounterSinkDiameterToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsCounterSinkDiameterToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isCounterBoreDepthToleranceApplied", holeData.IsCounterBoreDepthToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsCounterBoreDepthToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isCounterBoreDiameterToleranceApplied", holeData.IsCounterBoreDiameterToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsCounterBoreDiameterToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isThreadTaperDepthToleranceApplied", holeData.IsThreadTaperDepthToleranceApplied)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsThreadTaperDepthToleranceApplied: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("isThreadLeftHanded", holeData.IsThreadLeftHanded)); }
                catch (Exception ex) { Console.WriteLine($"HoleData IsThreadLeftHanded: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                //Counter Bore Hole Data 
                try
                { 
                
                    XElement counterBoreElements = new XElement("counterBore");

                    counterBoreElements.Add(new XElement("diameter", holeData.CounterboreDiameter),
                                                new XElement("diaToleranceHoleClassName", holeData.CounterBoreDiameterToleranceHoleClassName),
                                                new XElement("diaToleranceHoleClassLower", holeData.CounterBoreDiameterToleranceHoleClassLower),
                                                new XElement("diaToleranceHoleClassUpper", holeData.CounterBoreDiameterToleranceHoleClassUpper),
                                                new XElement("diaToleranceShaftClassName", holeData.CounterBoreDiameterToleranceShaftClassName),
                                                new XElement("diaTolerance", new XElement("type", holeData.CounterBoreDiameterToleranceType),
                                                new XElement("lower", holeData.CounterBoreDiameterUnitToleranceLower),
                                                new XElement("upper", holeData.CounterBoreDiameterUnitToleranceUpper)));

                    counterBoreElements.Add(new XElement("depth", holeData.CounterboreDepth),
                                                new XElement("depthToleranceHoleClassName", holeData.CounterBoreDepthToleranceHoleClassName),
                                                new XElement("depthToleranceHoleClassLower", holeData.CounterBoreDepthToleranceHoleClassLower),
                                                new XElement("depthToleranceHoleClassUpper", holeData.CounterBoreDepthToleranceHoleClassUpper),
                                                new XElement("depthToleranceShaftClassName", holeData.CounterBoreDepthToleranceShaftClassName),
                                                new XElement("depthTolerance", new XElement("type", holeData.CounterBoreDepthToleranceType),
                                                new XElement("lower", holeData.CounterBoreDepthUnitToleranceLower),
                                                new XElement("upper", holeData.CounterBoreDepthUnitToleranceUpper)));

                    counterBoreElements.Add(new XElement("profileLocationType", holeData.CounterboreProfileLocationType));

                    holeDataElements.Add(counterBoreElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData CounterBore: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                //Counter Sink Hole Data
                try
                {
                    holeDataElements.Add(new XElement("counterSink", new XElement("diameter", holeData.CountersinkDiameter),
                                                                        new XElement("angle", holeData.CountersinkAngle),
                                                                        new XElement("toleranceHoleClassName", holeData.CounterSinkDiameterToleranceHoleClassName),
                                                                        new XElement("toleranceHoleClassLower", holeData.CounterSinkDiameterToleranceHoleClassLower),
                                                                        new XElement("toleranceHoleClassUpper", holeData.CounterSinkDiameterToleranceHoleClassUpper),
                                                                        new XElement("toleranceShaftClassName", holeData.CounterSinkDiameterToleranceShaftClassName),
                                                                        new XElement("tolerance", new XElement("type", holeData.CounterSinkDiameterToleranceType),
                                                                        new XElement("lower", holeData.CounterSinkDiameterUnitToleranceLower),
                                                                        new XElement("upper", holeData.CounterSinkDiameterUnitToleranceUpper))));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData CounterSink: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                //Head Clearance Data
                try
                {
                    holeDataElements.Add(new XElement("headClearance", new XElement("clearance", holeData.HeadClearance),
                                                                        new XElement("toleranceHoleClassName", holeData.HeadClearanceToleranceHoleClassName),
                                                                        new XElement("toleranceHoleClassLower", holeData.HeadClearanceToleranceHoleClassLower),
                                                                        new XElement("toleranceHoleClassUpper", holeData.HeadClearanceToleranceHoleClassUpper),
                                                                        new XElement("toleranceShaftClassName", holeData.HeadClearanceToleranceShaftClassName),
                                                                        new XElement("tolerance", new XElement("type", holeData.HeadClearanceToleranceType),
                                                                        new XElement("lower", holeData.HeadClearanceUnitToleranceLower),
                                                                        new XElement("upper", holeData.HeadClearanceUnitToleranceUpper))));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData HeadClearance: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                //Hole Depth Data
                try
                {
                    holeDataElements.Add(new XElement("holeDepth", new XElement("toleranceHoleClassName", holeData.HoleDepthToleranceHoleClassName),
                                                                        new XElement("toleranceHoleClassLower", holeData.HoleDepthToleranceHoleClassLower),
                                                                        new XElement("toleranceHoleClassUpper", holeData.HoleDepthToleranceHoleClassUpper),
                                                                        new XElement("toleranceShaftClassName", holeData.HoleDepthToleranceShaftClassName),
                                                                        new XElement("tolerance", new XElement("type", holeData.HoleDepthToleranceType),
                                                                        new XElement("lower", holeData.HoleDepthUnitToleranceLower),
                                                                        new XElement("upper", holeData.HoleDepthUnitToleranceUpper))));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData HoleDepth: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                //Hole Diameter Data
                try
                {
                    holeDataElements.Add(new XElement("holeDiameter", new XElement("diameter", holeData.HoleDiameter),
                                                                        new XElement("toleranceHoleClassName", holeData.HoleDiameterToleranceHoleClassName),
                                                                        new XElement("toleranceHoleClassLower", holeData.HoleDiameterToleranceHoleClassLower),
                                                                        new XElement("toleranceHoleClassUpper", holeData.HoleDiameterToleranceHoleClassUpper),
                                                                        new XElement("toleranceShaftClassName", holeData.HoleDiameterToleranceShaftClassName),
                                                                        new XElement("tolerance", new XElement("type", holeData.HoleDiameterToleranceType),
                                                                        new XElement("lower", holeData.HoleDiameterUnitToleranceLower),
                                                                        new XElement("upper", holeData.HoleDiameterUnitToleranceUpper))));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData HoleDiameter: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try { holeDataElements.Add(new XElement("bottomAngle", holeData.BottomAngle.ToString())); }
                catch (Exception ex) { Console.WriteLine($"HoleData BottomAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { holeDataElements.Add(new XElement("vBottomDimType", holeData.VBottomDimType)); }
                catch (Exception ex) { Console.WriteLine($"HoleData VBottomDimType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                //Taper Information
                try
                {
                    XElement taperElements = new XElement("taperData");

                    taperElements.Add(new XElement("taper", holeData.Taper));
                    taperElements.Add(new XElement("DimType", holeData.TaperDimType));
                    taperElements.Add(new XElement("LValue", holeData.TaperLValue));
                    taperElements.Add(new XElement("method", holeData.TaperMethod));
                    taperElements.Add(new XElement("RValue", holeData.TaperRValue));

                    holeDataElements.Add(taperElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData Taper: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                //Thread Information
                try
                {
                    XElement threadElements = new XElement("threadData");

                    threadElements.Add(new XElement("depth", holeData.ThreadDepth));
                    threadElements.Add(new XElement("depthMethod", holeData.ThreadDepthMethod));
                    threadElements.Add(new XElement("description", holeData.ThreadDescription));
                    threadElements.Add(new XElement("diamterOption", holeData.ThreadDiameterOption));
                    threadElements.Add(new XElement("externalDiameter", holeData.ThreadExternalDiameter));
                    threadElements.Add(new XElement("minorDiameter", holeData.ThreadMinorDiameter));
                    threadElements.Add(new XElement("nominalDiameter", holeData.ThreadNominalDiameter));
                    threadElements.Add(new XElement("height", holeData.ThreadHeight));
                    threadElements.Add(new XElement("outsideEffectiveThreadLength", holeData.OutsideEffectiveThreadLength));
                    threadElements.Add(new XElement("insideEffectiveThreadLength", holeData.InsideEffectiveThreadLength));
                    threadElements.Add(new XElement("offset", holeData.ThreadOffset));
                    threadElements.Add(new XElement("setting", holeData.ThreadSetting));
                    threadElements.Add(new XElement("tapDrillDiameter", holeData.ThreadTapDrillDiameter));
                    threadElements.Add(new XElement("taperAngle", holeData.ThreadTaperAngle));
                    threadElements.Add(new XElement("taperDepthToleranceHole", new XElement("className", holeData.ThreadTaperDepthToleranceHoleClassName),
                                                                                new XElement("classLower", holeData.ThreadTaperDepthToleranceHoleClassLower),
                                                                                new XElement("classUpper", holeData.ThreadTaperDepthToleranceHoleClassUpper),
                                                                                new XElement("shaftClass", holeData.ThreadTaperDepthToleranceShaftClassName),
                                                                                new XElement("type", holeData.ThreadTaperDepthToleranceType),
                                                                                new XElement("Lower", holeData.ThreadTaperDepthUnitToleranceLower),
                                                                                new XElement("Upper", holeData.ThreadTaperDepthUnitToleranceUpper)));

                    holeDataElements.Add(threadElements);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData ThreadData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    holeData.GetStartChamfer(out int sChamferOn, out double sSetback, out double sAngle);
                    holeDataElements.Add(new XElement("startChamfer", new XElement("chamferOn", sChamferOn),
                                                            new XElement("setback", sSetback),
                                                            new XElement("angle", sAngle)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData GetStartChamfer: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    holeData.GetEndChamfer(out int eChamferOn, out double eSetback, out double eAngle);
                    holeDataElements.Add(new XElement("endChamfer", new XElement("chamferOn", eChamferOn),
                                                            new XElement("setback", eSetback),
                                                            new XElement("angle", eAngle)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData GetEndChamfer: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    holeData.GetNeckChamfer(out int nChamferOn, out double nSetback, out double nAngle);
                    holeDataElements.Add(new XElement("neckChamfer", new XElement("chamferOn", nChamferOn),
                                                            new XElement("setback", nSetback),
                                                            new XElement("angle", nAngle)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData GetNeckChamfer: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    holeData.GetPhysicalThreadClearance(out PhysicalThreadClearanceTypeConstants ePTCType, out double dClearance);
                    holeDataElements.Add(new XElement("physicalThreadClearanceType", new XElement("type", ePTCType),
                                                            new XElement("clearance", dClearance)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HoleData GetPhysicalThreadClearance: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"HoleData: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("HoleData", "Error");
            }
            finally
            {
                if (holeData != null)
                {
                    Marshal.ReleaseComObject(holeData);
                    holeData = null;
                }
            }

            Console.WriteLine($"Created Hole Data XML list");
            return holeDataElements;
        }
    }
}