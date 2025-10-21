using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgeConstants;
using SolidEdgePart;

namespace xml_data_extraction.Features
{
    internal class FE04_edge_features_extractor
    {
        public static XElement Chamfer(Chamfer chamfer)
        {
            XElement chamferElements = new XElement("Chamfer", new XAttribute("Type", 462094742));

            try
            {
                var chamfer_ref_face = chamfer.ChamferReferenceFace;
                //Console.WriteLine($"CHAMFER. Referenceface: {chamfer_ref_face}");

                var chamfer_type = chamfer.ChamferType;
                chamferElements.Add(new XElement("type", chamfer_type));
                Console.WriteLine($"CHAMFER. Chamfer Type: {chamfer_type}");

                var chamfer_Name = chamfer.Name;
                chamferElements.Add(new XElement("Name", chamfer_type));
                Console.WriteLine($"CHAMFER. Name: {chamfer_Name}");

                var chamfer_set_val_1 = chamfer.ChamferSetbackValue1;
                if (chamfer_set_val_1 != null)
                {
                    chamferElements.Add(new XElement("setback_value_1", chamfer_set_val_1));
                    Console.WriteLine($"CHAMFER. Setback Value 1: {chamfer_set_val_1}");
                }

                var chamfer_set_val_2 = chamfer.ChamferSetbackValue2;
                if (chamfer_set_val_2 != null)
                {
                    chamferElements.Add(new XElement("setback_value_2", chamfer_set_val_2));
                    Console.WriteLine($"CHAMFER. Setback Value 2: {chamfer_set_val_2}");
                }

                var chamfer_set_angle = chamfer.ChamferSetbackAngle;
                if (chamfer_set_angle != null)
                {
                    chamferElements.Add(new XElement("setback_angle", chamfer_set_angle));
                    Console.WriteLine($"CHAMFER. Setback Angle: {chamfer_set_angle}");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Chamfer: Error Message:{ex.Message}");
                return new XElement("Chamfer");
            }

            finally
            {
                if (chamfer != null)
                {
                    Marshal.ReleaseComObject(chamfer);
                    chamfer = null;
                }

            }

            return chamferElements;
        }

        public static XElement Round (Round round)
        {
            XElement roundElements = new XElement("Round", new XAttribute("Type", 462094738));

            try
            {
                var round_type = round.Type;
                roundElements.Add(new XElement("type", round_type));
                Console.WriteLine($"Round Type: {round_type}");

                var round_Name = round.Name;
                roundElements.Add(new XElement("name", round_Name));
                Console.WriteLine($"Round Name: {round_Name}");

                var round_EdgeSetCount = round.EdgeSetCount;
                roundElements.Add(new XElement("edge_set_count", round_EdgeSetCount));
                Console.WriteLine($"Round Edge Set Count: {round_EdgeSetCount}");

                //var round_BlendShape = round.BlendShape;
                //roundElements.Add(new XElement("blend_shape", round_BlendShape));
                //Console.WriteLine($"Round Blend Shape: {round_BlendShape}");
                 
                //var round_BlendShapeValue = round.BlendShapeValue;
                //roundElements.Add(new XElement("blend_shape_value", round_BlendShapeValue));
                //Console.WriteLine($"Round Blend Shape Value: {round_BlendShapeValue}");

                //Extracting the Radius information for the Round
                for (int idx = 1; idx <= round_EdgeSetCount; idx++)
                {
                    try
                    {
                        SolidEdgePart.RoundTypeConstants radiusType = round.RadiusType[idx];
                        Console.WriteLine($"Radius Type: {radiusType}");
                        roundElements.Add(new XElement("radius_type", radiusType));

                        if (radiusType == SolidEdgePart.RoundTypeConstants.igConstantRadius)
                        {
                            System.Array constant_Radii_raw = null;

                            round.GetConstantRadii(out int constant_RadiiCount, ref constant_Radii_raw);

                            double[] constant_Radii = (double[])constant_Radii_raw;

                            if (constant_RadiiCount != null && constant_Radii.Length > 0)
                            {
                                roundElements.Add(new XElement("constantradii_count", constant_RadiiCount));

                                Console.WriteLine("  Constant Radii found:");
                                foreach (double radius in constant_Radii)
                                {
                                    Console.WriteLine($"    Radius: {radius} mm");
                                    roundElements.Add(new XElement("radius", radius));
                                }
                            }
                        }

                        else if (radiusType == SolidEdgePart.RoundTypeConstants.igVariableRadius)
                        {
                            Array variable_Radii = null;

                            round.GetVariableRadii(out int variable_RadiiCount, ref variable_Radii);
                            if (variable_RadiiCount != null && variable_Radii.Length > 0)
                            {
                                roundElements.Add(new XElement("variableradii_count", variable_RadiiCount));

                                Console.WriteLine("  Varaiable Radii found:");
                                foreach (double radius in variable_Radii)
                                {
                                    Console.WriteLine($"    Radius: {radius} mm");
                                    roundElements.Add(new XElement("variable_radii", radius));
                                }
                            }
                        }

                        else
                        {
                            Console.WriteLine("No Radius Found");
                        }

                    }

                    catch (COMException ex) when ((uint)ex.ErrorCode == 0x8002000B)
                    {
                        Console.WriteLine($"EdgeSetIndex {idx}: Index out of bounds. No more radius types for this Round feature.");
                        break;
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Round: Error Message:{ex.Message}");
                return new XElement("Round");
            }

            finally
            {
                if (round != null)
                {
                    Marshal.ReleaseComObject(round);
                    round = null;
                }

            }

            Console.WriteLine($"Created Round Feature XML list");
            return roundElements;
        }
    }
}
