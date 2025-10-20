using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgePart;

namespace xml_data_extraction.Features
{
    internal class FE04_chamfer_extractor
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
    }
}
