using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace xml_data_extraction.Features
{
    internal class FE10_mirror_extractor
    {
        public static XElement MirrorPart(MirrorPart mirrorPart)
        {
            XElement mirrorPartElements = new XElement("Mirror_Part", new XAttribute("Type", 1908287958));

            try
            {
                var name = mirrorPart.Name;
                mirrorPartElements.Add(new XElement("name", name));

                var type = mirrorPart.Type;
                mirrorPartElements.Add(new XElement("type", type));

                var mirrorPlane = mirrorPart.MirrorPlane;
                mirrorPartElements.Add(new XElement("mirrorPlane", mirrorPlane));
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Mirror Part: Error Message:{ex.Message}");
                return new XElement("Miror_Part", "Error");
            }

            finally
            {
                if (mirrorPart != null)
                {
                    Marshal.ReleaseComObject(mirrorPart);
                    mirrorPart = null;
                }
            }

            Console.WriteLine($"Created Mirror Part XML list");
            return mirrorPartElements;
        }

        public static XElement MirrorCopy(MirrorCopy mirrorCopy)
        {
            XElement mirrorCopyElements = new XElement("MirrorCopy", new XAttribute("Type", 66247736));

            try
            {
                var name = mirrorCopy.Name;
                mirrorCopyElements.Add(new XElement("name", name));

                var type = mirrorCopy.Type;
                mirrorCopyElements.Add(new XElement("type", type));

                var mirrorVisible = mirrorCopy.Visible;
                mirrorCopyElements.Add(new XElement("isVisible", mirrorVisible));

                var mirrorModelingModeType = mirrorCopy.ModelingModeType;
                mirrorCopyElements.Add(new XElement("modelingModeType", mirrorModelingModeType));

                //var mirrorPatternPlane = mirrorCopy.PatternPlane;
                //int objType = mirrorPatternPlane.ObjectType;
                //mirrorCopyElements.Add(new XElement("patternPlaneType", objType));

                var mirrorNumberOfInputs = mirrorCopy.NumberInputFeatures;
                mirrorCopyElements.Add(new XElement("NumberOfInputFeatures", mirrorNumberOfInputs));

                mirrorCopy.GetNumberOfOccurrences(out int numOcc, out int numFea);
                mirrorCopyElements.Add(new XElement("numberOfOccurrences", numOcc));
                mirrorCopyElements.Add(new XElement("numberOfFeaturesPerOccurrences", numFea));
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Mirror Copy: Error Message:{ex.Message}");
                return new XElement("Miror_Copy", "Error");
            }

            finally
            {
                if (mirrorCopy != null)
                {
                    Marshal.ReleaseComObject(mirrorCopy);
                    mirrorCopy = null;
                }
            }

            Console.WriteLine($"Created Mirror Copy XML list");
            return mirrorCopyElements;
        }
    }
}
