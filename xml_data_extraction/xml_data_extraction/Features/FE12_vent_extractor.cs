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
    internal class FE12_vent_extractor
    {
        public static XElement Vent(Vent vent)
        {
            XElement ventElements = new XElement("Vent", new XAttribute("Type", -85880079));

            //Array dimensionsArray = Array.CreateInstance(typeof(double), 0);

            try
            {
                var name = vent.Name;
                ventElements.Add(new XElement("name", name));

                var type = vent.Type;
                ventElements.Add(new XElement("type", type));

                var extentDepth = vent.ExtentDepth;
                ventElements.Add(new XElement("extentDepth", extentDepth));

                var extentSide = vent.ExtentSide;
                ventElements.Add(new XElement("extentSide", extentSide));

                var extentType = vent.ExtentType;
                ventElements.Add(new XElement("extentType", extentType));

                var ventModelingModeType = vent.ModelingModeType;
                ventElements.Add(new XElement("modelingModeType", ventModelingModeType));

                var ribDepth = vent.RibDepth;
                ventElements.Add(new XElement("ribDepth", ribDepth));

                var ribThickness = vent.RibThickness;
                ventElements.Add(new XElement("ribThickness", ribThickness));

                var ribExtension = vent.RibExtension;
                ventElements.Add(new XElement("ribExtension", ribExtension));

                var ribOffset = vent.RibOffset;
                ventElements.Add(new XElement("ribOffset", ribOffset));

                var roundEnabled = vent.RoundEnabled;
                ventElements.Add(new XElement("roundEnabled", roundEnabled));

                var roundRadius = vent.RoundRadius;
                ventElements.Add(new XElement("roundRadius", roundRadius));

                var sparDepth = vent.SparDepth;
                ventElements.Add(new XElement("sparDepth", sparDepth));

                var sparExtension = vent.SparExtension;
                ventElements.Add(new XElement("sparExtension", sparExtension));

                var sparOffset = vent.SparOffset;
                ventElements.Add(new XElement("sparOffset", sparOffset));

                var sparThickness = vent.SparThickness;
                ventElements.Add(new XElement("sparThickness", sparThickness));

                var draftAngle = vent.DraftAngle;
                ventElements.Add(new XElement("draftAngle", draftAngle));

                var draftEnabled = vent.DraftEnabled;
                ventElements.Add(new XElement("draftEnabled", draftEnabled));

                var draftFromOutsideEdges = vent.DraftFromOutsideEdges;
                ventElements.Add(new XElement("draftFromOutsideEdges", draftFromOutsideEdges));

                var draftSide = vent.DraftSide;
                ventElements.Add(new XElement("draftSide", draftSide));

                //vent.GetDimensions(out int numDim, ref dimensionsArray);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Vent: Error Message:{ex.Message}");
                return new XElement("Vent", "Error");
            }
            finally
            {
                if (vent != null)
                {
                    Marshal.ReleaseComObject(vent);
                    vent = null;
                }
            }

            Console.WriteLine($"Created Vent XML list");
            return ventElements;

        }
    }
}
