using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgePart;
using SolidEdgeFrameworkSupport;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Geometries
{
    internal class GE01_dimensions_extractor
    {
        public static XElement Dimension_extract(Profile profile)
        {
            Dimensions seDimensions = null;
            Dimension seDimension = null;

            XElement dimensionsElement = new XElement("Dimension");

            try
            {
                seDimensions = (Dimensions)profile.Dimensions;

                for (int k = 1; k <= seDimensions.Count; k++)
                {
                    seDimension = seDimensions.Item(k);

                    dimensionsElement.Add(new XElement("index", seDimension.Index));

                    dimensionsElement.Add(new XElement("name", seDimension.Name?.ToString()));
                    //Console.WriteLine($"Dimension measured [{k}]: {seDimension.Name}");

                    dimensionsElement.Add(new XElement("type", seDimension.Type));

                    dimensionsElement.Add(new XElement("isConstrainted", seDimension.Constraint));

                    dimensionsElement.Add(new XElement("isAngleClockwise", seDimension.AngleClockwise));

                    dimensionsElement.Add(new XElement("dimensionType", seDimension.DimensionType));
                    //Console.WriteLine($"Dimension type [{k}]: {seDimension.DimensionType}");

                    dimensionsElement.Add(new XElement("value", seDimension.Value));
                    //Console.WriteLine($"Dimension value [{k}]: {seDimension.Value}");

                    dimensionsElement.Add(new XElement("unitsType", seDimension.UnitsType));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Dimensions: Error Message:{ex.Message}");           
            }
            finally
            {
                if (seDimension != null)
                {
                    Marshal.ReleaseComObject(seDimension);
                    seDimension = null;
                }
                if (seDimensions != null)
                {
                    Marshal.ReleaseComObject(seDimensions);
                    seDimensions = null;
                }
            }

            //XElement xmlDimension = new XElement("dimensions", dimensionsElement.Select(kv => new XElement("dimension",
            //                                new XAttribute("Name", kv.Key), kv.Value?.ToString() ?? String.Empty)));
            Console.WriteLine($"Created Dimension XML list");

            //return xmlDimension;

            return dimensionsElement;
        }
    }
}
