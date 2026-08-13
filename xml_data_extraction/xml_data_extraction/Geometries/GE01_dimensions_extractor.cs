using System;
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
            XElement dimensionsElement = new XElement("Dimension");

            try
            {
                seDimensions = (Dimensions)profile.Dimensions;

                for (int k = 1; k <= seDimensions.Count; k++)
                {
                    Dimension seDimension = null;
                    try
                    {
                        seDimension = seDimensions.Item(k);
                        ExtractSingleDimension(dimensionsElement, seDimension, k);
                    }
                    finally
                    {
                        if (seDimension != null)
                        {
                            Marshal.ReleaseComObject(seDimension);
                            seDimension = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Dimensions: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seDimensions != null)
                {
                    Marshal.ReleaseComObject(seDimensions);
                    seDimensions = null;
                }
            }

            Console.WriteLine($"\tCreated Dimension XML list");
            return dimensionsElement;
        }

        public static XElement Dimensions_extract_fromArray(Array dimensionsArray)
        {
            XElement dimensionsElement = new XElement("Dimensions");

            if (dimensionsArray == null)
                return dimensionsElement;

            dimensionsElement.Add(new XAttribute("Count", dimensionsArray.Length));

            for (int k = 0; k < dimensionsArray.Length; k++)
            {
                Dimension seDimension = null;
                try
                {
                    seDimension = dimensionsArray.GetValue(k) as Dimension;
                    if (seDimension == null)
                    {
                        Console.WriteLine($"Dimension[{k}]: could not cast to Dimension, skipping.");
                        continue;
                    }

                    XElement dimensionElement = new XElement("Dimension");
                    ExtractSingleDimension(dimensionElement, seDimension, k);
                    dimensionsElement.Add(dimensionElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Dimension[{k}]: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                }
                finally
                {
                    if (seDimension != null)
                    {
                        Marshal.ReleaseComObject(seDimension);
                        seDimension = null;
                    }
                }
            }

            Console.WriteLine("\t\t Created Feature Level Dimensions XML list");
            return dimensionsElement;
        }

        private static readonly DimTypeConstants[] AngleCapableDimTypes =
        {
            DimTypeConstants.igDimTypeAngular,
            DimTypeConstants.igDimTypeArcAngle,
            DimTypeConstants.igDimTypeAngularCoordinate
        };

        private static void ExtractSingleDimension(XElement parent, Dimension seDimension, int k)
        {
            DimTypeConstants dimensionType = default;
            bool dimensionTypeKnown = false;

            try { parent.Add(new XElement("index", seDimension.Index)); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { parent.Add(new XElement("name", seDimension.Name?.ToString())); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { parent.Add(new XElement("type", seDimension.Type)); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { parent.Add(new XElement("isConstrained", seDimension.Constraint)); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] Constraint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try
            {
                dimensionType = seDimension.DimensionType;
                dimensionTypeKnown = true;
                parent.Add(new XElement("dimensionType", dimensionType.ToString()));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Dimension[{k}] DimensionType: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            if (dimensionTypeKnown && Array.IndexOf(AngleCapableDimTypes, dimensionType) >= 0)
            {
                try { parent.Add(new XElement("isAngleClockwise", seDimension.AngleClockwise)); }
                catch (Exception ex) { Console.WriteLine($"Dimension[{k}] AngleClockwise: {ex.Message} | Inner: {ex.InnerException?.Message}"); }
            }

            try { parent.Add(new XElement("value", seDimension.Value)); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] Value: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { parent.Add(new XElement("unitsType", seDimension.UnitsType)); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] UnitsType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { parent.Add(new XElement("isReadOnly", seDimension.IsReadOnly)); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] IsReadOnly: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { parent.Add(new XElement("updateStatus", seDimension.UpdateStatus())); }
            catch (Exception ex) { Console.WriteLine($"Dimension[{k}] UpdateStatus: {ex.Message} | Inner: {ex.InnerException?.Message}"); }
        }
    }
}