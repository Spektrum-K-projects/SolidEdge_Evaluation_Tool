using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using SolidEdgeFileProperties;
using SolidEdgePart;
using SolidEdgeFramework;

namespace xml_data_extraction.Properties
{
    internal class PR01_file_properties_extract
    {
        private static readonly HashSet<string> InterestingPropertyNames = new HashSet<string>
        {
            "Title", "Document Number", "Material", "Density", "Face Style", "Fill Style",
            "Virtual Style", "Thermal Conductivity", "Specific Heat", "Modulus of Elasticity",
            "Poisson's Ratio", "Yield Stress", "Ultimate Stress", "Elongation", "Grouping"
        };

        public static XElement Properties(string? subFile)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;

            var propList = new List<(string Name, object Value)>();

            if (string.IsNullOrEmpty(subFile))
            {
                Console.WriteLine("Properties: no file path provided.");
                return new XElement("properties");
            }

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(subFile, true);

                for (int i = 0; i < propertySets.Count; i++)
                {
                    SolidEdgeFileProperties.Properties properties = null;

                    try
                    {
                        properties = (SolidEdgeFileProperties.Properties)propertySets[i];

                        for (int j = 0; j < properties.Count; j++)
                        {
                            SolidEdgeFileProperties.Property property = null;

                            try
                            {
                                property = (SolidEdgeFileProperties.Property)properties[j];

                                if (InterestingPropertyNames.Contains(property.Name))
                                {
                                    propList.Add((property.Name, property.Value));
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Property [{i}][{j}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                            }
                            finally
                            {
                                if (property != null)
                                {
                                    Marshal.ReleaseComObject(property);
                                    property = null;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"PropertySet [{i}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                    }
                    finally
                    {
                        if (properties != null)
                        {
                            Marshal.ReleaseComObject(properties);
                            properties = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Properties: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (propertySets != null)
                {
                    Marshal.ReleaseComObject(propertySets);
                    propertySets = null;
                }
            }

            var xmlProps = new XElement("properties", propList.Select(p => new XElement("property",
                                            new XAttribute("Name", p.Name), p.Value?.ToString() ?? string.Empty)));
            return xmlProps;
        }

        public static XElement UnitsOfMeasure_Extract(PartDocument partDoc)
        {
            var unitsElement = new XElement("UnitsOfMeasure");
            SolidEdgeFramework.UnitsOfMeasure unitsOfMeasure = null;

            try
            {
                unitsOfMeasure = partDoc.UnitsOfMeasure;
                unitsElement.Add(new XAttribute("Count", unitsOfMeasure.Count));

                for (int u = 1; u <= unitsOfMeasure.Count; u++)
                {
                    var unit = unitsOfMeasure.Item(u);
                    var unitElement = new XElement("Unit");

                    try
                    {
                        unitElement.Add(new XAttribute("Type", unit.Type));
                        unitElement.Add(new XAttribute("Units", unit.Units));
                        unitElement.Add(new XAttribute("Precision", unit.Precision));

                        if (unit.Type == UnitTypeConstants.igUnitDistance)
                        {
                            unitElement.Add(new XAttribute("UnitsName", ((UnitOfMeasureLengthReadoutConstants)unit.Units).ToString()));
                        }
                        else if (unit.Type == UnitTypeConstants.igUnitAngle)
                        {
                            unitElement.Add(new XAttribute("UnitsName", ((UnitOfMeasureAngleReadoutConstants)unit.Units).ToString()));
                        }
                    }
                    catch (Exception ex)
                    {
                        unitElement.Add(new XAttribute("Error", ex.Message));
                    }
                    finally
                    {
                        if (unit != null)
                        {
                            Marshal.ReleaseComObject(unit);
                            unit = null;
                        }
                    }

                    unitsElement.Add(unitElement);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"UnitsOfMeasure: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (unitsOfMeasure != null)
                {
                    Marshal.ReleaseComObject(unitsOfMeasure);
                    unitsOfMeasure = null;
                }
            }

            return unitsElement;
        }

        public static XElement BaseStyle_Extract(PartDocument partDoc)
        {
            var baseStylesElement = new XElement("BaseStyles");

            foreach (PartBaseStylesConstants styleType in Enum.GetValues(typeof(PartBaseStylesConstants)))
            {
                var styleElement = new XElement("BaseStyle", new XAttribute("Category", styleType));
                FaceStyle baseStyle = null;

                try
                {
                    partDoc.GetBaseStyle(styleType, out baseStyle);

                    if (baseStyle == null)
                    {
                        styleElement.Add(new XAttribute("NotSet", true));
                    }
                    else
                    {
                        try { styleElement.Add(new XAttribute("StyleName", baseStyle.StyleName ?? "")); }
                        catch (Exception ex) { styleElement.Add(new XAttribute("StyleNameError", ex.Message)); }

                        try { styleElement.Add(new XAttribute("Material", baseStyle.Material ?? "")); }
                        catch (Exception ex) { styleElement.Add(new XAttribute("MaterialError", ex.Message)); }

                        try
                        {
                            styleElement.Add(new XElement("DiffuseColor",
                                new XAttribute("R", baseStyle.DiffuseRed),
                                new XAttribute("G", baseStyle.DiffuseGreen),
                                new XAttribute("B", baseStyle.DiffuseBlue)));
                        }
                        catch (Exception ex) { styleElement.Add(new XAttribute("DiffuseColorError", ex.Message)); }

                        try
                        {
                            styleElement.Add(new XElement("SpecularColor",
                                new XAttribute("R", baseStyle.SpecularRed),
                                new XAttribute("G", baseStyle.SpecularGreen),
                                new XAttribute("B", baseStyle.SpecularBlue)));
                        }
                        catch (Exception ex) { styleElement.Add(new XAttribute("SpecularColorError", ex.Message)); }

                        try
                        {
                            styleElement.Add(new XElement("AmbientColor",
                                new XAttribute("R", baseStyle.AmbientRed),
                                new XAttribute("G", baseStyle.AmbientGreen),
                                new XAttribute("B", baseStyle.AmbientBlue)));
                        }
                        catch (Exception ex) { styleElement.Add(new XAttribute("AmbientColorError", ex.Message)); }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"BaseStyle {styleType}: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                    styleElement.Add(new XAttribute("Error", ex.Message));
                }
                finally
                {
                    if (baseStyle != null)
                    {
                        Marshal.ReleaseComObject(baseStyle);
                        baseStyle = null;
                    }
                }

                baseStylesElement.Add(styleElement);
            }

            return baseStylesElement;
        }
    }
}