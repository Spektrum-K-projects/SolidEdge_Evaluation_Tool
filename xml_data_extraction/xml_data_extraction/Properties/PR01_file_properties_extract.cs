using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgeFileProperties;
using SolidEdgePart;
using SolidEdgeFramework;

namespace xml_data_extraction.Properties
{
    internal class PR01_file_properties_extract
    {
        [STAThread]
        public static XElement Properties(string? subFile)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;

            var prop_dict = new Dictionary<string, object>();

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();

                //Command to specify file path
                string file_path = subFile;
                Console.WriteLine($"  Extracting properties from: {file_path}");

                //Command to access the Solid Edge file in the path
                propertySets.Open(@file_path, true);

                for (int i = 0; i < propertySets.Count; i++)
                {
                    properties = (SolidEdgeFileProperties.Properties)propertySets[i];
                    // Console.WriteLine($"Properties [{i}]: {properties.Name}");


                    for (int j = 0; j < properties.Count; j++)
                    {
                        property = (SolidEdgeFileProperties.Property)properties[j];

                        if (property.Name == "Title")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Document Number")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Material")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Density")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Face Style")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Fill Style")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Virtual Style")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Thermal Conductivity")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Specific Heat")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Modulus of Elasticity")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Poisson's Ratio")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Yield Stress")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Ultimate Stress")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Elongation")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                        if (property.Name == "Grouping")
                        {
                            prop_dict.Add(property.Name, property.Value);
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                if (property != null)
                {
                    Marshal.ReleaseComObject(property);
                    property = null;
                }
                if (properties != null)
                {
                    Marshal.ReleaseComObject(properties);
                    properties = null;
                }
                if (propertySets != null)
                {
                    Marshal.ReleaseComObject(propertySets);
                    propertySets = null;
                }
            }

            var xmlProps = new XElement("properties", prop_dict.Select(kv => new XElement("property", 
                                            new XAttribute("Name", kv.Key), kv.Value?.ToString() ?? string.Empty)));
            return xmlProps;
        }

        public static XElement UnitsOfMeasure_Extract(PartDocument partDoc)
        {
            var unitsElement = new XElement("UnitsOfMeasure");

            try
            {
                var unitsOfMeasure = partDoc.UnitsOfMeasure;
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

                        // NEW: friendly decode for the two categories that matter most
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

            return unitsElement;
        }

        public static XElement BaseStyle_Extract(PartDocument partDoc)
        {
            var baseStylesElement = new XElement("BaseStyles");

            foreach (PartBaseStylesConstants styleType in Enum.GetValues(typeof(PartBaseStylesConstants)))
            {
                var styleElement = new XElement("BaseStyle", new XAttribute("Category", styleType));

                try
                {
                    partDoc.GetBaseStyle(styleType, out FaceStyle baseStyle);

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

                baseStylesElement.Add(styleElement);
            }

            return baseStylesElement;
        }
    }
}
