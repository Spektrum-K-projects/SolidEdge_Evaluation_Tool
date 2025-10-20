using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgeFileProperties;
using SolidEdgePart;

namespace xml_data_extraction.Properties
{
    internal class PR01_file_properties_extract
    {
        [STAThread]
        public static XElement Properties(string? subFile)
        {
            PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            Property property = null;

            var prop_dict = new Dictionary<string, object>();

            try
            {
                propertySets = new PropertySets();

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
                        property = (Property)properties[j];

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
    }
}
