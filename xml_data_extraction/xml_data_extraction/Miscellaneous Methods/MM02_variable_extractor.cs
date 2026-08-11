using SolidEdgeFramework;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Miscellaneous_Methods
{
    internal class MM02_variable_extractor
    {
        public static XElement Variables_extract(PartDocument partDoc)
        {
            Variables variables = null;
            XElement variablesElement = null;

            try
            {
                try { variables = partDoc.Variables; }
                catch (Exception ex) { Console.WriteLine($"Variables Collection: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (variables != null && variables.Count > 0)
                {
                    variablesElement = new XElement("Variables",
                        new XAttribute("Count", variables.Count));

                    for (int i = 1; i <= variables.Count; i++)
                    {
                        variable varItem = null;
                        try
                        {
                            varItem = (variable)variables.Item(i);

                            XElement variableElement = new XElement("Variable",
                                new XAttribute("index", i));

                            try { variableElement.Add(new XElement("name", varItem.Name)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("displayName", varItem.DisplayName)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] DisplayName: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("systemName", varItem.SystemName)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] SystemName: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("value", varItem.Value)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] Value: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("formula", varItem.Formula)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] Formula: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("unitsType", varItem.UnitsType)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] UnitsType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("type", ((ObjectType)varItem.Type).ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("properties", varItem.Properties)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] Properties: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("variableType", varItem.VariableType)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] VariableType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("variableTableName", varItem.VariableTableName)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] VariableTableName: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("expose", varItem.Expose)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] Expose: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("exposeName", varItem.ExposeName)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] ExposeName: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("isReadOnly", varItem.IsReadOnly)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] IsReadOnly: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { variableElement.Add(new XElement("isSuppressVariable", varItem.IsSuppressVariable)); }
                            catch (Exception ex) { Console.WriteLine($"Variable[{i}] IsSuppressVariable: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            variablesElement.Add(variableElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Variable Item [{i}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (varItem != null)
                            {
                                Marshal.ReleaseComObject(varItem);
                                varItem = null;
                            }
                        }
                    }
                    Console.WriteLine($"Created Variables XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Variables: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (variables != null)
                {
                    Marshal.ReleaseComObject(variables);
                    variables = null;
                }
            }

            return variablesElement;
        }
    }
}