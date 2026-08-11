using SolidEdgeFrameworkSupport;
using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Miscellaneous_Methods
{
    internal class MM01_geometry_methods
    {
        public static XElement GetPlaneData(object planeObj, string elementName = "Plane")
        {
            if (planeObj == null)
                return new XElement(elementName, "null");

            try
            {
                if (planeObj is RefPlane refPlane)
                {
                    var planeElement = new XElement(elementName);

                    try { planeElement.Add(new XAttribute("Name", refPlane.Name ?? "Unnamed")); }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try { planeElement.Add(new XAttribute("Global", refPlane.Global)); }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane Global: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try { planeElement.Add(new XAttribute("Visible", refPlane.Visible)); }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane Visible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try { planeElement.Add(new XAttribute("Type", refPlane.Type.ToString())); }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try { planeElement.Add(new XAttribute("TangentAngle", refPlane.TangentAngle)); }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane TangentAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try { planeElement.Add(new XAttribute("ModelingModeType", refPlane.ModelingModeType.ToString())); }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try
                    {
                        Array normal = Array.CreateInstance(typeof(double), 0);
                        refPlane.GetNormal(ref normal);
                        if (normal.Length >= 3)
                            planeElement.Add(new XElement("Normal",
                                new XAttribute("X", normal.GetValue(0)), new XAttribute("Y", normal.GetValue(1)), new XAttribute("Z", normal.GetValue(2))));
                    }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane GetNormal: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try
                    {
                        Array rootPoint = Array.CreateInstance(typeof(double), 0);
                        refPlane.GetRootPoint(ref rootPoint);
                        if (rootPoint.Length >= 3)
                            planeElement.Add(new XElement("RootPoint",
                                new XAttribute("X", rootPoint.GetValue(0)), new XAttribute("Y", rootPoint.GetValue(1)), new XAttribute("Z", rootPoint.GetValue(2))));
                    }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane GetRootPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    try
                    {
                        Array refDir = Array.CreateInstance(typeof(double), 0);
                        refPlane.GetReferenceDirection(ref refDir);
                        if (refDir.Length >= 3)
                            planeElement.Add(new XElement("ReferenceDirection",
                                new XAttribute("X", refDir.GetValue(0)), new XAttribute("Y", refDir.GetValue(1)), new XAttribute("Z", refDir.GetValue(2))));
                    }
                    catch (Exception ex) { Console.WriteLine($"{elementName} RefPlane GetReferenceDirection: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                    Marshal.ReleaseComObject(refPlane);
                    return planeElement;
                }
                else if (planeObj is Faces facesCollection)
                {
                    var facesElement = new XElement(elementName, new XAttribute("Kind", "Faces"), new XAttribute("Count", facesCollection.Count));

                    for (int i = 1; i <= facesCollection.Count; i++)
                    {
                        Face faceItem = null;
                        try
                        {
                            faceItem = (Face)facesCollection.Item(i);
                            facesElement.Add(Face_Data(faceItem));
                        }
                        catch (Exception ex)
                        {
                            facesElement.Add(new XElement("Face", new XAttribute("Error", ex.Message)));
                            Console.WriteLine($"{elementName} Faces[{i}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (faceItem != null)
                            {
                                Marshal.ReleaseComObject(faceItem);
                                faceItem = null;
                            }
                        }
                    }

                    Marshal.ReleaseComObject(facesCollection);
                    return facesElement;
                }
                else if (planeObj is Face face)
                {
                    var singleFaceElement = new XElement(elementName, new XAttribute("Kind", "Face"));
                    singleFaceElement.Add(Face_Data(face).Elements());
                    singleFaceElement.Add(Face_Data(face).Attributes());
                    Marshal.ReleaseComObject(face);
                    return singleFaceElement;
                }
                else
                {
                    var unknownElement = new XElement(elementName, new XAttribute("Kind", "Unknown"));

                    try
                    {
                        dynamic dynObj = planeObj;
                        unknownElement.Add(new XAttribute("Type", dynObj.Type));
                    }
                    catch (Exception ex)
                    {
                        unknownElement.Add(new XAttribute("Note", $"Type not readable: {ex.Message}"));
                    }

                    if (Marshal.IsComObject(planeObj))
                    {
                        Marshal.ReleaseComObject(planeObj);
                    }

                    return unknownElement;
                }
            }
            catch (Exception ex)
            {
                return new XElement(elementName, new XAttribute("Error", ex.Message));
            }
        }

        public static XElement Face_Data(Face face)
        {
            var faceElement = new XElement("Face");

            if (face == null)
            {
                faceElement.Add(new XAttribute("Error", "null face"));
                return faceElement;
            }

            try { faceElement.Add(new XAttribute("Type", face.Type.ToString())); }
            catch (Exception ex) { Console.WriteLine($"Face Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { faceElement.Add(new XAttribute("Area", face.Area)); }
            catch (Exception ex) { Console.WriteLine($"Face Area: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { faceElement.Add(new XAttribute("GeometryForm", face.GeometryForm)); }
            catch (Exception ex) { Console.WriteLine($"Face GeometryForm: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { faceElement.Add(new XAttribute("Continuity", face.Continuity)); }
            catch (Exception ex) { Console.WriteLine($"Face Continuity: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { faceElement.Add(new XAttribute("ID", face.ID)); }
            catch (Exception ex) { Console.WriteLine($"Face ID: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { faceElement.Add(new XAttribute("Tag", face.Tag)); }
            catch (Exception ex) { Console.WriteLine($"Face Tag: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try
            {
                Array minRange = Array.CreateInstance(typeof(double), 0);
                Array maxRange = Array.CreateInstance(typeof(double), 0);
                face.GetRange(ref minRange, ref maxRange);
                if (minRange.Length >= 3 && maxRange.Length >= 3)
                {
                    faceElement.Add(new XElement("Range",
                        new XAttribute("MinX", minRange.GetValue(0)), new XAttribute("MinY", minRange.GetValue(1)), new XAttribute("MinZ", minRange.GetValue(2)),
                        new XAttribute("MaxX", maxRange.GetValue(0)), new XAttribute("MaxY", maxRange.GetValue(1)), new XAttribute("MaxZ", maxRange.GetValue(2))));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Face GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            return faceElement;
        }
    }
}