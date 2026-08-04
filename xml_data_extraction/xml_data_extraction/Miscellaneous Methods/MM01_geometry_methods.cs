using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace xml_data_extraction.Miscellaneous_Methods
{
    internal class MM01_geometry_methods
    {
        public static XElement GetCenterPoint2d(object geometry2d)
        {
            XElement centerPoint2dElement = new XElement("CenterPoint2d");
            try
            {
                double xCoord;
                double yCoord;

                try
                {
                    if (geometry2d is Arc2d arc2d)
                    {
                        arc2d.GetCenterPoint(out xCoord, out yCoord);
                        centerPoint2dElement.Add(new XElement("X", xCoord));
                        centerPoint2dElement.Add(new XElement("Y", yCoord));
                    }
                    else if (geometry2d is Circle2d profiles)
                    {
                        profiles.GetCenterPoint(out xCoord, out yCoord);
                        centerPoint2dElement.Add(new XElement("X", xCoord));
                        centerPoint2dElement.Add(new XElement("Y", yCoord));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"GetCenterPoint2d profile identification: Error Message:{ex.Message}");
                }
                //var centerPoint2d = geometry2d.GetCenterPoint(out double xCoord, out double yCoord);

                //if (geometry2d != null)
                //{
                //    var centerPoint2d = geometry2d.GetCenterPoint;
                //    if (centerPoint2d != null)
                //    {
                //        centerPoint2dElement.Add(new XElement("X", centerPoint2d.X));
                //        centerPoint2dElement.Add(new XElement("Y", centerPoint2d.Y));
                //    }
                //    else
                //    {
                //        centerPoint2dElement.Add(new XElement("X", "N/A"));
                //        centerPoint2dElement.Add(new XElement("Y", "N/A"));
                //    }
                //}
                //else
                //{
                //    centerPoint2dElement.Add(new XElement("X", "N/A"));
                //    centerPoint2dElement.Add(new XElement("Y", "N/A"));
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetCenterPoint2d: Error Message:{ex.Message}");
                centerPoint2dElement.Add(new XElement("X", "Error"));
                centerPoint2dElement.Add(new XElement("Y", "Error"));
            }
            return centerPoint2dElement;
        }

        public static XElement GetPlaneData(object planeObj, string elementName = "Plane")
        {
            if (planeObj == null)
                return new XElement(elementName, "null");

            try
            {
                if (planeObj is RefPlane refPlane)
                {
                    Array normal = Array.CreateInstance(typeof(double), 0);
                    Array rootPoint = Array.CreateInstance(typeof(double), 0);
                    Array refDir = Array.CreateInstance(typeof(double), 0);

                    refPlane.GetNormal(ref normal);
                    refPlane.GetRootPoint(ref rootPoint);
                    refPlane.GetReferenceDirection(ref refDir);

                    var planeElement = new XElement(elementName,
                        new XAttribute("Name", refPlane.Name ?? "Unnamed"),
                        new XAttribute("Global", refPlane.Global));

                    if (normal.Length >= 3)
                        planeElement.Add(new XElement("Normal",
                            new XAttribute("X", normal.GetValue(0)), new XAttribute("Y", normal.GetValue(1)), new XAttribute("Z", normal.GetValue(2))));
                    if (rootPoint.Length >= 3)
                        planeElement.Add(new XElement("RootPoint",
                            new XAttribute("X", rootPoint.GetValue(0)), new XAttribute("Y", rootPoint.GetValue(1)), new XAttribute("Z", rootPoint.GetValue(2))));
                    if (refDir.Length >= 3)
                        planeElement.Add(new XElement("ReferenceDirection",
                            new XAttribute("X", refDir.GetValue(0)), new XAttribute("Y", refDir.GetValue(1)), new XAttribute("Z", refDir.GetValue(2))));

                    Marshal.ReleaseComObject(refPlane);
                    return planeElement;
                }
                else
                {
                    // Not a RefPlane - most likely a planar Face was picked instead of a reference plane.
                    // Not extracting full face geometry here - just recording what it actually was.
                    return new XElement(elementName, new XAttribute("UnderlyingType", planeObj.GetType().Name));
                }
            }
            catch (Exception ex)
            {
                return new XElement(elementName, new XAttribute("Error", ex.Message));
            }
        }

    }
}
