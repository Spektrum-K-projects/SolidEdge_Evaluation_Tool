using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;

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

    }
}
