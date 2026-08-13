using System;
using System.Xml.Linq;
using SolidEdgePart;

namespace xml_data_extraction.Properties
{
    internal class PR02_part_physical_properties_extractor
    {
        public static XElement Properties(Model model)
        {
            XElement physicalpropElements = new XElement("PhysicalProperties");

            int status = 0;
            double density = 0, accuracy = 0, volume = 0, area = 0, mass = 0, relativeAccuracy = 0;

            Array centerOfGravity = Array.CreateInstance(typeof(double), 3);
            Array centerOfVolume = Array.CreateInstance(typeof(double), 3);
            Array globalMOI = Array.CreateInstance(typeof(double), 6);
            Array principalMOI = Array.CreateInstance(typeof(double), 3);
            Array principalAxes = Array.CreateInstance(typeof(double), 9);
            Array radiiOfGyration = Array.CreateInstance(typeof(double), 3);

            try
            {
                model.GetPhysicalProperties(out status, out density, out accuracy, out volume,
                                    out area, out mass, ref centerOfGravity, ref centerOfVolume, ref globalMOI,
                                        ref principalMOI, ref principalAxes, ref radiiOfGyration, out relativeAccuracy);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties GetPhysicalProperties: {ex.Message} | Inner: {ex.InnerException?.Message}");
                return physicalpropElements;
            }

            try { physicalpropElements.Add(new XElement("Status", status)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties Status: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { physicalpropElements.Add(new XElement("Density", density)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties Density: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { physicalpropElements.Add(new XElement("Accuracy", accuracy)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties Accuracy: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { physicalpropElements.Add(new XElement("Volume", volume)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties Volume: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { physicalpropElements.Add(new XElement("Area", area)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties Area: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try { physicalpropElements.Add(new XElement("Mass", mass)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties Mass: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            try
            {
                physicalpropElements.Add(new XElement("CenterofGravity",
                    new XAttribute("CoGX", centerOfGravity.GetValue(0)),
                    new XAttribute("CoGY", centerOfGravity.GetValue(1)),
                    new XAttribute("CoGZ", centerOfGravity.GetValue(2))));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties CenterOfGravity: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            try
            {
                physicalpropElements.Add(new XElement("CenterofVolume",
                    new XAttribute("CoVX", centerOfVolume.GetValue(0)),
                    new XAttribute("CoVY", centerOfVolume.GetValue(1)),
                    new XAttribute("CoVZ", centerOfVolume.GetValue(2))));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties CenterOfVolume: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            try
            {
                physicalpropElements.Add(new XElement("GlobalMomentofInertia",
                    new XAttribute("Ixx", globalMOI.GetValue(0)),
                    new XAttribute("Iyy", globalMOI.GetValue(1)),
                    new XAttribute("Izz", globalMOI.GetValue(2)),
                    new XAttribute("Ixy", globalMOI.GetValue(3)),
                    new XAttribute("Ixz", globalMOI.GetValue(4)),
                    new XAttribute("Iyz", globalMOI.GetValue(5))));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties GlobalMomentOfInertia: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            try
            {
                physicalpropElements.Add(new XElement("PrincipalMomentofInertia",
                    new XAttribute("Ix", principalMOI.GetValue(0)),
                    new XAttribute("Iy", principalMOI.GetValue(1)),
                    new XAttribute("Iz", principalMOI.GetValue(2))));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties PrincipalMomentOfInertia: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            try
            {
                physicalpropElements.Add(new XElement("PrincipalAxes",
                    new XAttribute("Pxx", principalAxes.GetValue(0)), new XAttribute("Pxy", principalAxes.GetValue(1)), new XAttribute("Pxz", principalAxes.GetValue(2)),
                    new XAttribute("Pyx", principalAxes.GetValue(3)), new XAttribute("Pyy", principalAxes.GetValue(4)), new XAttribute("Pyz", principalAxes.GetValue(5)),
                    new XAttribute("Pzx", principalAxes.GetValue(6)), new XAttribute("Pzy", principalAxes.GetValue(7)), new XAttribute("Pzz", principalAxes.GetValue(8))));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties PrincipalAxes: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            try
            {
                physicalpropElements.Add(new XElement("RadiiofGyration",
                    new XAttribute("Rx", radiiOfGyration.GetValue(0)),
                    new XAttribute("Ry", radiiOfGyration.GetValue(1)),
                    new XAttribute("Rz", radiiOfGyration.GetValue(2))));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PhysicalProperties RadiiOfGyration: {ex.Message} | Inner: {ex.InnerException?.Message}");
            }

            try { physicalpropElements.Add(new XElement("RelativeAccuracyAchieved", relativeAccuracy)); }
            catch (Exception ex) { Console.WriteLine($"PhysicalProperties RelativeAccuracyAchieved: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

            return physicalpropElements;
        }
    }
}