using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SolidEdgePart;

namespace xml_data_extraction.Properties
{
    internal class PR02_part_physical_properties_extractor
    {
        public static XElement Properties(Model model)
        {
            int status;

            double density,
                    accuracy,
                    volume,
                    area,
                    mass,
                    relativeAccuracy;

            Array centerOfGravity = Array.CreateInstance(typeof(double), 3);
            Array centerOfVolume = Array.CreateInstance(typeof(double), 3);
            Array globalMOI = Array.CreateInstance(typeof(double), 6);
            Array principalMOI = Array.CreateInstance(typeof(double), 3);
            Array principalAxes = Array.CreateInstance(typeof(double), 9);
            Array radiiOfGyration = Array.CreateInstance(typeof(double), 9);

            XElement physicalpropElements = new XElement("PhysicalProperties");

            try
            {
                model.GetPhysicalProperties(out status, out density, out accuracy, out volume,
                                    out area, out mass, ref centerOfGravity, ref centerOfVolume, ref globalMOI,
                                        ref principalMOI, ref principalAxes, ref radiiOfGyration, out relativeAccuracy);

                physicalpropElements.Add(new XElement ("Status", status));
                physicalpropElements.Add(new XElement("Density", density));
                physicalpropElements.Add(new XElement("Accuracy", accuracy));
                physicalpropElements.Add(new XElement("Volume", volume));
                physicalpropElements.Add(new XElement("Area", area));
                physicalpropElements.Add(new XElement("Mass", mass));

                physicalpropElements.Add(new XElement("CenterofGravity", new XAttribute("CoGX", centerOfGravity.GetValue(0)),
                                                                        new XAttribute("CoGY", centerOfGravity.GetValue(1)),
                                                                        new XAttribute("CoGZ", centerOfGravity.GetValue(2))));

                physicalpropElements.Add(new XElement("CenterofVolume", new XAttribute("CoVX", centerOfVolume.GetValue(0)),
                                                                        new XAttribute("CoVY", centerOfVolume.GetValue(1)),
                                                                        new XAttribute("CoVZ", centerOfVolume.GetValue(2))));

                physicalpropElements.Add(new XElement("GlobalMomentofInertia", new XAttribute("Ixx", globalMOI.GetValue(0)),
                                                                            new XAttribute("Iyy", globalMOI.GetValue(1)),
                                                                            new XAttribute("Izz", globalMOI.GetValue(2)),
                                                                            new XAttribute("Ixy", globalMOI.GetValue(3)),
                                                                            new XAttribute("Ixz", globalMOI.GetValue(4)),
                                                                            new XAttribute("Iyz", globalMOI.GetValue(5))));

                physicalpropElements.Add(new XElement("PrincipalMomentofInertia", new XAttribute("Ix", principalMOI.GetValue(0)),
                                                                                new XAttribute("Iy", principalMOI.GetValue(1)),
                                                                                new XAttribute("Iz", principalMOI.GetValue(2))));

                physicalpropElements.Add(new XElement("PrincipalAxes", new XAttribute("Pxx", principalAxes.GetValue(0)),
                                                                    new XAttribute("Pxy", principalAxes.GetValue(1)),
                                                                    new XAttribute("Pxz", principalAxes.GetValue(2)),
                                                                    new XAttribute("Pyx", principalAxes.GetValue(3)),
                                                                    new XAttribute("Pyy", principalAxes.GetValue(4)),
                                                                    new XAttribute("Pyz", principalAxes.GetValue(5)),
                                                                    new XAttribute("Pzx", principalAxes.GetValue(6)),
                                                                    new XAttribute("Pzy", principalAxes.GetValue(7)),
                                                                    new XAttribute("Pzz", principalAxes.GetValue(8))));

                physicalpropElements.Add(new XElement("RadiiofGyration", new XAttribute("Rxx", radiiOfGyration.GetValue(0)),
                                                                        new XAttribute("Rxy", radiiOfGyration.GetValue(1)),
                                                                        new XAttribute("Rxz", radiiOfGyration.GetValue(2)),
                                                                        new XAttribute("Ryx", radiiOfGyration.GetValue(3)),
                                                                        new XAttribute("Ryy", radiiOfGyration.GetValue(4)),
                                                                        new XAttribute("Ryz", radiiOfGyration.GetValue(5)),
                                                                        new XAttribute("Rzx", radiiOfGyration.GetValue(6)),
                                                                        new XAttribute("Rzy", radiiOfGyration.GetValue(7)),
                                                                        new XAttribute("Rzz", radiiOfGyration.GetValue(8))));

                physicalpropElements.Add(new XElement("RelativeAccuracyAchieved", relativeAccuracy));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Physical Properties Error Message: {ex.Message}");
            }

            return physicalpropElements;
        }
    }
}
