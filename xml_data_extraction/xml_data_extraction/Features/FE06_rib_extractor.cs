using SolidEdgePart;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE06_rib_extractor
    {
        public static XElement Rib(Rib rib)
        {
            XElement ribElements = new XElement("Rib", new XAttribute("Type", 462094730));

            try
            {
                var name = rib.Name;
                ribElements.Add(new XElement("name", name));

                var type = rib.Type;
                ribElements.Add(new XElement("type", type));

                var thickness = rib.Thickness;
                ribElements.Add(new XElement("thickness", thickness));

                var thicknessSide = rib.ThicknessSide;
                ribElements.Add(new XElement("thicknessside", thicknessSide));

                var thicknessType = rib.ThicknessType;
                ribElements.Add(new XElement("thicknesstype", thicknessType));

                //SolidEdgeGeometry.Edges edges = rib.Edges;
                //ribElements.Add(new XElement("edges", edges));

                //var faces = rib.Faces;
                //ribElements.Add(new XElement("faces", faces));

                //var facesByRay = rib.FacesByRay;
                //ribElements.Add(new XElement("thickness", thickness));

                var materialSide = rib.MaterialSide;
                ribElements.Add(new XElement("materialside", materialSide));

                var modelingModeType = rib.ModelingModeType;
                ribElements.Add(new XElement("modelingModeType", modelingModeType));

                Profile profile = rib.Profile;

                XElement profileElement = new XElement("Profiles");
                profileElement.Add(new XElement("profile_name", profile.Name));
                profileElement.Add(new XElement("profile_type", profile.Type));

                var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                profileElement.Add(dim_extract);

                Marshal.ReleaseComObject(profile);

            }

            catch (Exception ex) 
            {
                Console.WriteLine($"Ribs: Error Message:{ex.Message}");
                return new XElement("Rib");
            }

            finally
            {
                if (rib != null)
                {
                    Marshal.ReleaseComObject(rib);
                    rib = null;
                }
            }

            Console.WriteLine($"Created Rib XML list");
            return ribElements;
        }
    }
}
