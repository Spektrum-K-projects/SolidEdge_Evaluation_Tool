using SolidEdgePart;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE11_webnetwork_extractor
    {
        public static XElement WebNetwork (WebNetwork webnetwork)
        {
            XElement webNetworkElements = new XElement("Web_Network", new XAttribute("Type", 1718424353));

            try
            {
                var name = webnetwork.Name;
                webNetworkElements.Add(new XElement("name", name));

                var type = webnetwork.Type;
                webNetworkElements.Add(new XElement("type", type));

                var thickness = webnetwork.Thickness;
                webNetworkElements.Add(new XElement("thickness", thickness));

                var fintieDepth = webnetwork.FiniteDepth;
                webNetworkElements.Add(new XElement("fintieDepth", fintieDepth));

                var profileExtensionType = webnetwork.ProfileExtensionType;
                webNetworkElements.Add(new XElement("profileExtensionType", profileExtensionType));

                var webDirection = webnetwork.WebDirection;
                webNetworkElements.Add(new XElement("webDirection", webDirection));

                var modelingModeType = webnetwork.ModelingModeType;
                webNetworkElements.Add(new XElement("modelingModeType", modelingModeType));

                var extentType = webnetwork.ExtentType;
                webNetworkElements.Add(new XElement("extentType", extentType));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(webnetwork);
                webNetworkElements.Add(profile_extract);

                webnetwork.GetDraft(out DraftSideConstants draftSide, out double draftAngle);
                webNetworkElements.Add(new XElement("draftSide", draftSide));
                webNetworkElements.Add(new XElement("draftAngle", draftAngle));

                //Profile profile = webnetwork.Profile;

                //XElement profileElement = new XElement("Profiles");
                //profileElement.Add(new XElement("profile_name", profile.Name));
                //profileElement.Add(new XElement("profile_type", profile.Type));

                //var dim_extract = GE01_dimensions_extractor.Dimension_extract(profile);
                //profileElement.Add(dim_extract);

                //Marshal.ReleaseComObject(profile);

            }

            catch (Exception ex)
            {
                Console.WriteLine($"Web Network: Error Message:{ex.Message}");
                return new XElement("Web_Network", "Error");
            }

            finally
            {
                if (webnetwork != null)
                {
                    Marshal.ReleaseComObject(webnetwork);
                    webnetwork = null;
                }
            }

            Console.WriteLine($"Created Web Networks XML list");
            return webNetworkElements;
        }
    }
}

