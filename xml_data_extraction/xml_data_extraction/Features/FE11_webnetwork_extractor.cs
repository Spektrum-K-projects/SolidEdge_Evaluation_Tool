using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE11_webnetwork_extractor
    {
        public static XElement WebNetwork(WebNetwork webnetwork)
        {
            XElement webNetworkElements = new XElement("Web_Network", new XAttribute("Type", 1718424353));

            try
            {
                webNetworkElements.Add(new XElement("name", webnetwork.Name));
                webNetworkElements.Add(new XElement("type", webnetwork.Type));
                webNetworkElements.Add(new XElement("thickness", webnetwork.Thickness));
                webNetworkElements.Add(new XElement("fintieDepth", webnetwork.FiniteDepth));
                webNetworkElements.Add(new XElement("profileExtensionType", webnetwork.ProfileExtensionType));
                webNetworkElements.Add(new XElement("webDirection", webnetwork.WebDirection));
                webNetworkElements.Add(new XElement("modelingModeType", webnetwork.ModelingModeType));
                webNetworkElements.Add(new XElement("extentType", webnetwork.ExtentType));

                var profile_extract = GE04_getProfiles_extractor.getProfile_extract(webnetwork);
                webNetworkElements.Add(profile_extract);

                webnetwork.GetDraft(out DraftSideConstants draftSide, out double draftAngle);
                webNetworkElements.Add(new XElement("draftSide", draftSide));
                webNetworkElements.Add(new XElement("draftAngle", draftAngle));

                try
                {
                    dynamic dynWebNetwork = webnetwork;
                    webNetworkElements.Add(new XElement("status", dynWebNetwork.Status));
                    webNetworkElements.Add(new XElement("suppress", dynWebNetwork.Suppress));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WebNetwork GetStatus/Suppress: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = webnetwork.GetStatusEx(out object description);
                    webNetworkElements.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WebNetwork GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    Array dims = Array.CreateInstance(typeof(object), 0);
                    webnetwork.GetDimensions(out int numDims, ref dims);
                    webNetworkElements.Add(new XElement("Dimensions", new XAttribute("Count", numDims)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WebNetwork GetDimensions: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    webnetwork.Range(out double x1, out double y1, out double z1, out double x2, out double y2, out double z2);
                    webNetworkElements.Add(new XElement("Range",
                        new XAttribute("X1", x1), new XAttribute("Y1", y1), new XAttribute("Z1", z1),
                        new XAttribute("X2", x2), new XAttribute("Y2", y2), new XAttribute("Z2", z2)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WebNetwork GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Web Network: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
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

            Console.WriteLine("Created Web Networks XML list");
            return webNetworkElements;
        }
    }
}