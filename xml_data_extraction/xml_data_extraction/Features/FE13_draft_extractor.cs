using SolidEdgePart;
using System.Dynamic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    //Draft Data Extraction
    internal class FE13_draft_extractor
    {
        public static XElement draft_Data(Draft draft)
        {
            XElement draftElements = new XElement("Draft", new XAttribute("Type", 462094746));

            Array draftAngles = Array.CreateInstance(typeof(double), 0);

            try
            {
                draftElements.Add("name", draft.Name);

                draftElements.Add("type", draft.Type);

                draftElements.Add("modelingModeType", draft.ModelingModeType);

                draftElements.Add("draftSide", draft.DraftSide.ToString());

                draft.GetDraftAngles(out int draftAngleCount, ref draftAngles);
                if (draftAngles.Length > 0)
                {
                    draftElements.Add(new XElement("draftAngleCount", draftAngleCount));

                    foreach (var angles in draftAngles)
                    {
                        draftElements.Add(new XElement("draftAngles", new XAttribute ("values", string.Join(" ", (double[])angles))));
                    }
                }

                //----Add Edges----

                //----Add Faces----

                //----Add getDimensions----
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Thread: Error Message:{ex.Message}");
            }
            finally
            {
                if (draft != null)
                {
                    Marshal.ReleaseComObject(draft);
                    draft = null;
                }
            }

            return draftElements;
        }
    }
}