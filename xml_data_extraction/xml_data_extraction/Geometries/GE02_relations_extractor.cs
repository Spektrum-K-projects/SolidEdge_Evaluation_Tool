using SolidEdgeFrameworkSupport;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Geometries
{
    internal class GE02_relations_extractor
    {
        public static XElement Relations2d_extract(dynamic? profile)
        {
            Relations2d seRelations = null;
            Relation2d seRelation = null;

            XElement relations2dElement = new XElement("Relations2d");

            try
            {
                seRelations = (Relations2d)profile.Relations2d;

                for (int k = 1; k <= seRelations.Count; k++)
                {
                    seRelation = seRelations.Item(k);

                    relations2dElement.Add(new XElement("index", seRelation.Index));
                    //Console.WriteLine($"Relation Index [{k}]: {seRelation.Index}");

                    relations2dElement.Add(new XElement("name", seRelation.Name?.ToString()));
                    //Console.WriteLine($"Relation Name [{k}]: {seRelation.Name}");

                    relations2dElement.Add(new XElement("type", seRelation.Type));
                    //Console.WriteLine($"Relation type [{k}]: {seRelation.Type}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Relations2d Error Message:{ex.Message}");
            }
            finally
            {
                if (seRelation != null)
                {
                    Marshal.ReleaseComObject(seRelation);
                    seRelation = null;
                }
                if (seRelations != null)
                {
                    Marshal.ReleaseComObject(seRelations);
                    seRelations = null;
                }
            }

            Console.WriteLine($"Created Relations XML list");

            return relations2dElement;
        }
    }
}
