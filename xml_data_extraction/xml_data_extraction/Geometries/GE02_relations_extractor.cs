using SolidEdgePart;
using SolidEdgeFrameworkSupport;
using SolidEdgeConstants;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Geometries
{
    internal class GE02_relations_extractor
    {
        public static XElement Relations2d_extract(Profile profile)
        {
            Relations2d seRelations = null;
            XElement relations2dElement = new XElement("Relations2d");

            try
            {
                seRelations = (Relations2d)profile.Relations2d;

                for (int k = 1; k <= seRelations.Count; k++)
                {
                    Relation2d seRelation = null;
                    try
                    {
                        seRelation = seRelations.Item(k);

                        XElement relationElement = new XElement("Relation");

                        try { relationElement.Add(new XElement("index", seRelation.Index)); }
                        catch (Exception ex) { Console.WriteLine($"Relation[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                        try { relationElement.Add(new XElement("name", seRelation.Name?.ToString())); }
                        catch (Exception ex) { Console.WriteLine($"Relation[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                        try { relationElement.Add(new XElement("type", ((ObjectType)seRelation.Type).ToString())); }
                        catch (Exception ex) { Console.WriteLine($"Relation[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                        try { relationElement.Add(new XElement("zOrder", seRelation.ZOrder)); }
                        catch (Exception ex) { Console.WriteLine($"Relation[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                        try { relationElement.Add(new XElement("key", seRelation.Key)); }
                        catch (Exception ex) { Console.WriteLine($"Relation[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                        int relatedObjectCount = 0;
                        try
                        {
                            seRelation.GetRelatedObjectCount(out relatedObjectCount);
                            relationElement.Add(new XElement("relatedObjectCount", relatedObjectCount));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Relation[{k}] RelatedObjectCount: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }

                        if (relatedObjectCount > 0)
                        {
                            XElement relatedObjectsElement = new XElement("RelatedObjects", new XAttribute("Count", relatedObjectCount));

                            for (int r = 1; r <= relatedObjectCount; r++)
                            {
                                object graphicObject = null;
                                try
                                {
                                    seRelation.GetRelatedObject(r, out graphicObject, out int keypointIndex);
                                    relatedObjectsElement.Add(new XElement($"relatedObject{r}",
                                        new XAttribute("UnderlyingType", ClassifyGraphicObject(graphicObject)),
                                        new XAttribute("KeypointIndex", keypointIndex)));
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Relation[{k}] GetRelatedObject[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                }
                                finally
                                {
                                    if (graphicObject != null)
                                    {
                                        Marshal.ReleaseComObject(graphicObject);
                                        graphicObject = null;
                                    }
                                }
                            }

                            relationElement.Add(relatedObjectsElement);
                        }

                        relations2dElement.Add(relationElement);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Relation[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                    }
                    finally
                    {
                        if (seRelation != null)
                        {
                            Marshal.ReleaseComObject(seRelation);
                            seRelation = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Relations2d Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seRelations != null)
                {
                    Marshal.ReleaseComObject(seRelations);
                    seRelations = null;
                }
            }

            Console.WriteLine($"Created Relations XML list");
            return relations2dElement;
        }

        private static string ClassifyGraphicObject(object graphicObject)
        {
            if (graphicObject == null) return "null";

            try
            {
                dynamic dynObject = graphicObject;
                return ((ObjectType)dynObject.Type).ToString();
            }
            catch (Exception ex)
            {
                return $"Unknown ({ex.Message})";
            }
        }
    }
}