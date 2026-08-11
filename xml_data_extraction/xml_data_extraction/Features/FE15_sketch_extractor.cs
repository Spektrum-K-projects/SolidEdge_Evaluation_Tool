using SolidEdgePart;
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using xml_data_extraction.Geometries;

namespace xml_data_extraction.Features
{
    internal class FE15_sketch_extractor
    {
        public static XElement Sketch_Extract(Sketch sketch)
        {
            XElement sketchElement = new XElement("Sketch");

            try
            {
                try { sketchElement.Add(new XElement("name", sketch.Name)); }
                catch (Exception ex) { Console.WriteLine($"Sketch Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sketchElement.Add(new XElement("isUnderDefined", sketch.IsUnderDefined)); }
                catch (Exception ex) { Console.WriteLine($"Sketch IsUnderDefined: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sketchElement.Add(new XElement("locked", sketch.Locked)); }
                catch (Exception ex) { Console.WriteLine($"Sketch Locked: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try { sketchElement.Add(new XElement("modelingModeType", sketch.ModelingModeType)); }
                catch (Exception ex) { Console.WriteLine($"Sketch ModelingModeType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                try
                {
                    dynamic dynSketch = sketch;
                    sketchElement.Add(new XElement("status", dynSketch.Status));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Sketch GetStatus: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    FeatureStatusConstants statusEx = sketch.GetStatusEx(out object description);
                    sketchElement.Add(new XElement("statusEx",
                        new XAttribute("Code", statusEx), new XAttribute("Description", description?.ToString() ?? "")));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Sketch GetStatusEx: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }

                try
                {
                    var profiles = sketch.Profiles;
                    var profilesElement = new XElement("SketchProfiles", new XAttribute("Count", profiles.Count));

                    for (int p = 1; p <= profiles.Count; p++)
                    {
                        var profile = profiles.Item(p);
                        profilesElement.Add(GE04_getProfiles_extractor.Profile_Data(profile));
                        Marshal.ReleaseComObject(profile);
                    }

                    sketchElement.Add(profilesElement);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Sketch GetProfiles: {ex.Message} | Inner: {ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Sketch: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
                return new XElement("Sketch", "Error");
            }
            finally
            {
                if (sketch != null)
                {
                    Marshal.ReleaseComObject(sketch);
                    sketch = null;
                }
            }

            return sketchElement;
        }
    }
}