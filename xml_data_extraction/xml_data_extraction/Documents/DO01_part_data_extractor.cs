using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using SolidEdgePart;
using xml_data_extraction.Features;
using xml_data_extraction.Properties;

namespace xml_data_extraction.Documents
{
    internal class DO01_part_data_extractor
    {
        public static XElement PartExtract(PartDocument partDoc)
        {
            Models models = null;
            var partElements = new XElement("Part");

            try
            {
                Console.WriteLine("  Extracting features...");
                models = partDoc.Models;

                for (int m = 1; m <= models.Count; m++)
                {
                    Model model = models.Item(m);
                    try
                    {
                        Console.WriteLine($"   Model [{m}]: {model.Name}");

                        partElements.Add(PR02_part_physical_properties_extractor.Properties(model));

                        SolidEdgePart.Features features = model.Features;
                        try
                        {
                            for (int f = 1; f <= features.Count; f++)
                            {
                                var feat = features.Item(f);
                                int output = feat.Type;

                                //Call Feature Types:
                                if (output == 462094706)
                                {
                                    partElements.Add(FE01_protrusion_extractor.Protrusion_Extrude((ExtrudedProtrusion)feat));
                                    continue;
                                }
                                else if (output == -1204891230)
                                {
                                    partElements.Add(FE01_protrusion_extractor.Helix_Protrusion_Extrude((HelixProtrusion)feat));
                                    continue;
                                }
                                else if (output == 462094722)
                                {
                                    partElements.Add(FE02_hole_extractor.Hole((Hole)feat));
                                    continue;
                                }
                                else if (output == 462094714)
                                {
                                    partElements.Add(FE03_cutout_extractor.Cutout_Extrude((ExtrudedCutout)feat));
                                    continue;
                                }
                                else if (output == 1197717883)
                                {
                                    partElements.Add(FE03_cutout_extractor.Helix_Cutout_Extrude((HelixCutout)feat));
                                    continue;
                                }
                                else if (output == 462094742)
                                {
                                    partElements.Add(FE04_edge_features_extractor.Chamfer((Chamfer)feat));
                                    continue;
                                }
                                else if (output == 462094710)
                                {
                                    partElements.Add(FE05_revolve_extractor.Revolved_protrusion((RevolvedProtrusion)feat));
                                    continue;
                                }
                                else if (output == 462094718)
                                {
                                    partElements.Add(FE05_revolve_extractor.Revolved_Cutout((RevolvedCutout)feat));
                                    continue;
                                }
                                else if (output == 462094730)
                                {
                                    partElements.Add(FE06_rib_extractor.Rib((Rib)feat));
                                    continue;
                                }
                                else if (output == 462094738)
                                {
                                    partElements.Add(FE04_edge_features_extractor.Round((Round)feat));
                                    continue;
                                }
                                else if (output == -416228998)
                                {
                                    partElements.Add(FE07_pattern_extractor.Pattern((Pattern)feat));
                                    continue;
                                }
                                else if (output == 462094734)
                                {
                                    partElements.Add(FE08_thinwall_extractor.ThinWall((ThinWall)feat));
                                    continue;
                                }
                                else if (output == 438630050)
                                {
                                    partElements.Add(FE09_thinregion_extractor.ThinRegion((Thin)feat));
                                    continue;
                                }
                                else if (output == 1908287958)
                                {
                                    partElements.Add(FE10_mirror_extractor.MirrorPart((MirrorPart)feat));
                                    continue;
                                }
                                else if (output == 66247736)
                                {
                                    partElements.Add(FE10_mirror_extractor.MirrorCopy((MirrorCopy)feat));
                                    continue;
                                }
                                else if (output == 1718424353)
                                {
                                    partElements.Add(FE11_webnetwork_extractor.WebNetwork((WebNetwork)feat));
                                    continue;
                                }
                                else if (output == -85880079)
                                {
                                    partElements.Add(FE12_vent_extractor.Vent((Vent)feat));
                                    continue;
                                }
                                else if (output == 462094746)
                                {
                                    partElements.Add(FE13_draft_extractor.draft_Data((Draft)feat));
                                    continue;
                                }
                                else if (output == 1336556513)
                                {
                                    partElements.Add(FE03_cutout_extractor.Slot_Extract((Slot)feat));
                                    continue;
                                }
                                else if (output == 339115113)
                                {
                                    partElements.Add(FE03_cutout_extractor.SlotGroup_Extract((SlotGroup)feat));
                                    continue;
                                }
                                else if (output == -292547215)
                                {
                                    partElements.Add(FE03_cutout_extractor.Normal_Cutout((NormalCutout)feat));
                                    continue;
                                }
                                else if (output == -483645223)
                                {
                                    partElements.Add(FE16_lip_extractor.Lip_Extract((Lip)feat));
                                    continue;
                                }
                                else if (output == 2057842144)
                                {
                                    partElements.Add(FE17_loft_extractor.Lofted_Protrusion((LoftedProtrusion)feat));
                                    continue;
                                }
                                else if (output == 2057842149)
                                {
                                    partElements.Add(FE17_loft_extractor.Lofted_Cutout((LoftedCutout)feat));
                                    continue;
                                }
                                else if (output == -2101194894)
                                {
                                    partElements.Add(FE18_sweep_extractor.Swept_Protrusion((SweptProtrusion)feat));
                                    continue;
                                }
                                else if (output == -398746894)
                                {
                                    partElements.Add(FE18_sweep_extractor.Swept_Cutout((SweptCutout)feat));
                                    continue;
                                }
                                else if (output == -2101998503)
                                {
                                    partElements.Add(FE19_emboss_extractor.Emboss_Extract((EmbossFeature)feat));
                                    continue;
                                }
                                else if (output == -1468087919)
                                {
                                    partElements.Add(FE07_pattern_extractor.UserDefinedPattern((UserDefinedPattern)feat));
                                    continue;
                                }

                                //Add more feature types as needed...

                                else
                                {
                                    Console.WriteLine($"Skipping unhandled feature type: {output}");
                                    if (feat != null)
                                    {
                                        Marshal.ReleaseComObject(feat);
                                    }
                                    continue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Error extracting features for Model [{m}]: {ex.Message}");
                        }
                        finally
                        {
                            if (features != null)
                            {
                                Marshal.ReleaseComObject(features);
                                features = null;
                            }
                        }

                        object threadsObj = null;
                        try
                        {
                            var threads = model.Threads;
                            threadsObj = threads;

                            for (int t = 1; t <= threads.Count; t++)
                            {
                                SolidEdgePart.Thread thread = null;
                                try
                                {
                                    thread = threads.Item(t);
                                    partElements.Add(FE20_thread_extractor.Thread_Extract(thread));
                                }
                                finally
                                {
                                    if (thread != null)
                                    {
                                        Marshal.ReleaseComObject(thread);
                                        thread = null;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Error extracting Threads for Model [{m}]: {ex.Message}");
                        }
                        finally
                        {
                            if (threadsObj != null)
                            {
                                Marshal.ReleaseComObject(threadsObj);
                                threadsObj = null;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Error processing Model [{m}]: {ex.Message}");
                    }
                    finally
                    {
                        if (model != null)
                        {
                            Marshal.ReleaseComObject(model);
                            model = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("  PartDataExtractor error: " + ex.Message);
            }
            finally
            {
                if (models != null)
                {
                    Marshal.ReleaseComObject(models);
                    models = null;
                }
            }

            return partElements;
        }
    }
}