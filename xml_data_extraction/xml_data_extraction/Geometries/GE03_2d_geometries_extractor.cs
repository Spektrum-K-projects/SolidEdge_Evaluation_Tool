using SolidEdgeFrameworkSupport;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace xml_data_extraction.Geometries
{
    internal class GE03_2d_geometries_extractor
    {
        //Lines2d Extract
        public static XElement Line2d_extract(dynamic? profile)
        {
            Lines2d seLines = null;
            Line2d seLine = null;

            XElement lines2dElement = new XElement("Lines2d");

            try
            {
                seLines = (Lines2d)profile.Lines2d;

                for (int k = 1; k <= seLines.Count; k++)
                {
                    seLine = seLines.Item(k);

                    lines2dElement.Add(new XElement("index", seLine.Index));
                    //Console.WriteLine($"Line Index [{k}]: {seLine.Index}");

                    lines2dElement.Add(new XElement("name", seLine.Name?.ToString()));
                    //Console.WriteLine($"Line Name [{k}]: {seLine.Name}");

                    lines2dElement.Add(new XElement("length", seLine.Length));
                    //Console.WriteLine($"Line length [{k}]: {seLine.Length}");

                    lines2dElement.Add(new XElement("keypointcount", seLine.KeyPointCount));
                    //Console.WriteLine($"Line Key Point Count [{k}]: {seLine.KeyPointCount}");

                    lines2dElement.Add(new XElement("type", seLine.Type));
                    //Console.WriteLine($"Line type [{k}]: {seLine.Type}");

                    Relationships2d lines2drelationships = seLine.Relationships;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lines2d Error Message:{ex.Message}");
            }
            finally
            {
                if (seLine != null)
                {
                    Marshal.ReleaseComObject(seLine);
                    seLine = null;
                }
                if (seLines != null)
                {
                    Marshal.ReleaseComObject(seLines);
                    seLines = null;
                }
            }

            Console.WriteLine($"Created Lines XML list");

            return lines2dElement;
        }

        //Circle2d Extract
        public static XElement Circle2d_extract(dynamic? profile)
        {
            Circles2d seCircles = null;
            Circle2d seCircle = null;

            XElement circles2dElement = new XElement("Circles2d");

            try
            {
                seCircles = (Circles2d)profile.Circles2d;

                for (int l = 1; l <= seCircles.Count; l++)
                {
                    seCircle = seCircles.Item(l);

                    circles2dElement.Add(new XElement("index", seCircle.Index));
                    //Console.WriteLine($"Circle Index [{l}]: {seCircle.Index}");

                    circles2dElement.Add(new XElement("name", seCircle.Name?.ToString()));
                    //Console.WriteLine($"Circle Name [{l}]: {seCircle.Name}");

                    circles2dElement.Add(new XElement("Diameter", seCircle.Diameter));
                    //Console.WriteLine($"Circle length [{l}]: {seCircle.Diameter}");

                    circles2dElement.Add(new XElement("keypointcount", seCircle.KeyPointCount));
                    //Console.WriteLine($"Circle Key Point Count [{l}]: {seCircle.KeyPointCount}");

                    circles2dElement.Add(new XElement("type", seCircle.Type));
                    //Console.WriteLine($"Circle type [{l}]: {seCircle.Type}");

                    Relationships2d circles2drelationships = seCircle.Relationships;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Circles2d Error Message:{ex.Message}");
            }
            finally
            {
                if (seCircle != null)
                {
                    Marshal.ReleaseComObject(seCircle);
                    seCircle = null;
                }
                if (seCircles != null)
                {
                    Marshal.ReleaseComObject(seCircles);
                    seCircles = null;
                }
            }

            Console.WriteLine($"Created Circles XML list");

            return circles2dElement;
        }

        //Arc2d
        public static XElement Arc2d_extract(dynamic? profile)
        {
            Arcs2d seArcs = null;
            Arc2d seArc = null;

            XElement arcs2dElement = new XElement("Arcs2d");

            try
            {
                seArcs = (Arcs2d)profile.Arcs2d;

                for (int m = 1; m <= seArcs.Count; m++)
                {
                    seArc = seArcs.Item(m);

                    arcs2dElement.Add(new XElement("index", seArc.Index));
                    //Console.WriteLine($"Arc Index [{m}]: {seArc.Index}");

                    arcs2dElement.Add(new XElement("name", seArc.Name?.ToString()));
                    //Console.WriteLine($"Arc Name [{m}]: {seArc.Name}");

                    arcs2dElement.Add(new XElement("radius", seArc.Radius));
                    //Console.WriteLine($"Arc Radius [{m}]: {seArc.Radius}");

                    arcs2dElement.Add(new XElement("length", seArc.Length));
                    //Console.WriteLine($"Arc Length [{m}]: {seArc.Length}");

                    arcs2dElement.Add(new XElement("orientation", seArc.Orientation));
                    //Console.WriteLine($"Arc Orientation [{m}]: {seArc.Orientation}");

                    arcs2dElement.Add(new XElement("sweepangle", seArc.SweepAngle));
                    //Console.WriteLine($"Arc Sweep Angle [{m}]: {seArc.SweepAngle}");

                    arcs2dElement.Add(new XElement("startangle", seArc.StartAngle));
                    //Console.WriteLine($"Arc Start Angle [{m}]: {seArc.StartAngle}");

                    arcs2dElement.Add(new XElement("keypointcount", seArc.KeyPointCount));
                    //Console.WriteLine($"Arc Key Point Count [{m}]: {seArc.KeyPointCount}");

                    arcs2dElement.Add(new XElement("type", seArc.Type));
                    //Console.WriteLine($"Arc type [{m}]: {seArc.Type}");

                    arcs2dElement.Add(new XElement("zprder", seArc.ZOrder));
                    //Console.WriteLine($"Arc ZOrder [{m}]: {seArc.ZOrder}");

                    Relationships2d arcs2drelationships = seArc.Relationships;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Arcs2d Error Message:{ex.Message}");
            }
            finally
            {
                if (seArc != null)
                {
                    Marshal.ReleaseComObject(seArc);
                    seArc = null;
                }
                if (seArcs != null)
                {
                    Marshal.ReleaseComObject(seArcs);
                    seArcs = null;
                }
            }

            Console.WriteLine($"Created Arcs XML list");

            return arcs2dElement;
        }

        //Ellipse Extract
        public static XElement Ellipse2d_extract(dynamic? profile)
        {
            Ellipses2d seEllipses = null;
            Ellipse2d seEllipse = null;

            XElement ellipses2dElement = new XElement("Ellipses2d");

            try
            {
                seEllipses = (Ellipses2d)profile.Ellipses2d;

                for (int n = 1; n <= seEllipses.Count; n++)
                {
                    seEllipse = seEllipses.Item(n);

                    ellipses2dElement.Add(new XElement("index", seEllipse.Index));
                    //Console.WriteLine($"Circle Index [{n}]: {seEllipse.Index}");

                    ellipses2dElement.Add(new XElement("name", seEllipse.Name?.ToString()));
                    //Console.WriteLine($"Circle Name [{n}]: {seEllipse.Name}");

                    ellipses2dElement.Add(new XElement("majorradius", seEllipse.MajorRadius));
                    //Console.WriteLine($"Major Radius [{n}]: {seEllipse.MajorRadius}");

                    ellipses2dElement.Add(new XElement("minorradius", seEllipse.MinorRadius));
                    //Console.WriteLine($"Minor Radius [{n}]: {seEllipse.MinorRadius}");

                    ellipses2dElement.Add(new XElement("ratio", seEllipse.MinorMajorRatio));
                    //Console.WriteLine($"Ratio [{n}]: {seEllipse.MinorMajorRatio}");

                    ellipses2dElement.Add(new XElement("orientation", seEllipse.Orientation));
                    //Console.WriteLine($"Ellipse Orientation [{n}]: {seEllipse.Orientation}");

                    ellipses2dElement.Add(new XElement("rotationangle", seEllipse.RotationAngle));
                    //Console.WriteLine($"Rotation Angle [{n}]: {seEllipse.RotationAngle}");

                    ellipses2dElement.Add(new XElement("keypointcount", seEllipse.KeyPointCount));
                    //Console.WriteLine($"Ellipse Key Point Count [{n}]: {seEllipse.KeyPointCount}");

                    ellipses2dElement.Add(new XElement("type", seEllipse.Type));
                    //Console.WriteLine($"Ellipse type [{n}]: {seEllipse.Type}");

                    Relationships2d ellipses2drelationships = seEllipse.Relationships;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Circles2d Error Message:{ex.Message}");
            }
            finally
            {
                if (seEllipse != null)
                {
                    Marshal.ReleaseComObject(seEllipse);
                    seEllipse = null;
                }
                if (seEllipses != null)
                {
                    Marshal.ReleaseComObject(seEllipses);
                    seEllipses = null;
                }
            }

            Console.WriteLine($"Created Circles XML list");

            return ellipses2dElement;
        }
    }
}
