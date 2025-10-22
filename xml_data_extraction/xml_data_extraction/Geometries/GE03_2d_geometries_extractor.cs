using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
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

            XElement lines2dElement = null;

            try
            {
                seLines = (Lines2d)profile.Lines2d;

                if (seLines != null && seLines.Count > 0)
                { 

                    lines2dElement = new XElement("Lines2d");

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

                        lines2dElement.Add(new XElement("angle", seLine.Angle));

                        lines2dElement.Add(new XElement("isProjection", seLine.Projection));

                        lines2dElement.Add(new XElement("isChamfer", seLine.IsChamfer));

                        lines2dElement.Add(new XElement("type", seLine.Type));
                        //Console.WriteLine($"Line type [{k}]: {seLine.Type}");

                        lines2dElement.Add(new XElement("zOrder", seLine.ZOrder));

                        seLine.GetStartPoint(out double strX, out double strY);
                        lines2dElement.Add(new XElement("startPoint", new XAttribute("Sx", strX),
                                                                        new XAttribute("Sy", strY)));

                        seLine.GetEndPoint(out double endX, out double endY);
                        lines2dElement.Add(new XElement("endPoint", new XAttribute("Ex", endX),
                                                                        new XAttribute("Ey", endY)));

                        for (int f = 0; f < seLine.KeyPointCount; f++)
                        {
                            seLine.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            lines2dElement.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        Relationships2d lines2drelationships = seLine.Relationships;

                    }
                    Console.WriteLine($"Created Line2d XML list");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Lines2d Error Message:{ex.Message}");
                return new XElement("Lines2d", "Error");
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
            
            return lines2dElement;
        }

        //Circle2d Extract
        public static XElement Circle2d_extract(dynamic? profile)
        {
            Circles2d seCircles = null;
            Circle2d seCircle = null;

            XElement circles2dElement = null;

            try
            {
                seCircles = (Circles2d)profile.Circles2d;

                if (seCircles != null && seCircles.Count > 0)
                { 

                    circles2dElement = new XElement("Circles2d");

                    for (int l = 1; l <= seCircles.Count; l++)
                    {
                        seCircle = seCircles.Item(l);

                        circles2dElement.Add(new XElement("index", seCircle.Index));
                        //Console.WriteLine($"Circle Index [{l}]: {seCircle.Index}");

                        circles2dElement.Add(new XElement("name", seCircle.Name?.ToString()));
                        //Console.WriteLine($"Circle Name [{l}]: {seCircle.Name}");

                        circles2dElement.Add(new XElement("Diameter", seCircle.Diameter));
                        //Console.WriteLine($"Circle length [{l}]: {seCircle.Diameter}");

                        circles2dElement.Add(new XElement("Circumference", seCircle.Circumference));

                        circles2dElement.Add(new XElement("Area", seCircle.Area));

                        circles2dElement.Add(new XElement("Length", seCircle.Length));

                        circles2dElement.Add(new XElement("Radius", seCircle.Radius));

                        circles2dElement.Add(new XElement("keypointcount", seCircle.KeyPointCount));
                        //Console.WriteLine($"Circle Key Point Count [{l}]: {seCircle.KeyPointCount}");

                        circles2dElement.Add(new XElement("type", seCircle.Type));
                        //Console.WriteLine($"Circle type [{l}]: {seCircle.Type}");

                        circles2dElement.Add(new XElement("zOrder", seCircle.ZOrder));

                        seCircle.GetCenterPoint(out double x, out double y);
                        circles2dElement.Add(new XElement("centerpoint", new XAttribute("Xcoord", x),
                                                                        new XAttribute("Ycoord", y)));

                        seCircle.GetMajorAxis(out double majX, out double majY);
                        circles2dElement.Add(new XElement("majorAxis", new XAttribute("Mx", majX),
                                                                        new XAttribute("My", majY)));

                        for (int f = 0; f < seCircle.KeyPointCount; f++)
                        {
                            seCircle.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            circles2dElement.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        Relationships2d circles2drelationships = seCircle.Relationships;
                    }
                    Console.WriteLine($"Created Circle2d XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Circles2d Error Message:{ex.Message}");
                return null;
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
                        
            return circles2dElement;
        }

        //Arc2d
        public static XElement Arc2d_extract(dynamic? profile)
        {
            Arcs2d seArcs = null;
            Arc2d seArc = null;

            XElement arcs2dElement = null;

            try
            {
                seArcs = (Arcs2d)profile.Arcs2d;

                if (seArcs != null && seArcs.Count > 0)
                {

                    arcs2dElement = new XElement("Arcs2d");

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

                        arcs2dElement.Add(new XElement("zorder", seArc.ZOrder));
                        //Console.WriteLine($"Arc ZOrder [{m}]: {seArc.ZOrder}");

                        arcs2dElement.Add(new XElement("isFillet", seArc.IsFillet));

                        //Get Center Point
                        //arcs2dElement.Add(Miscellaneous_Methods.MM01_geometry_methods.GetCenterPoint2d(seArc));

                        seArc.GetCenterPoint(out double cenX, out double cenY);
                        arcs2dElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX),
                                                                        new XAttribute("Cy", cenY)));

                        seArc.GetStartPoint(out double strX, out double strY);
                        arcs2dElement.Add(new XElement("startPoint", new XAttribute("Sx", strX),
                                                                        new XAttribute("Sy", strY)));

                        seArc.GetEndPoint(out double endX, out double endY);
                        arcs2dElement.Add(new XElement("endPoint", new XAttribute("Ex", endX),
                                                                        new XAttribute("Ey", endY)));

                        seArc.GetMajorAxis(out double majX, out double majY);
                        arcs2dElement.Add(new XElement("majorAxis", new XAttribute("Mx", majX),
                                                                        new XAttribute("My", majY)));

                        for (int f = 0; f < seArc.KeyPointCount; f++)
                        {
                            seArc.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            arcs2dElement.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        //seArc.GetKeyPoint(index_ed, out double keyX, out double keyY, out double keyZ, 
                        //                                                out KeyPointType keyPointType, out int handleType);
                        //arcs2dElement.Add(new XElement("centerpoint", new XAttribute("Ex", endX),
                        //                                                new XAttribute("Ey", endY)));



                        Relationships2d arcs2drelationships = seArc.Relationships;
                    }
                    Console.WriteLine($"Created Arc2d XML list");
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

            return arcs2dElement;
        }

        //Ellipse Extract
        public static XElement Ellipse2d_extract(dynamic? profile)
        {
            Ellipses2d seEllipses = null;
            Ellipse2d seEllipse = null;

            XElement ellipses2dElement = null;

            try
            {
                seEllipses = (Ellipses2d)profile.Ellipses2d;

                if (seEllipses != null && seEllipses.Count > 0)
                {

                    ellipses2dElement = new XElement("Ellipses2d");

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

                        ellipses2dElement.Add(new XElement("length", seEllipse.Length));

                        ellipses2dElement.Add(new XElement("circumference", seEllipse.Circumference));

                        ellipses2dElement.Add(new XElement("area", seEllipse.Area));

                        ellipses2dElement.Add(new XElement("orientation", seEllipse.Orientation));
                        //Console.WriteLine($"Ellipse Orientation [{n}]: {seEllipse.Orientation}");

                        ellipses2dElement.Add(new XElement("rotationangle", seEllipse.RotationAngle));
                        //Console.WriteLine($"Rotation Angle [{n}]: {seEllipse.RotationAngle}");

                        ellipses2dElement.Add(new XElement("keypointcount", seEllipse.KeyPointCount));
                        //Console.WriteLine($"Ellipse Key Point Count [{n}]: {seEllipse.KeyPointCount}");

                        ellipses2dElement.Add(new XElement("type", seEllipse.Type));
                        //Console.WriteLine($"Ellipse type [{n}]: {seEllipse.Type}");

                        ellipses2dElement.Add(new XElement("zOrder", seEllipse.ZOrder));

                        seEllipse.GetCenterPoint(out double cenX, out double cenY);
                        ellipses2dElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX),
                                                                        new XAttribute("Cy", cenY)));

                        seEllipse.GetMajorAxis(out double majX, out double majY);
                        ellipses2dElement.Add(new XElement("majorAxis", new XAttribute("MAx", majX),
                                                                        new XAttribute("MAy", majY)));

                        seEllipse.GetMinorAxis(out double minX, out double minY);
                        ellipses2dElement.Add(new XElement("minorAxis", new XAttribute("MIx", minX),
                                                                        new XAttribute("MIy", minY)));

                        for (int f = 0; f < seEllipse.KeyPointCount; f++)
                        {
                            seEllipse.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            ellipses2dElement.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        Relationships2d ellipses2drelationships = seEllipse.Relationships;
                    }
                    Console.WriteLine($"Created Ellipse2d XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ellipse2d Error Message:{ex.Message}");
                return new XElement("Ellipse2d", "error");
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
            return ellipses2dElement;
        }

        //Elliptical Arcs Extract
        public static XElement EllipticalArc2d_extract(dynamic? profile)
        {
            EllipticalArcs2d seEllipticalArcs = null;
            EllipticalArc2d seEllipticalArc = null;

            XElement ellipticalArc2dElement = null;

            try
            {
                seEllipticalArcs = (EllipticalArcs2d)profile.ElliticalArcs2d;

                if (seEllipticalArcs != null && seEllipticalArcs.Count > 0)
                {

                    ellipticalArc2dElement = new XElement("EllipticalArc2d");

                    for (int n = 1; n <= seEllipticalArcs.Count; n++)
                    {
                        seEllipticalArc = seEllipticalArcs.Item(n);

                        ellipticalArc2dElement.Add(new XElement("index", seEllipticalArc.Index));
                        //Console.WriteLine($"Circle Index [{n}]: {seEllipse.Index}");

                        ellipticalArc2dElement.Add(new XElement("name", seEllipticalArc.Name?.ToString()));
                        //Console.WriteLine($"Circle Name [{n}]: {seEllipse.Name}");

                        ellipticalArc2dElement.Add(new XElement("majorradius", seEllipticalArc.MajorRadius));
                        //Console.WriteLine($"Major Radius [{n}]: {seEllipse.MajorRadius}");

                        ellipticalArc2dElement.Add(new XElement("minorradius", seEllipticalArc.MinorRadius));
                        //Console.WriteLine($"Minor Radius [{n}]: {seEllipse.MinorRadius}");

                        ellipticalArc2dElement.Add(new XElement("length", seEllipticalArc.Length));

                        ellipticalArc2dElement.Add(new XElement("sweepAngle", seEllipticalArc.SweepAngle));

                        ellipticalArc2dElement.Add(new XElement("startAngle", seEllipticalArc.StartAngle));

                        ellipticalArc2dElement.Add(new XElement("orientation", seEllipticalArc.Orientation));
                        //Console.WriteLine($"Ellipse Orientation [{n}]: {seEllipse.Orientation}");

                        ellipticalArc2dElement.Add(new XElement("rotationangle", seEllipticalArc.RotationAngle));
                        //Console.WriteLine($"Rotation Angle [{n}]: {seEllipse.RotationAngle}");

                        ellipticalArc2dElement.Add(new XElement("keypointcount", seEllipticalArc.KeyPointCount));
                        //Console.WriteLine($"Ellipse Key Point Count [{n}]: {seEllipse.KeyPointCount}");

                        ellipticalArc2dElement.Add(new XElement("type", seEllipticalArc.Type));
                        //Console.WriteLine($"Ellipse type [{n}]: {seEllipse.Type}");

                        ellipticalArc2dElement.Add(new XElement("zOrder", seEllipticalArc.ZOrder));

                        seEllipticalArc.GetCenterPoint(out double cenX, out double cenY);
                        ellipticalArc2dElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX),
                                                                        new XAttribute("Cy", cenY)));

                        seEllipticalArc.GetStartPoint(out double strX, out double strY);
                        ellipticalArc2dElement.Add(new XElement("startPoint", new XAttribute("Sx", strX),
                                                                        new XAttribute("Sy", strY)));

                        seEllipticalArc.GetEndPoint(out double endX, out double endY);
                        ellipticalArc2dElement.Add(new XElement("endPoint", new XAttribute("Ex", endX),
                                                                        new XAttribute("Ey", endY)));

                        seEllipticalArc.GetMajorAxis(out double majX, out double majY);
                        ellipticalArc2dElement.Add(new XElement("majorAxis", new XAttribute("MAx", majX),
                                                                        new XAttribute("MAy", majY)));

                        seEllipticalArc.GetMinorAxis(out double minX, out double minY);
                        ellipticalArc2dElement.Add(new XElement("minorAxis", new XAttribute("MIx", minX),
                                                                        new XAttribute("MIy", minY)));

                        for (int f = 0; f < seEllipticalArc.KeyPointCount; f++)
                        {
                            seEllipticalArc.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            ellipticalArc2dElement.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        Relationships2d ellipticalArcdrelationships = seEllipticalArc.Relationships;
                    }
                    Console.WriteLine($"Created EllipticalArcs2d XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EllipticalArc2d Error Message:{ex.Message}");
                return new XElement("EllipticalArc2d", "error");
            }
            finally
            {
                if (seEllipticalArc != null)
                {
                    Marshal.ReleaseComObject(seEllipticalArc);
                    seEllipticalArc = null;
                }
                if (seEllipticalArcs != null)
                {
                    Marshal.ReleaseComObject(seEllipticalArcs);
                    seEllipticalArcs = null;
                }
            }
            return ellipticalArc2dElement;
        }

        //B-SPline Curves Extract
        public static XElement BSplineCurve2d_extract(dynamic? profile)
        {
            BSplineCurves2d seBSplines = null;
            BSplineCurve2d seBSpline = null;

            XElement bSplineCurve2dElement = null;

            try
            {
                seBSplines = (BSplineCurves2d)profile.BSplinesCurves2d;

                if (seBSplines != null && seBSplines.Count > 0)
                { 

                    bSplineCurve2dElement = new XElement("BSplineCurve2d");

                    for (int k = 1; k <= seBSplines.Count; k++)
                    {
                        seBSpline = seBSplines.Item(k);

                        bSplineCurve2dElement.Add(new XElement("index", seBSpline.Index));
                        //Console.WriteLine($"Line Index [{k}]: {seLine.Index}");

                        bSplineCurve2dElement.Add(new XElement("name", seBSpline.Name?.ToString()));
                        //Console.WriteLine($"Line Name [{k}]: {seLine.Name}");

                        bSplineCurve2dElement.Add(new XElement("length", seBSpline.Length));
                        //Console.WriteLine($"Line length [{k}]: {seLine.Length}");

                        bSplineCurve2dElement.Add(new XElement("area", seBSpline.Area));

                        bSplineCurve2dElement.Add(new XElement("keyPointCount", seBSpline.KeyPointCount));
                        //Console.WriteLine($"Line Key Point Count [{k}]: {seLine.KeyPointCount}");

                        bSplineCurve2dElement.Add(new XElement("nodeCount", seBSpline.NodeCount));

                        bSplineCurve2dElement.Add(new XElement("poleCount", seBSpline.PoleCount));

                        bSplineCurve2dElement.Add(new XElement("segmentedStyleCount", seBSpline.SegmentedStyleCount));

                        bSplineCurve2dElement.Add(new XElement("shapeEdit", seBSpline.ShapeEdit));

                        bSplineCurve2dElement.Add(new XElement("isTangentaillyClosedCurve", seBSpline.IsTangentiallyClosedCurve));

                        bSplineCurve2dElement.Add(new XElement("flexible", seBSpline.Flexible));

                        bSplineCurve2dElement.Add(new XElement("type", seBSpline.Type));
                        //Console.WriteLine($"Line type [{k}]: {seLine.Type}");

                        bSplineCurve2dElement.Add(new XElement("zOrder", seBSpline.ZOrder));

                        seBSpline.GetCentroid(out double cenX, out double cenY);
                        bSplineCurve2dElement.Add(new XElement("centroid", new XAttribute("Cx", cenX),
                                                                        new XAttribute("Cy", cenY)));

                        for (int a = 0; a < seBSpline.NodeCount; a++)
                        {
                            seBSpline.GetNode(a, out double nodeX, out double nodeY);
                            bSplineCurve2dElement.Add(new XElement("node", new XAttribute("Nindex", a),
                                                                            new XAttribute("Nx", nodeX),
                                                                            new XAttribute("Ny", nodeY)));
                        }

                        for (int b = 0; b < seBSpline.PoleCount; b++)
                        {
                            seBSpline.GetPole(b, out double poleX, out double poleY);
                            bSplineCurve2dElement.Add(new XElement("node", new XAttribute("Pindex", b),
                                                                            new XAttribute("Px", poleX),
                                                                            new XAttribute("Py", poleY)));
                        }

                        for (int f = 0; f < seBSpline.KeyPointCount; f++)
                        {
                            seBSpline.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            bSplineCurve2dElement.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        //for (int d = 0; d < seBSpline.SegmentedStyleCount; d++)
                        //{
                        //    seBSpline.GetSegmentedStyle(d, out double x1, out double y1, out double x2, out double y2, out object Style);

                        //    bSplineCurve2dElement.Add(new XElement("segmentedStyle", new XAttribute("Sindex", d),
                        //                                                    new XAttribute("X1", x1),
                        //                                                    new XAttribute("Y1", y1),
                        //                                                    new XAttribute("X2", x2),
                        //                                                    new XAttribute("Y2", y2)));
                        //}

                        Relationships2d bSplineCurve2drelationships = seBSpline.Relationships;
                    }
                    Console.WriteLine($"Created B-Spline Curves2d XML list");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"BSplinesCurve2d Error Message:{ex.Message}");
                return new XElement("BSplineCurve2d", "Error");
            }

            finally
            {
                if (seBSpline != null)
                {
                    Marshal.ReleaseComObject(seBSpline);
                    seBSpline = null;
                }
                if (seBSplines != null)
                {
                    Marshal.ReleaseComObject(seBSplines);
                    seBSplines = null;
                }
            }

            return bSplineCurve2dElement;
        }

        //Conics Extract
        public static XElement Conic2d_extract(dynamic? profile)
        {
            Conics2d seConics = null;
            Conic2d seConic = null;

            XElement conic2dElements = null;

            try
            {
                seConics = (Conics2d)profile.Conics2d;

                if (seConics! != null && seConics.Count > 0)
                { 

                    conic2dElements = new XElement("Conic2d");

                    for (int k = 1; k <= seConics.Count; k++)
                    {
                        seConic = seConics.Item(k);

                        conic2dElements.Add(new XElement("index", seConic.Index));

                        conic2dElements.Add(new XElement("name", seConic.Name?.ToString()));

                        conic2dElements.Add(new XElement("rhoValue", seConic.RhoValue));

                        conic2dElements.Add(new XElement("type", seConic.Type));

                        conic2dElements.Add(new XElement("zOrder", seConic.ZOrder));

                        seConic.GetControlPoint(out double conX, out double conY);
                        conic2dElements.Add(new XElement("controlPoint", new XAttribute("CPx", conX),
                                                                        new XAttribute("CPy", conY)));

                        seConic.GetStartPoint(out double strX, out double strY);
                        conic2dElements.Add(new XElement("startPoint", new XAttribute("Sx", strX),
                                                                        new XAttribute("Sy", strY)));

                        seConic.GetEndPoint(out double endX, out double endY);
                        conic2dElements.Add(new XElement("endPoint", new XAttribute("Ex", endX),
                                                                        new XAttribute("Ey", endY)));

                        Relationships2d conic2drelationships = seConic.Relationships;
                    }
                    Console.WriteLine($"Created Conics2d XML list");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Conics2d Error Message:{ex.Message}");
                return new XElement("Conic2d", "Error");
            }

            finally
            {
                if (seConic != null)
                {
                    Marshal.ReleaseComObject(seConic);
                    seConic = null;
                }
                if (seConics != null)
                {
                    Marshal.ReleaseComObject(seConics);
                    seConics = null;
                }
            }

            return conic2dElements;
        }

        //Points2d Extract
        public static XElement Point2d_extract(dynamic? profile)
        {
            Points2d sePoints = null;
            Point2d sePoint = null;

            XElement point2dElements = null;

            try
            {
                sePoints = (Points2d)profile.Points2d;

                if (sePoints! != null && sePoints.Count > 0)
                {

                    point2dElements = new XElement("Point2d");

                    for (int k = 1; k <= sePoints.Count; k++)
                    {
                        sePoint = sePoints.Item(k);

                        point2dElements.Add(new XElement("index", sePoint.Index));

                        point2dElements.Add(new XElement("name", sePoint.Name?.ToString()));

                        point2dElements.Add(new XElement("x-Coordinate", sePoint.x));

                        point2dElements.Add(new XElement("y-Coordinate", sePoint.y));

                        point2dElements.Add(new XElement("keyPointCount", sePoint.KeyPointCount));

                        point2dElements.Add(new XElement("type", sePoint.Type));

                        point2dElements.Add(new XElement("zOrder", sePoint.ZOrder));

                        for (int f = 0; f < sePoint.KeyPointCount; f++)
                        {
                            sePoint.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            point2dElements.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        sePoint.GetPoint(out double pX, out double pY);
                        point2dElements.Add(new XElement("controlPoint", new XAttribute("Px", pX),
                                                                        new XAttribute("Py", pY)));

                        Relationships2d cpoint2drelationships = sePoint.Relationships;
                    }
                    Console.WriteLine($"Created Points2d XML list");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine($"Lines2d Error Message:{ex.Message}");
                return new XElement("Point2d", "Error");
            }

            finally
            {
                if (sePoint != null)
                {
                    Marshal.ReleaseComObject(sePoint);
                    sePoint = null;
                }
                if (sePoints != null)
                {
                    Marshal.ReleaseComObject(sePoints);
                    sePoints = null;
                }
            }

            return point2dElements;
        }

        public static XElement Hole2d_extract(dynamic? profile)
        {
            Holes2d seHoles2d = null;
            Hole2d seHole2d = null;

            XElement hole2dElements = null;

           try
           {
                seHoles2d = (Holes2d)profile.Holes2d;

                if (seHoles2d! != null && seHoles2d.Count > 0)
                {
                    hole2dElements = new XElement("Holes2d");
                    for (int k = 1; k <= seHoles2d.Count; k++)
                    {
                        seHole2d = seHoles2d.Item(k);
                        hole2dElements.Add(new XElement("name", seHole2d.Name?.ToString()));
                        hole2dElements.Add(new XElement("type", seHole2d.Type));
                        hole2dElements.Add(new XElement("zOrder", seHole2d.ZOrder));
                        hole2dElements.Add(new XElement("keyPointCount", seHole2d.KeyPointCount));

                        for (int f = 0; f < seHole2d.KeyPointCount; f++)
                        {
                            seHole2d.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                                            out KeyPointType keyPointType, out int handleType);
                            hole2dElements.Add(new XElement("keyPoint", new XAttribute("Kindex", f),
                                                                            new XAttribute("Kx", keyX),
                                                                            new XAttribute("Ky", keyY),
                                                                            new XAttribute("Kz", keyZ),
                                                                            new XAttribute("Ktype", keyPointType),
                                                                            new XAttribute("Htype", handleType)));
                        }

                        seHole2d.GetCenterPoint(out double cenX, out double cenY);
                        hole2dElements.Add(new XElement("centerPoint", new XAttribute("Cx", cenX),
                                                                        new XAttribute("Cy", cenY)));
                    }
                    Console.WriteLine($"Created Holes2d XML list");
                }
           }
           catch (Exception ex)
           {
                Console.WriteLine($"Holes2d Error Message:{ex.Message}");
                return new XElement("Holes2d", "Error");
           }
           finally
           {
                if (seHole2d != null)
                {
                    Marshal.ReleaseComObject(seHole2d);
                    seHole2d = null;
                }
                if (seHoles2d != null)
                {
                    Marshal.ReleaseComObject(seHoles2d);
                    seHoles2d = null;
                }
           }
            return hole2dElements;
        }
    }
}
