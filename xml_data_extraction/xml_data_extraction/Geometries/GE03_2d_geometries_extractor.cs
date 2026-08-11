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
        public static XElement Line2d_extract(Profile profile)
        {
            Lines2d seLines = null;
            XElement lines2dElement = null;

            try
            {
                try { seLines = (Lines2d)profile.Lines2d; }
                catch (Exception ex) { Console.WriteLine($"Lines2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seLines != null && seLines.Count > 0)
                {
                    lines2dElement = new XElement("Lines2d", new XAttribute("Count", seLines.Count));

                    for (int k = 1; k <= seLines.Count; k++)
                    {
                        Line2d seLine = null;
                        try
                        {
                            seLine = seLines.Item(k);
                            XElement lineElement = new XElement("Line");

                            try { lineElement.Add(new XElement("index", seLine.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("name", seLine.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("key", seLine.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("length", seLine.Length)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Length: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("keypointcount", seLine.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("segmentedStyleCount", seLine.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("angle", seLine.Angle)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Angle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("isProjection", seLine.Projection)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Projection: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("projectionDashName", seLine.ProjectionDashName)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] ProjectionDashName: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("isChamfer", seLine.IsChamfer)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] IsChamfer: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("type", seLine.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { lineElement.Add(new XElement("zOrder", seLine.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seLine.GetStartPoint(out double strX, out double strY);
                                lineElement.Add(new XElement("startPoint", new XAttribute("Sx", strX), new XAttribute("Sy", strY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] GetStartPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seLine.GetEndPoint(out double endX, out double endY);
                                lineElement.Add(new XElement("endPoint", new XAttribute("Ex", endX), new XAttribute("Ey", endY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] GetEndPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seLine.KeyPointCount; f++)
                                {
                                    seLine.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                lineElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = seLine.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Line[{k}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    lineElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"Line[{k}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            lines2dElement.Add(lineElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Line[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seLine != null)
                            {
                                Marshal.ReleaseComObject(seLine);
                                seLine = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lines2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seLines != null)
                {
                    Marshal.ReleaseComObject(seLines);
                    seLines = null;
                }
            }

            return lines2dElement;
        }

        //Circle2d Extract
        public static XElement Circle2d_extract(Profile profile)
        {
            Circles2d seCircles = null;
            XElement circles2dElement = null;

            try
            {
                try { seCircles = (Circles2d)profile.Circles2d; }
                catch (Exception ex) { Console.WriteLine($"Circles2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seCircles != null && seCircles.Count > 0)
                {
                    circles2dElement = new XElement("Circles2d", new XAttribute("Count", seCircles.Count));

                    for (int l = 1; l <= seCircles.Count; l++)
                    {
                        Circle2d seCircle = null;
                        try
                        {
                            seCircle = seCircles.Item(l);
                            XElement circleElement = new XElement("Circle");

                            try { circleElement.Add(new XElement("index", seCircle.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("name", seCircle.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("key", seCircle.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("diameter", seCircle.Diameter)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Diameter: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("circumference", seCircle.Circumference)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Circumference: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("area", seCircle.Area)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Area: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("length", seCircle.Length)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Length: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("radius", seCircle.Radius)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Radius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("keypointcount", seCircle.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("segmentedStyleCount", seCircle.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("showCurvatureComb", seCircle.ShowCurvatureComb)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] ShowCurvatureComb: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("type", seCircle.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { circleElement.Add(new XElement("zOrder", seCircle.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seCircle.GetCenterPoint(out double x, out double y);
                                circleElement.Add(new XElement("centerpoint", new XAttribute("Xcoord", x), new XAttribute("Ycoord", y)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] GetCenterPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seCircle.GetMajorAxis(out double majX, out double majY);
                                circleElement.Add(new XElement("majorAxis", new XAttribute("Mx", majX), new XAttribute("My", majY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] GetMajorAxis: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seCircle.KeyPointCount; f++)
                                {
                                    seCircle.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                circleElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = seCircle.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Circle[{l}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    circleElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"Circle[{l}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            circles2dElement.Add(circleElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Circle[{l}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seCircle != null)
                            {
                                Marshal.ReleaseComObject(seCircle);
                                seCircle = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Circles2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seCircles != null)
                {
                    Marshal.ReleaseComObject(seCircles);
                    seCircles = null;
                }
            }

            return circles2dElement;
        }

        //Arc2d
        public static XElement Arc2d_extract(Profile profile)
        {
            Arcs2d seArcs = null;
            XElement arcs2dElement = null;

            try
            {
                try { seArcs = (Arcs2d)profile.Arcs2d; }
                catch (Exception ex) { Console.WriteLine($"Arcs2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seArcs != null && seArcs.Count > 0)
                {
                    arcs2dElement = new XElement("Arcs2d", new XAttribute("Count", seArcs.Count));

                    for (int m = 1; m <= seArcs.Count; m++)
                    {
                        Arc2d seArc = null;
                        try
                        {
                            seArc = seArcs.Item(m);
                            XElement arcElement = new XElement("Arc");

                            try { arcElement.Add(new XElement("index", seArc.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("name", seArc.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("key", seArc.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("radius", seArc.Radius)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Radius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("length", seArc.Length)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Length: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("orientation", seArc.Orientation)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Orientation: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("sweepangle", seArc.SweepAngle)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] SweepAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("startangle", seArc.StartAngle)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] StartAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("keypointcount", seArc.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("segmentedStyleCount", seArc.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("showCurvatureComb", seArc.ShowCurvatureComb)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] ShowCurvatureComb: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("type", seArc.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("zorder", seArc.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { arcElement.Add(new XElement("isFillet", seArc.IsFillet)); }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] IsFillet: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seArc.GetCenterPoint(out double cenX, out double cenY);
                                arcElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX), new XAttribute("Cy", cenY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] GetCenterPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seArc.GetStartPoint(out double strX, out double strY);
                                arcElement.Add(new XElement("startPoint", new XAttribute("Sx", strX), new XAttribute("Sy", strY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] GetStartPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seArc.GetEndPoint(out double endX, out double endY);
                                arcElement.Add(new XElement("endPoint", new XAttribute("Ex", endX), new XAttribute("Ey", endY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] GetEndPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seArc.GetMajorAxis(out double majX, out double majY);
                                arcElement.Add(new XElement("majorAxis", new XAttribute("Mx", majX), new XAttribute("My", majY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] GetMajorAxis: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seArc.KeyPointCount; f++)
                                {
                                    seArc.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                arcElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = seArc.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Arc[{m}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    arcElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"Arc[{m}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            arcs2dElement.Add(arcElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Arc[{m}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seArc != null)
                            {
                                Marshal.ReleaseComObject(seArc);
                                seArc = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Arcs2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seArcs != null)
                {
                    Marshal.ReleaseComObject(seArcs);
                    seArcs = null;
                }
            }

            return arcs2dElement;
        }

        //Ellipse Extract
        public static XElement Ellipse2d_extract(Profile profile)
        {
            Ellipses2d seEllipses = null;
            XElement ellipses2dElement = null;

            try
            {
                try { seEllipses = (Ellipses2d)profile.Ellipses2d; }
                catch (Exception ex) { Console.WriteLine($"Ellipses2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seEllipses != null && seEllipses.Count > 0)
                {
                    ellipses2dElement = new XElement("Ellipses2d", new XAttribute("Count", seEllipses.Count));

                    for (int n = 1; n <= seEllipses.Count; n++)
                    {
                        Ellipse2d seEllipse = null;
                        try
                        {
                            seEllipse = seEllipses.Item(n);
                            XElement ellipseElement = new XElement("Ellipse");

                            try { ellipseElement.Add(new XElement("index", seEllipse.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("name", seEllipse.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("key", seEllipse.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("majorradius", seEllipse.MajorRadius)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] MajorRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("minorradius", seEllipse.MinorRadius)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] MinorRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("ratio", seEllipse.MinorMajorRatio)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] MinorMajorRatio: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("length", seEllipse.Length)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Length: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("circumference", seEllipse.Circumference)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Circumference: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("area", seEllipse.Area)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Area: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("orientation", seEllipse.Orientation)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Orientation: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("rotationangle", seEllipse.RotationAngle)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] RotationAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("keypointcount", seEllipse.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("segmentedStyleCount", seEllipse.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("showCurvatureComb", seEllipse.ShowCurvatureComb)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] ShowCurvatureComb: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("type", seEllipse.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipseElement.Add(new XElement("zOrder", seEllipse.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipse.GetCenterPoint(out double cenX, out double cenY);
                                ellipseElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX), new XAttribute("Cy", cenY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] GetCenterPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipse.GetMajorAxis(out double majX, out double majY);
                                ellipseElement.Add(new XElement("majorAxis", new XAttribute("MAx", majX), new XAttribute("MAy", majY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] GetMajorAxis: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipse.GetMinorAxis(out double minX, out double minY);
                                ellipseElement.Add(new XElement("minorAxis", new XAttribute("MIx", minX), new XAttribute("MIy", minY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] GetMinorAxis: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seEllipse.KeyPointCount; f++)
                                {
                                    seEllipse.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                ellipseElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = seEllipse.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Ellipse[{n}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    ellipseElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"Ellipse[{n}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            ellipses2dElement.Add(ellipseElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ellipse[{n}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seEllipse != null)
                            {
                                Marshal.ReleaseComObject(seEllipse);
                                seEllipse = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ellipse2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seEllipses != null)
                {
                    Marshal.ReleaseComObject(seEllipses);
                    seEllipses = null;
                }
            }

            return ellipses2dElement;
        }

        //Elliptical Arcs Extract
        public static XElement EllipticalArc2d_extract(Profile profile)
        {
            EllipticalArcs2d seEllipticalArcs = null;
            XElement ellipticalArc2dElement = null;

            try
            {
                try { seEllipticalArcs = (EllipticalArcs2d)profile.EllipticalArcs2d; }
                catch (Exception ex) { Console.WriteLine($"EllipticalArcs2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seEllipticalArcs != null && seEllipticalArcs.Count > 0)
                {
                    ellipticalArc2dElement = new XElement("EllipticalArcs2d", new XAttribute("Count", seEllipticalArcs.Count));

                    for (int n = 1; n <= seEllipticalArcs.Count; n++)
                    {
                        EllipticalArc2d seEllipticalArc = null;
                        try
                        {
                            seEllipticalArc = seEllipticalArcs.Item(n);
                            XElement ellipticalArcElement = new XElement("EllipticalArc");

                            try { ellipticalArcElement.Add(new XElement("index", seEllipticalArc.Index)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("name", seEllipticalArc.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("key", seEllipticalArc.Key)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("majorradius", seEllipticalArc.MajorRadius)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] MajorRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("minorradius", seEllipticalArc.MinorRadius)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] MinorRadius: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("length", seEllipticalArc.Length)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Length: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("sweepAngle", seEllipticalArc.SweepAngle)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] SweepAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("startAngle", seEllipticalArc.StartAngle)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] StartAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("orientation", seEllipticalArc.Orientation)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Orientation: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("rotationangle", seEllipticalArc.RotationAngle)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] RotationAngle: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("keypointcount", seEllipticalArc.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("segmentedStyleCount", seEllipticalArc.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("showCurvatureComb", seEllipticalArc.ShowCurvatureComb)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] ShowCurvatureComb: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("type", seEllipticalArc.Type)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { ellipticalArcElement.Add(new XElement("zOrder", seEllipticalArc.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipticalArc.GetCenterPoint(out double cenX, out double cenY);
                                ellipticalArcElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX), new XAttribute("Cy", cenY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] GetCenterPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipticalArc.GetStartPoint(out double strX, out double strY);
                                ellipticalArcElement.Add(new XElement("startPoint", new XAttribute("Sx", strX), new XAttribute("Sy", strY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] GetStartPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipticalArc.GetEndPoint(out double endX, out double endY);
                                ellipticalArcElement.Add(new XElement("endPoint", new XAttribute("Ex", endX), new XAttribute("Ey", endY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] GetEndPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipticalArc.GetMajorAxis(out double majX, out double majY);
                                ellipticalArcElement.Add(new XElement("majorAxis", new XAttribute("MAx", majX), new XAttribute("MAy", majY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] GetMajorAxis: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seEllipticalArc.GetMinorAxis(out double minX, out double minY);
                                ellipticalArcElement.Add(new XElement("minorAxis", new XAttribute("MIx", minX), new XAttribute("MIy", minY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] GetMinorAxis: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seEllipticalArc.KeyPointCount; f++)
                                {
                                    seEllipticalArc.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                ellipticalArcElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = seEllipticalArc.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"EllipticalArc[{n}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    ellipticalArcElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"EllipticalArc[{n}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            ellipticalArc2dElement.Add(ellipticalArcElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"EllipticalArc[{n}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seEllipticalArc != null)
                            {
                                Marshal.ReleaseComObject(seEllipticalArc);
                                seEllipticalArc = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EllipticalArcs2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seEllipticalArcs != null)
                {
                    Marshal.ReleaseComObject(seEllipticalArcs);
                    seEllipticalArcs = null;
                }
            }

            return ellipticalArc2dElement;
        }

        //B-Spline Curves Extract
        public static XElement BSplineCurve2d_extract(Profile profile)
        {
            BSplineCurves2d seBSplines = null;
            XElement bSplineCurve2dElement = null;

            try
            {
                try { seBSplines = (BSplineCurves2d)profile.BSplineCurves2d; }
                catch (Exception ex) { Console.WriteLine($"BSplineCurves2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seBSplines != null && seBSplines.Count > 0)
                {
                    bSplineCurve2dElement = new XElement("BSplineCurves2d", new XAttribute("Count", seBSplines.Count));

                    for (int k = 1; k <= seBSplines.Count; k++)
                    {
                        BSplineCurve2d seBSpline = null;
                        try
                        {
                            seBSpline = seBSplines.Item(k);
                            XElement bSplineElement = new XElement("BSplineCurve");

                            try { bSplineElement.Add(new XElement("index", seBSpline.Index)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("name", seBSpline.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("key", seBSpline.Key)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("length", seBSpline.Length)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Length: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("area", seBSpline.Area)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Area: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("order", seBSpline.Order)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Order: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("form", seBSpline.Form.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Form: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("scope", seBSpline.Scope.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Scope: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("keyPointCount", seBSpline.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("nodeCount", seBSpline.NodeCount)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] NodeCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("poleCount", seBSpline.PoleCount)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] PoleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("segmentedStyleCount", seBSpline.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("shapeEdit", seBSpline.ShapeEdit)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] ShapeEdit: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("isTangentiallyClosedCurve", seBSpline.IsTangentiallyClosedCurve)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] IsTangentiallyClosedCurve: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("flexible", seBSpline.Flexible)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Flexible: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("derived", seBSpline.Derived)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Derived: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("showCurvatureComb", seBSpline.ShowCurvatureComb)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] ShowCurvatureComb: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("showControlPolygon", seBSpline.ShowControlPolygon)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] ShowControlPolygon: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("type", seBSpline.Type)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { bSplineElement.Add(new XElement("zOrder", seBSpline.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seBSpline.GetCentroid(out double cenX, out double cenY);
                                bSplineElement.Add(new XElement("centroid", new XAttribute("Cx", cenX), new XAttribute("Cy", cenY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] GetCentroid: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seBSpline.GetParameterRange(out double startParam, out double endParam);
                                bSplineElement.Add(new XElement("parameterRange", new XAttribute("Start", startParam), new XAttribute("End", endParam)));
                            }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] GetParameterRange: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement nodesElement = new XElement("Nodes");
                                for (int a = 0; a < seBSpline.NodeCount; a++)
                                {
                                    seBSpline.GetNode(a, out double nodeX, out double nodeY);
                                    nodesElement.Add(new XElement("node", new XAttribute("Nindex", a), new XAttribute("Nx", nodeX), new XAttribute("Ny", nodeY)));
                                }
                                bSplineElement.Add(nodesElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] GetNode: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement polesElement = new XElement("Poles");
                                for (int b = 0; b < seBSpline.PoleCount; b++)
                                {
                                    seBSpline.GetPole(b, out double poleX, out double poleY);
                                    polesElement.Add(new XElement("pole", new XAttribute("Pindex", b), new XAttribute("Px", poleX), new XAttribute("Py", poleY)));
                                }
                                bSplineElement.Add(polesElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] GetPole: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seBSpline.KeyPointCount; f++)
                                {
                                    seBSpline.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                bSplineElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                object numberOfNodes, nodes, numberOfPoles, poles, knots, rational, weights, degree, form, scope;
                                seBSpline.GetData(out numberOfNodes, out nodes, out numberOfPoles, out poles,
                                                    out knots, out rational, out weights, out degree, out form, out scope);

                                Console.WriteLine($"BSplineCurve[{k}] GetData types - Knots:{knots?.GetType().Name}, Weights:{weights?.GetType().Name}, Rational:{rational?.GetType().Name}, Degree:{degree?.GetType().Name}");

                                XElement dataElement = new XElement("NurbsData");
                                dataElement.Add(new XElement("degree", degree?.ToString() ?? ""));
                                dataElement.Add(new XElement("rational", rational?.ToString() ?? ""));

                                dataElement.Add(BuildArrayOrScalarElement("Knots", knots));
                                dataElement.Add(BuildArrayOrScalarElement("Weights", weights));

                                bSplineElement.Add(dataElement);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"BSplineCurve[{k}] GetData: {ex.Message} | Inner: {ex.InnerException?.Message}");
                            }

                            try
                            {
                                Relationships2d relationships = seBSpline.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"BSplineCurve[{k}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    bSplineElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"BSplineCurve[{k}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            bSplineCurve2dElement.Add(bSplineElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"BSplineCurve[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seBSpline != null)
                            {
                                Marshal.ReleaseComObject(seBSpline);
                                seBSpline = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BSplineCurves2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seBSplines != null)
                {
                    Marshal.ReleaseComObject(seBSplines);
                    seBSplines = null;
                }
            }

            return bSplineCurve2dElement;
        }

        private static XElement BuildArrayOrScalarElement(string label, object value)
        {
            var element = new XElement(label);

            if (value == null)
            {
                element.Add(new XAttribute("Value", "null"));
                return element;
            }

            if (value is Array arr && arr.Rank == 1)
            {
                element.Add(new XAttribute("Count", arr.Length));
                for (int i = 0; i < arr.Length; i++)
                {
                    element.Add(new XElement("v", new XAttribute("i", i), arr.GetValue(i)));
                }
            }
            else
            {
                element.Add(new XAttribute("UnexpectedType", value.GetType().Name));
                element.Add(new XAttribute("RawValue", value.ToString()));
            }

            return element;
        }

        //Conics Extract
        public static XElement Conic2d_extract(Profile profile)
        {
            Conics2d seConics = null;
            XElement conic2dElements = null;

            try
            {
                try { seConics = (Conics2d)profile.Conics2d; }
                catch (Exception ex) { Console.WriteLine($"Conics2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seConics != null && seConics.Count > 0)
                {
                    conic2dElements = new XElement("Conics2d", new XAttribute("Count", seConics.Count));

                    for (int k = 1; k <= seConics.Count; k++)
                    {
                        Conic2d seConic = null;
                        try
                        {
                            seConic = seConics.Item(k);
                            XElement conicElement = new XElement("Conic");

                            try { conicElement.Add(new XElement("index", seConic.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { conicElement.Add(new XElement("name", seConic.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { conicElement.Add(new XElement("key", seConic.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { conicElement.Add(new XElement("rhoValue", seConic.RhoValue)); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] RhoValue: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { conicElement.Add(new XElement("segmentedStyleCount", seConic.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { conicElement.Add(new XElement("type", seConic.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { conicElement.Add(new XElement("zOrder", seConic.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seConic.GetControlPoint(out double conX, out double conY);
                                conicElement.Add(new XElement("controlPoint", new XAttribute("CPx", conX), new XAttribute("CPy", conY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] GetControlPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seConic.GetStartPoint(out double strX, out double strY);
                                conicElement.Add(new XElement("startPoint", new XAttribute("Sx", strX), new XAttribute("Sy", strY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] GetStartPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seConic.GetEndPoint(out double endX, out double endY);
                                conicElement.Add(new XElement("endPoint", new XAttribute("Ex", endX), new XAttribute("Ey", endY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] GetEndPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = seConic.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Conic[{k}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    conicElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"Conic[{k}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            conic2dElements.Add(conicElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Conic[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seConic != null)
                            {
                                Marshal.ReleaseComObject(seConic);
                                seConic = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conics2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seConics != null)
                {
                    Marshal.ReleaseComObject(seConics);
                    seConics = null;
                }
            }

            return conic2dElements;
        }

        //Points2d Extract
        public static XElement Point2d_extract(Profile profile)
        {
            Points2d sePoints = null;
            XElement point2dElements = null;

            try
            {
                try { sePoints = (Points2d)profile.Points2d; }
                catch (Exception ex) { Console.WriteLine($"Points2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (sePoints != null && sePoints.Count > 0)
                {
                    point2dElements = new XElement("Points2d", new XAttribute("Count", sePoints.Count));

                    for (int k = 1; k <= sePoints.Count; k++)
                    {
                        Point2d sePoint = null;
                        try
                        {
                            sePoint = sePoints.Item(k);
                            XElement pointElement = new XElement("Point");

                            try { pointElement.Add(new XElement("index", sePoint.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("name", sePoint.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("key", sePoint.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("x-Coordinate", sePoint.x)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] X-Coordinate: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("y-Coordinate", sePoint.y)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] Y-Coordinate: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("keyPointCount", sePoint.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("segmentedStyleCount", sePoint.SegmentedStyleCount)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] SegmentedStyleCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("type", sePoint.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { pointElement.Add(new XElement("zOrder", sePoint.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < sePoint.KeyPointCount; f++)
                                {
                                    sePoint.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                pointElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                sePoint.GetPoint(out double pX, out double pY);
                                pointElement.Add(new XElement("controlPoint", new XAttribute("Px", pX), new XAttribute("Py", pY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] GetPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                Relationships2d relationships = sePoint.Relationships;
                                if (relationships != null && relationships.Count > 0)
                                {
                                    XElement relationshipsElement = new XElement("Relationships", new XAttribute("Count", relationships.Count));
                                    for (int r = 1; r <= relationships.Count; r++)
                                    {
                                        object relObj = null;
                                        try
                                        {
                                            relObj = relationships.Item(r);
                                            if (relObj is Relation2d rel)
                                            {
                                                relationshipsElement.Add(new XElement("Relationship",
                                                    new XAttribute("Index", rel.Index),
                                                    new XAttribute("Type", ((ObjectType)rel.Type).ToString())));
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Point[{k}] Relationship[{r}]: {ex.Message} | Inner: {ex.InnerException?.Message}");
                                        }
                                        finally
                                        {
                                            if (relObj != null)
                                            {
                                                Marshal.ReleaseComObject(relObj);
                                                relObj = null;
                                            }
                                        }
                                    }
                                    pointElement.Add(relationshipsElement);
                                }
                                if (relationships != null) Marshal.ReleaseComObject(relationships);
                            }
                            catch (Exception ex) { Console.WriteLine($"Point[{k}] Relationships: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            point2dElements.Add(pointElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Point[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (sePoint != null)
                            {
                                Marshal.ReleaseComObject(sePoint);
                                sePoint = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Points2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (sePoints != null)
                {
                    Marshal.ReleaseComObject(sePoints);
                    sePoints = null;
                }
            }

            return point2dElements;
        }


        //Hole2d Extract
        public static XElement Hole2d_extract(Profile profile)
        {
            Holes2d seHoles2d = null;
            XElement hole2dElements = null;

            try
            {
                try { seHoles2d = (Holes2d)profile.Holes2d; }
                catch (Exception ex) { Console.WriteLine($"Holes2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seHoles2d != null && seHoles2d.Count > 0)
                {
                    hole2dElements = new XElement("Holes2d", new XAttribute("Count", seHoles2d.Count));

                    for (int k = 1; k <= seHoles2d.Count; k++)
                    {
                        Hole2d seHole2d = null;
                        try
                        {
                            seHole2d = seHoles2d.Item(k);
                            XElement holeElement = new XElement("Hole");

                            try { holeElement.Add(new XElement("index", seHole2d.index)); }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { holeElement.Add(new XElement("name", seHole2d.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { holeElement.Add(new XElement("key", seHole2d.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { holeElement.Add(new XElement("type", seHole2d.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { holeElement.Add(new XElement("zOrder", seHole2d.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { holeElement.Add(new XElement("keyPointCount", seHole2d.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                XElement keyPointsElement = new XElement("KeyPoints");
                                for (int f = 0; f < seHole2d.KeyPointCount; f++)
                                {
                                    seHole2d.GetKeyPoint(f, out double keyX, out double keyY, out double keyZ,
                                                        out KeyPointType keyPointType, out int handleType);
                                    keyPointsElement.Add(new XElement("keyPoint",
                                        new XAttribute("Kindex", f), new XAttribute("Kx", keyX), new XAttribute("Ky", keyY),
                                        new XAttribute("Kz", keyZ), new XAttribute("Ktype", keyPointType), new XAttribute("Htype", handleType)));
                                }
                                holeElement.Add(keyPointsElement);
                            }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] GetKeyPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seHole2d.GetCenterPoint(out double cenX, out double cenY);
                                holeElement.Add(new XElement("centerPoint", new XAttribute("Cx", cenX), new XAttribute("Cy", cenY)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] GetCenterPoint: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                seHole2d.Range(out double xMin, out double yMin, out double xMax, out double yMax);
                                holeElement.Add(new XElement("Range",
                                    new XAttribute("XMin", xMin), new XAttribute("YMin", yMin),
                                    new XAttribute("XMax", xMax), new XAttribute("YMax", yMax)));
                            }
                            catch (Exception ex) { Console.WriteLine($"Hole[{k}] GetRange: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            hole2dElements.Add(holeElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Hole[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seHole2d != null)
                            {
                                Marshal.ReleaseComObject(seHole2d);
                                seHole2d = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Holes2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seHoles2d != null)
                {
                    Marshal.ReleaseComObject(seHoles2d);
                    seHoles2d = null;
                }
            }

            return hole2dElements;
        }

        //Boundaries2d Extract
        public static XElement Boundaries2d_extract(Profile profile)
        {
            Boundaries2d seBoundaries = null;
            XElement boundaries2dElement = null;

            try
            {
                try { seBoundaries = (Boundaries2d)profile.Boundaries2d; }
                catch (Exception ex) { Console.WriteLine($"Boundaries2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (seBoundaries != null && seBoundaries.Count > 0)
                {
                    boundaries2dElement = new XElement("Boundaries2d", new XAttribute("Count", seBoundaries.Count));

                    for (int k = 1; k <= seBoundaries.Count; k++)
                    {
                        Boundary2d seBoundary = null;
                        try
                        {
                            seBoundary = seBoundaries.Item(k);
                            XElement boundaryElement = new XElement("Boundary");

                            try { boundaryElement.Add(new XElement("index", seBoundary.Index)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] Index: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("name", seBoundary.Name?.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] Name: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("key", seBoundary.Key)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] Key: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("type", seBoundary.Type)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("zOrder", seBoundary.ZOrder)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] ZOrder: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("perimeter", seBoundary.Perimeter)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] Perimeter: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("area", seBoundary.Area)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] Area: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("state", seBoundary.State.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] State: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { boundaryElement.Add(new XElement("keyPointCount", seBoundary.KeyPointCount)); }
                            catch (Exception ex) { Console.WriteLine($"Boundary[{k}] KeyPointCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try
                            {
                                BoundingObjects2d boundingObjects = seBoundary.BoundingObjects;
                                boundaryElement.Add(new XElement("BoundingObjects", new XAttribute("Count", boundingObjects.Count)));
                                Marshal.ReleaseComObject(boundingObjects);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Boundary[{k}] BoundingObjects: {ex.Message} | Inner: {ex.InnerException?.Message}");
                            }

                            boundaries2dElement.Add(boundaryElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Boundary[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (seBoundary != null)
                            {
                                Marshal.ReleaseComObject(seBoundary);
                                seBoundary = null;
                            }
                        }
                    }
                    Console.WriteLine("Created Boundaries2d XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Boundaries2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (seBoundaries != null)
                {
                    Marshal.ReleaseComObject(seBoundaries);
                    seBoundaries = null;
                }
            }

            return boundaries2dElement;
        }

        //RectangularPatterns2d Extract
        public static XElement RectangularPatterns2d_extract(Profile profile)
        {
            RectangularPatterns2d sePatterns = null;
            XElement patterns2dElement = null;

            try
            {
                try { sePatterns = (RectangularPatterns2d)profile.RectangularPatterns2d; }
                catch (Exception ex) { Console.WriteLine($"RectangularPatterns2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (sePatterns != null && sePatterns.Count > 0)
                {
                    patterns2dElement = new XElement("RectangularPatterns2d", new XAttribute("Count", sePatterns.Count));

                    for (int k = 1; k <= sePatterns.Count; k++)
                    {
                        RectangularPattern2d sePattern = null;
                        try
                        {
                            sePattern = sePatterns.Item(k);
                            XElement patternElement = new XElement("RectangularPattern");

                            try { patternElement.Add(new XElement("type", sePattern.Type)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("layer", sePattern.Layer)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] Layer: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("xCount", sePattern.XCount)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] XCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("yCount", sePattern.YCount)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] YCount: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("xSpace", sePattern.XSpace)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] XSpace: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("ySpace", sePattern.YSpace)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] YSpace: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("width", sePattern.Width)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] Width: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("height", sePattern.Height)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] Height: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("offsetType", sePattern.OffsetType.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] OffsetType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("staggerType", sePattern.StaggerType.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] StaggerType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("staggerOffset", sePattern.StaggerOffset)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] StaggerOffset: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("staggerOffsetHalf", sePattern.StaggerOffsetHalf)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] StaggerOffsetHalf: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("staggerIncludeLast", sePattern.StaggerIncludeLast)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] StaggerIncludeLast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("referenceOccurrence", sePattern.ReferenceOccurrence)); }
                            catch (Exception ex) { Console.WriteLine($"RectangularPattern[{k}] ReferenceOccurrence: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            patterns2dElement.Add(patternElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"RectangularPattern[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (sePattern != null)
                            {
                                Marshal.ReleaseComObject(sePattern);
                                sePattern = null;
                            }
                        }
                    }
                    Console.WriteLine("Created RectangularPatterns2d XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"RectangularPatterns2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (sePatterns != null)
                {
                    Marshal.ReleaseComObject(sePatterns);
                    sePatterns = null;
                }
            }

            return patterns2dElement;
        }

        //CircularPatterns2d Extract
        public static XElement CircularPatterns2d_extract(Profile profile)
        {
            CircularPatterns2d sePatterns = null;
            XElement patterns2dElement = null;

            try
            {
                try { sePatterns = (CircularPatterns2d)profile.CircularPatterns2d; }
                catch (Exception ex) { Console.WriteLine($"CircularPatterns2d Cast: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                if (sePatterns != null && sePatterns.Count > 0)
                {
                    patterns2dElement = new XElement("CircularPatterns2d", new XAttribute("Count", sePatterns.Count));

                    for (int k = 1; k <= sePatterns.Count; k++)
                    {
                        CircularPattern2d sePattern = null;
                        try
                        {
                            sePattern = sePatterns.Item(k);
                            XElement patternElement = new XElement("CircularPattern");

                            try { patternElement.Add(new XElement("type", sePattern.Type)); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] Type: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("layer", sePattern.Layer)); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] Layer: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("geometryType", sePattern.GeometryType.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] GeometryType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("offsetType", sePattern.OffsetType.ToString())); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] OffsetType: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("count", sePattern.Count)); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] Count: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("angularSpacing", sePattern.AngularSpacing)); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] AngularSpacing: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            try { patternElement.Add(new XElement("referenceOccurrence", sePattern.ReferenceOccurrence)); }
                            catch (Exception ex) { Console.WriteLine($"CircularPattern[{k}] ReferenceOccurrence: {ex.Message} | Inner: {ex.InnerException?.Message}"); }

                            patterns2dElement.Add(patternElement);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"CircularPattern[{k}] Item: {ex.Message} | Inner: {ex.InnerException?.Message}");
                        }
                        finally
                        {
                            if (sePattern != null)
                            {
                                Marshal.ReleaseComObject(sePattern);
                                sePattern = null;
                            }
                        }
                    }
                    Console.WriteLine("Created CircularPatterns2d XML list");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"CircularPatterns2d: Error Message:{ex.Message} | Inner: {ex.InnerException?.Message}");
            }
            finally
            {
                if (sePatterns != null)
                {
                    Marshal.ReleaseComObject(sePatterns);
                    sePatterns = null;
                }
            }

            return patterns2dElement;
        }
    }
}
