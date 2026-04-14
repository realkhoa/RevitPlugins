using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MultiAlignedDIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private UIApplication m_uiApp;
        private Document m_doc;
        private Autodesk.Revit.DB.View m_activeView;
        private const double TOLERANCE = 0.0001;
        private const double DIM_OFFSET = 3.0;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            m_uiApp = commandData.Application;
            m_doc = m_uiApp.ActiveUIDocument.Document;
            m_activeView = m_doc.ActiveView;

            var selected = m_uiApp.ActiveUIDocument.Selection.GetElementIds()
                .Select(id => m_doc.GetElement(id)).ToList();

            if (selected.Count < 2)
            {
                message = "Chọn ít nhất 2 object";
                return Result.Failed;
            }

            var refs = new List<DimRefData>();

            foreach (var e in selected)
                refs.AddRange(ExtractReferences(e));

            var groups = GroupByDirection(refs);

            using (Transaction tx = new Transaction(m_doc, "Auto DIM"))
            {
                tx.Start();

                foreach (var g in groups)
                {
                    if (g.Value.Count < 2) continue;
                    CreateDimensionForGroup(g.Key, g.Value);
                }

                tx.Commit();
            }

            return Result.Succeeded;
        }

        // ================== EXTRACT ==================

        private List<DimRefData> ExtractReferences(Element elem)
        {
            var result = new List<DimRefData>();

            if (elem is Wall w)
                result.AddRange(ExtractWall(w));
            else if (elem is Grid g)
                result.AddRange(ExtractGrid(g));
            else if (elem is FamilyInstance fi)
                result.AddRange(ExtractFamily(fi));
            else
                result.AddRange(ExtractGeometry(elem));

            return result;
        }

        private List<DimRefData> ExtractWall(Wall wall)
        {
            var result = new List<DimRefData>();

            var faces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior)
                .Concat(HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior));

            foreach (var r in faces)
            {
                var geo = wall.GetGeometryObjectFromReference(r);
                if (geo is PlanarFace pf)
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = Flatten(pf.FaceNormal),
                        Point = pf.Origin
                    });
                }
            }

            return result;
        }

        private List<DimRefData> ExtractGrid(Grid grid)
        {
            var result = new List<DimRefData>();

            var opt = new Options { ComputeReferences = true };
            var geo = grid.get_Geometry(opt);

            foreach (var g in geo)
            {
                if (g is Line line && line.Reference != null)
                {
                    var dir = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
                    var normal = new XYZ(-dir.Y, dir.X, 0);

                    result.Add(new DimRefData
                    {
                        Ref = line.Reference,
                        Normal = normal,
                        Point = line.GetEndPoint(0)
                    });
                    break;
                }
            }

            return result;
        }

        // 🔥 MAIN UPGRADE HERE
        private List<DimRefData> ExtractFamily(FamilyInstance fi)
        {
            var result = new List<DimRefData>();

            var opt = new Options
            {
                ComputeReferences = true,
                View = m_activeView
            };

            var geo = fi.get_Geometry(opt);

            // 1. Faces
            foreach (var g in geo)
            {
                if (g is Solid s)
                {
                    foreach (Face f in s.Faces)
                    {
                        if (f is PlanarFace pf && pf.Reference != null)
                        {
                            result.Add(new DimRefData
                            {
                                Ref = pf.Reference,
                                Normal = Flatten(pf.FaceNormal),
                                Point = pf.Origin
                            });
                        }
                    }
                }
            }

            // 2. Family references (xịn nhất)
            try
            {
                foreach (Reference r in fi.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = new XYZ(1, 0, 0),
                        Point = fi.GetTransform().Origin
                    });
                }

                foreach (Reference r in fi.GetReferences(FamilyInstanceReferenceType.CenterFrontBack))
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = new XYZ(0, 1, 0),
                        Point = fi.GetTransform().Origin
                    });
                }
            }
            catch { }

            // 3. BoundingBox fallback
            var bb = fi.get_BoundingBox(m_activeView);
            if (bb != null)
            {
                var min = bb.Min;
                var max = bb.Max;

                foreach (Reference r in fi.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = new XYZ(1, 0, 0),
                        Point = fi.GetTransform().Origin
                    });
                }

                foreach (Reference r in fi.GetReferences(FamilyInstanceReferenceType.CenterFrontBack))
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = new XYZ(0, 1, 0),
                        Point = fi.GetTransform().Origin
                    });
                }
            }

            // 4. Center fallback (VALID API)
            if (fi.Location is LocationPoint lp)
            {
                var refsLR = fi.GetReferences(FamilyInstanceReferenceType.CenterLeftRight);
                foreach (var r in refsLR)
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = new XYZ(1, 0, 0),
                        Point = lp.Point
                    });
                }

                var refsFB = fi.GetReferences(FamilyInstanceReferenceType.CenterFrontBack);
                foreach (var r in refsFB)
                {
                    result.Add(new DimRefData
                    {
                        Ref = r,
                        Normal = new XYZ(0, 1, 0),
                        Point = lp.Point
                    });
                }
            }

            return result;
        }

        private List<DimRefData> ExtractGeometry(Element e)
        {
            var result = new List<DimRefData>();

            var opt = new Options { ComputeReferences = true };
            var geo = e.get_Geometry(opt);

            foreach (var g in geo)
            {
                if (g is Line line && line.Reference != null)
                {
                    var dir = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
                    var normal = new XYZ(-dir.Y, dir.X, 0);

                    result.Add(new DimRefData
                    {
                        Ref = line.Reference,
                        Normal = normal,
                        Point = line.GetEndPoint(0)
                    });
                }
            }

            return result;
        }

        // ================== DIM ==================

        private Dictionary<XYZ, List<DimRefData>> GroupByDirection(List<DimRefData> refs)
        {
            var dict = new Dictionary<XYZ, List<DimRefData>>(new XYZComparer());

            foreach (var r in refs)
            {
                var n = Canonical(r.Normal);

                if (!dict.ContainsKey(n))
                    dict[n] = new List<DimRefData>();

                dict[n].Add(r);
            }

            return dict;
        }

        private void CreateDimensionForGroup(XYZ normal, List<DimRefData> refs)
        {
            refs = refs.OrderBy(r => r.Point.DotProduct(normal)).ToList();

            var arr = new ReferenceArray();
            foreach (var r in refs)
                arr.Append(r.Ref);

            var p1 = refs.First().Point;
            var p2 = refs.Last().Point;

            var perp = new XYZ(-normal.Y, normal.X, 0).Normalize();

            var offset = perp * DIM_OFFSET;

            var line = Line.CreateBound(p1 + offset, p2 + offset);

            m_doc.Create.NewDimension(m_activeView, line, arr);
        }

        // ================== HELPERS ==================

        private XYZ Flatten(XYZ v)
        {
            return new XYZ(v.X, v.Y, 0).Normalize();
        }

        private XYZ Canonical(XYZ v)
        {
            if (v.X < 0) return v.Negate();
            return v;
        }

        private class XYZComparer : IEqualityComparer<XYZ>
        {
            public bool Equals(XYZ a, XYZ b)
            {
                return a.IsAlmostEqualTo(b);
            }

            public int GetHashCode(XYZ obj)
            {
                return (Math.Round(obj.X, 2), Math.Round(obj.Y, 2)).GetHashCode();
            }
        }

        private class DimRefData
        {
            public Reference Ref;
            public XYZ Normal;
            public XYZ Point;
        }
    }
}