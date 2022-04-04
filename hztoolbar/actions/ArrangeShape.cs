#nullable enable

using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Linq;


namespace hztoolbar.actions
{

    public abstract class AbstractArrangeShapeAction : ToolbarAction
    {
        protected AbstractArrangeShapeAction(string id) : base(id) { }

        public override bool IsEnabled(string arg = "")
        {
            return GetSelectedShapes().Take(2).Count() == 2;
        }
    }


    public class ArrangeHorizontalAction : AbstractArrangeShapeAction
    {

        public ArrangeHorizontalAction() : base("arrange_horizontal") { }



        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                shapes.Sort((a, b) => a.Left < b.Left ? -1 : a.Left == b.Left ? 0 : 1);
                var left = reference.Left + reference.Width + Properties.Settings.Default.arrange_horizontal_gutter;
                foreach (var shape in shapes)
                {
                    if (shape == reference)
                    {
                        continue;
                    }
                    shape.Top = reference.Top;
                    shape.Left = left;
                    left = shape.Left + shape.Width + Properties.Settings.Default.arrange_horizontal_gutter;
                }
            }
        }

    }

    public class ArrangeVerticalAction : AbstractArrangeShapeAction
    {

        public ArrangeVerticalAction() : base("arrange_vertical") { }

        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                shapes.Sort((a, b) => a.Top < b.Top ? -1 : a.Top == b.Top ? 0 : 1);
                var top = reference.Top + reference.Height + Properties.Settings.Default.arrange_vertical_gutter;
                foreach (var shape in shapes)
                {
                    if (shape == reference)
                    {
                        continue;
                    }
                    shape.Left = reference.Left;
                    shape.Top = top;
                    top = shape.Top + shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
                }
            }
        }

    }

    public class ArrangeGridAction : AbstractArrangeShapeAction
    {

        public ArrangeGridAction() : base("arrange_grid") { }

        private List<List<PowerPoint.Shape>> MakeRows(List<PowerPoint.Shape> shapes)
        {
            shapes = new List<PowerPoint.Shape>(shapes);
            shapes.Sort((a, b) => a.Top.CompareTo(b.Top));
            var result = new List<List<PowerPoint.Shape>>();
            var currentRow = new List<PowerPoint.Shape>();
            var scanline = shapes[0].Top + shapes[0].Height;
            foreach (var s in shapes)
            {
                if (s.Top > scanline)
                {
                    result.Add(currentRow);
                    currentRow = new List<PowerPoint.Shape>();
                }
                currentRow.Add(s);
                scanline = Math.Max(scanline, s.Top + s.Height);
            }
            if (currentRow.Count > 0)
            {
                result.Add(currentRow);
            }
            return result;
        }

        private List<List<PowerPoint.Shape>> MakeColumns(List<PowerPoint.Shape> shapes)
        {
            shapes = new List<PowerPoint.Shape>(shapes);
            shapes.Sort((a, b) => a.Left.CompareTo(b.Left));
            var result = new List<List<PowerPoint.Shape>>();
            var currentColumn = new List<PowerPoint.Shape>();
            var scanline = shapes[0].Left + shapes[0].Width;
            foreach (var s in shapes)
            {
                if (s.Left > scanline)
                {
                    result.Add(currentColumn);
                    currentColumn = new List<PowerPoint.Shape>();
                }
                currentColumn.Add(s);
                scanline = Math.Max(scanline, s.Left + s.Width);
            }
            if (currentColumn.Count > 0)
            {
                result.Add(currentColumn);
            }
            return result;
        }

        private void AlignRows(List<PowerPoint.Shape> shapes)
        {
            var rows = MakeRows(shapes);
            var top = rows[0].Min(it => it.Top);
            foreach (var row in rows)
            {
                var next_top = top;
                foreach (var shape in row)
                {
                    shape.Top = top;
                    next_top = Math.Max(next_top, shape.Top + shape.Height + Properties.Settings.Default.arrange_vertical_gutter);
                }
                top = next_top;
            }
        }

        private void AlignColumns(List<PowerPoint.Shape> shapes)
        {
            var columns = MakeColumns(shapes);
            var left = columns[0].Min(it => it.Left);
            foreach (var col in columns)
            {
                var next_left = left;
                foreach (var shape in col)
                {
                    shape.Left = left;
                    next_left = Math.Max(next_left, shape.Left + shape.Width + Properties.Settings.Default.arrange_horizontal_gutter);
                }
                left = next_left;
            }
        }

        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                AlignRows(shapes);
                AlignColumns(shapes);
            }
        }
    }

    public class ShapeMagnet : AbstractArrangeShapeAction
    {

        public ShapeMagnet() : base("arrange_magnet") { }

        private (
            List<PowerPoint.Shape> right, List<PowerPoint.Shape> left,
            List<PowerPoint.Shape> top, List<PowerPoint.Shape> bottom
            ) Partition(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes)
        {
            var right = new List<PowerPoint.Shape>();
            var bottom = new List<PowerPoint.Shape>();
            var left = new List<PowerPoint.Shape>();
            var top = new List<PowerPoint.Shape>();

            var refLeft = reference.Left;
            var refRight = reference.Left + reference.Width;
            var refTop = reference.Top;
            var refBottom = reference.Top + reference.Height;
            foreach (var shape in shapes)
            {
                var cx = shape.Left + 0.5 * shape.Width;
                var cy = shape.Top + 0.5 * shape.Height;
                if (refRight < cx)
                {
                    right.Add(shape);
                }
                else if (cx < refLeft)
                {
                    left.Add(shape);
                }
                else if (refBottom < cy)
                {
                    bottom.Add(shape);
                }
                else if (cy < refTop)
                {
                    top.Add(shape);
                }
            }
            return (right, left, top, bottom);
        }

        private void RightMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes)
        {
            shapes.Sort((a, b) => a.Top < b.Top ? -1 : a.Top == b.Top ? 0 : 1);
            var left = reference.Left + reference.Width + Properties.Settings.Default.arrange_horizontal_gutter;
            var top = reference.Top;
            foreach (var shape in shapes)
            {
                shape.Top = top;
                shape.Left = left;
                top += shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
            }
        }

        private void LeftMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes)
        {
            shapes.Sort((a, b) => a.Top < b.Top ? -1 : a.Top == b.Top ? 0 : 1);
            var right = reference.Left + Properties.Settings.Default.arrange_horizontal_gutter;
            var top = reference.Top;
            foreach (var shape in shapes)
            {
                shape.Top = top;
                shape.Left = right - shape.Width;
                top += shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
            }
        }

        private void TopMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes)
        {
            shapes.Sort((a, b) => a.Left < b.Left ? -1 : a.Left == b.Left ? 0 : 1);
            var bottom = reference.Top - Properties.Settings.Default.arrange_vertical_gutter;
            var left = reference.Left;
            foreach (var shape in shapes)
            {
                shape.Top = bottom - shape.Height;
                shape.Left = left;
                left += shape.Width + Properties.Settings.Default.arrange_horizontal_gutter;
            }
        }
        private void BottomMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes)
        {
            shapes.Sort((a, b) => a.Left < b.Left ? -1 : a.Left == b.Left ? 0 : 1);
            var top = reference.Top + reference.Height + Properties.Settings.Default.arrange_vertical_gutter;
            var left = reference.Left;
            foreach (var shape in shapes)
            {
                shape.Top = top;
                shape.Left = left;
                left += shape.Width + Properties.Settings.Default.arrange_horizontal_gutter;
            }
        }
        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                var partitions = Partition(reference, shapes);
                RightMagnet(reference, partitions.right);
                BottomMagnet(reference, partitions.bottom);
                LeftMagnet(reference, partitions.left);
                TopMagnet(reference, partitions.top);
            }
        }
    }

}