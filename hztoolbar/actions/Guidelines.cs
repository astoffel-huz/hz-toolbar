#nullable enable

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using System.Collections.Generic;

namespace hztoolbar.actions
{


    public abstract class AbstractGuideAction : ToolbarAction
    {

        protected AbstractGuideAction(string id) : base(id) { }

        protected IEnumerable<PowerPoint.Guide> EnumerateGuides(PowerPoint.PpGuideOrientation? orientation)
        {
            var slide = Utils.GetActiveSlide();
            if (slide == null)
            {
                return Enumerable.Empty<PowerPoint.Guide>();
            }

            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActiveWindow.Presentation;
            var presentationGuides = presentation.Guides;
            var masterGuides = slide.Master.Guides;
            if (presentationGuides == null || masterGuides == null)
            {
                return Enumerable.Empty<PowerPoint.Guide>();
            }
            return from PowerPoint.Guide guide in Enumerable.Concat<PowerPoint.Guide>(
                from PowerPoint.Guide guide in presentationGuides select guide,
                from PowerPoint.Guide guide in masterGuides select guide
                )
                   where guide.Orientation == orientation
                   select guide;
        }

    }

    public class AlignGuideAction : AbstractGuideAction
    {

        private const string TOP = "top";
        private const string LEFT = "left";
        private const string RIGHT = "right";
        private const string BOTTOM = "bottom";

        public AlignGuideAction() : base("align_guide") { }

        private PowerPoint.PpGuideOrientation? GetOrientation(string arg)
        {
            return arg switch
            {
                TOP => PowerPoint.PpGuideOrientation.ppHorizontalGuide,
                BOTTOM => PowerPoint.PpGuideOrientation.ppHorizontalGuide,
                LEFT => PowerPoint.PpGuideOrientation.ppVerticalGuide,
                RIGHT => PowerPoint.PpGuideOrientation.ppVerticalGuide,
                _ => null
            };
        }

        public override bool IsEnabled(string arg = "")
        {
            var shapes = GetSelectedShapes();
            var guides = EnumerateGuides(GetOrientation(arg));
            return
                shapes.Take(1).Count() > 0
                && guides.Take(1).Count() > 0;
        }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            var guides = (
                from guide in EnumerateGuides(GetOrientation(arg))
                orderby guide.Position
                select guide
            ).ToList();

            if (shapes.Count == 0 || guides.Count == 0)
            {
                return false;
            }

            if (arg == TOP)
            {
                var minTop = shapes.Min(it => it.Top);
                var guide = Utils.FloorItem(minTop, guides, it => it.Position);
                if (guide == null)
                {
                    guide = guides[0];
                }
                foreach (var shape in shapes)
                {
                    shape.Top = guide.Position;
                }
            }
            else if (arg == BOTTOM)
            {
                var maxBottom = shapes.Max(it => it.Top + it.Height);
                var guide = Utils.CeilingItem(maxBottom, guides, it => it.Position);
                if (guide == null)
                {
                    guide = guides[guides.Count - 1];
                }
                foreach (var shape in shapes)
                {
                    shape.Top = guide.Position - shape.Height;
                }
            }
            else if (arg == LEFT)
            {
                var minLeft = shapes.Min(it => it.Left);
                var guide = Utils.FloorItem(minLeft, guides, it => it.Position);
                if (guide == null)
                {
                    guide = guides[0];
                }
                foreach (var shape in shapes)
                {
                    shape.Left = guide.Position;
                }
            }
            else if (arg == RIGHT)
            {
                var maxRight = shapes.Min(it => it.Left + it.Width);
                var guide = Utils.CeilingItem(maxRight, guides, it => it.Position);
                if (guide == null)
                {
                    guide = guides[guides.Count - 1];
                }
                foreach (var shape in shapes)
                {
                    shape.Left = guide.Position - shape.Width;
                }
            }
            return false;
        }
    }

    public class ResizeToGuideAction : AbstractGuideAction
    {
        public const string HORIZONTAL = "horizontal";
        public const string VERTICAL = "vertical";

        public ResizeToGuideAction() : base("resize_guide") { }

        private PowerPoint.PpGuideOrientation? GetOrientation(string arg)
        {
            return arg switch
            {
                HORIZONTAL => PowerPoint.PpGuideOrientation.ppVerticalGuide,
                VERTICAL => PowerPoint.PpGuideOrientation.ppHorizontalGuide,
                _ => null
            };
        }

        private (PowerPoint.Guide First, PowerPoint.Guide Last)? FindGuides(string arg, List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count < 1)
            {
                return null;
            }
            var guides = (
                from guide in EnumerateGuides(GetOrientation(arg))
                orderby guide.Position
                select guide
               ).ToList();
            if (guides.Count < 2)
            {
                return null;
            }
            (float min, float max)? extrema = arg switch
            {
                HORIZONTAL => (shapes.Min(it => it.Left), shapes.Max(it => it.Left + it.Width)),
                VERTICAL => (shapes.Min(it => it.Top), shapes.Max(it => it.Top + it.Height)),
                _ => null
            };
            if (extrema == null) { return null; }
            var result = (
                Utils.FloorItem(extrema.Value.min, guides, it => it.Position),
                Utils.CeilingItem(extrema.Value.max, guides, it => it.Position)
                );
            if (result.Item1 == null)
            {
                result.Item1 = guides[0];
            }
            if (result.Item2 == null)
            {
                result.Item2 = guides[guides.Count - 1];
            }
            if (result.Item1.Position == result.Item2.Position)
            {
                return null;
            }
            return (result.Item1, result.Item2);
        }

        public override bool IsEnabled(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            var guides = FindGuides(arg, shapes);
            return guides != null;
        }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            var guides = FindGuides(arg, shapes);

            if (guides != null)
            {
                var low = guides.Value.First.Position;
                var size = guides.Value.Last.Position - low;
                if (arg == HORIZONTAL)
                {
                    foreach (var shape in shapes)
                    {
                        shape.Left = low;
                        shape.Width = size;
                    }
                }
                else if (arg == VERTICAL)
                {
                    foreach (var shape in shapes)
                    {
                        shape.Top = low;
                        shape.Height = size;
                    }
                }
            }
            return false;
        }
    }

}