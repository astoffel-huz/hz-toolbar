#nullable enable

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using System.Collections.Generic;

namespace hztoolbar.actions
{

    public class AlignGuideAction : ToolbarAction
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

        private IEnumerable<PowerPoint.Guide> EnumerateGuides(PowerPoint.PpGuideOrientation? orientation)
        {
            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActiveWindow.Presentation;
            return from PowerPoint.Guide guide in presentation.Guides
                   where guide.Orientation == orientation
                   select guide;
        }

        public override bool IsEnabled(string arg = "")
        {
            var shapes = GetSelectedShapes();
            var guides = EnumerateGuides(GetOrientation(arg));
            return
                shapes.Take(1).Count() > 0
                && guides.Take(1).Count() > 0;
        }

        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            var guides = (
                from guide in EnumerateGuides(GetOrientation(arg))
                orderby guide.Position
                select guide
            ).ToList();

            if (shapes.Count == 0 || guides.Count == 0)
            {
                return;
            }

            if (arg == TOP)
            {
                var minTop = shapes.Min(it => it.Top);
                var guide = guides.FindLast(it => it.Position < minTop);
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
                var guide = guides.Find(it => it.Position > maxBottom);
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
                var guide = guides.FindLast(it => it.Position < minLeft);
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
                var guide = guides.Find(it => it.Position > maxRight);
                if (guide == null)
                {
                    guide = guides[guides.Count - 1];
                }
                foreach (var shape in shapes)
                {
                    shape.Left = guide.Position - shape.Width;
                }
            }
        }
    }

    public class ResizeToGuideAction : ToolbarAction
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

        private IEnumerable<PowerPoint.Guide> EnumerateGuides(PowerPoint.PpGuideOrientation? orientation)
        {
            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActiveWindow.Presentation;
            var slide = Utils.GetActiveSlide();
            if (slide == null)
            {
                return Enumerable.Empty<PowerPoint.Guide>();
            }
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


        public override bool IsEnabled(string arg = "")
        {
            var guides = EnumerateGuides(GetOrientation(arg)).ToList();
            var shapes = GetSelectedShapes().ToList();

            if (guides.Count < 2 || shapes.Count == 0)
            {
                return false;
            }

            (float min, float max)? extrema = arg switch
            {
                HORIZONTAL => (shapes.Min(it => it.Left), shapes.Max(it => it.Left + it.Width)),
                VERTICAL => (shapes.Min(it => it.Top), shapes.Max(it => it.Top + it.Height)),
                _ => null
            };

            if (extrema == null)
            {
                return false;
            }

            return guides.Any(it => it.Position <= extrema.Value.min) && guides.Any(it => extrema.Value.max <= it.Position);
        }

        public override void Run(string arg = "")
        {
            var guides = EnumerateGuides(GetOrientation(arg)).ToList();
            var shapes = GetSelectedShapes().ToList();

            (float min, float max)? extrema = arg switch
            {
                HORIZONTAL => (shapes.Min(it => it.Left), shapes.Max(it => it.Left + it.Width)),
                VERTICAL => (shapes.Min(it => it.Top), shapes.Max(it => it.Top + it.Height)),
                _ => null
            };

            if (extrema != null)
            {
                guides.Sort((a, b) => a.Position < b.Position ? -1 : a.Position == b.Position ? 0 : 1);
                var lowGuide = guides.FindLast(it => it.Position <= extrema.Value.min);
                var highGuide = guides.Find(it => extrema.Value.max <= it.Position);
                if (lowGuide != null && highGuide != null)
                {
                    var low = lowGuide.Position;
                    var size = highGuide.Position - low;
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
            }

        }
    }

}