using System.Collections.Generic;
using System.Collections.Immutable;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Linq;


namespace hztoolbar.actions
{

    public class HarmonizeAdjustmentsAction : ToolbarAction
    {
        public HarmonizeAdjustmentsAction() : base("harmonize_adjustments") { }

        protected override IEnumerable<PowerPoint.Shape> GetSelectedShapes()
        {
            return from shape in base.GetSelectedShapes()
                   where shape.Type == Office.MsoShapeType.msoAutoShape
                   select shape;
        }

        public override bool IsEnabled(string arg = "")
        {
            var adjustments = (
                from shape in GetSelectedShapes()
                select shape.Adjustments.Count
               ).ToList();
            return adjustments.Count > 1 && adjustments[0] > 0 && adjustments.All(it => it == adjustments[0]);
        }

        

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                if (reference.Adjustments.Count > 0
                    && !shapes.All(it => it.Adjustments.Count == reference.Adjustments.Count))
                {
                    return false;
                }
                foreach (var shape in shapes)
                {
                    if (shape == reference)
                    {
                        continue;
                    }
                    for (var i = 0; i < reference.Adjustments.Count; ++i)
                    {
                        shape.Adjustments[i + 1] = reference.Adjustments[i + 1];
                    }
                }
            }
            return false;
        }
    }




}
