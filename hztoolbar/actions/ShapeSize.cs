using System.Collections.Generic;
using System.Linq;


namespace hztoolbar.actions
{

    public abstract class AbstractShapeSizeAction : ToolbarAction
    {
        protected AbstractShapeSizeAction(string id) : base(id) { }

        public override bool IsEnabled(string arg = "")
        {
            return GetSelectedShapes().Take(2).Count() == 2;
        }
    }
    
	public class HarmonizeShapeWidthAction : AbstractShapeSizeAction
    {

        public HarmonizeShapeWidthAction() : base("harmonize_shape_width") { }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                foreach (var shape in shapes)
                {
                    if (shape == reference)
                    {
                        continue;
                    }
                    shape.Width = reference.Width;
                }
            }
            return false;
        }

    }

    public class HarmonizeShapeHeightAction : AbstractShapeSizeAction
    {

        public HarmonizeShapeHeightAction() : base("harmonize_shape_height") { }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                foreach (var shape in shapes)
                {
                    if (shape == reference)
                    {
                        continue;
                    }
                    shape.Height = reference.Height;
                }
            }
            return false;
        }

    }
    
	public class HarmonizeShapeSizeAction : AbstractShapeSizeAction
    {

        public HarmonizeShapeSizeAction() : base("harmonize_shape_size") { }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1)
            {
                var reference = shapes[0];
                foreach (var shape in shapes)
                {
                    if (shape == reference)
                    {
                        continue;
                    }
                    shape.Width = reference.Width;
                    shape.Height = reference.Height;
                }
            }
            return false;
        }

    }

}
