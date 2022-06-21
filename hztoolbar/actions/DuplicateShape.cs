#nullable enable


using System.Collections.Generic;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace hztoolbar.actions {

	public class DuplicateShapeAction : ToolbarAction {

		public const string TO_LEFT = "left";
		public const string TO_RIGHT = "right";
		public const string TO_TOP = "top";
		public const string TO_BOTTOM = "bottom";

		public DuplicateShapeAction() : base("duplicate_shape") { }

		protected override IEnumerable<PowerPoint.Shape> GetSelectedShapes() {
			return from shape in base.GetSelectedShapes()
				   where shape.Type == Office.MsoShapeType.msoAutoShape
				   select shape;
		}

		public override bool IsEnabled(string arg = "") {
			return GetSelectedShapes().Take(2).Count() == 1;
		}

		public override bool Run(string arg = "") {
			var shape = GetSelectedShapes().FirstOrDefault();
			var slide = Utils.GetActiveSlide();

			if (shape != null && slide != null) {
				var top = arg switch {
					TO_BOTTOM => shape.Top + shape.Height + Properties.Settings.Default.arrange_vertical_gutter,
					TO_TOP => shape.Top - shape.Height - Properties.Settings.Default.arrange_vertical_gutter,
					_ => shape.Top
				};
				var left = arg switch {
					TO_LEFT => shape.Left - shape.Width - Properties.Settings.Default.arrange_horizontal_gutter,
					TO_RIGHT => shape.Left + shape.Width + Properties.Settings.Default.arrange_horizontal_gutter,
					_ => shape.Left
				};

				var copy = slide.Shapes.AddShape(shape.AutoShapeType, left, top, shape.Width, shape.Height);
				shape.PickUp();
				copy.Apply();
				if (shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue) {
					shape.TextFrame.TextRange.Copy();
					copy.TextFrame.TextRange.Paste();
				}
			}
			return false;
		}
	}

}