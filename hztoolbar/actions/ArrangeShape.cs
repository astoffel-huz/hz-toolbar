#nullable enable

using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace hztoolbar.actions {

	public abstract class AbstractArrangeShapeAction : ToolbarAction {
		protected AbstractArrangeShapeAction(string id) : base(id) { }

		public override bool IsEnabled(string arg = "") {
			return GetSelectedShapes().Take(2).Count() == 2;
		}
	}


	public abstract class AbstractControllableArrangeShapeAction : AbstractArrangeShapeAction {

		private static ImmutableDictionary<string, int> GUTTERS = new Dictionary<string, int> {
			["none"] = 0,
			["small"] = 4,
			["medium"] = 10,
			["large"] = 20
		}.ToImmutableDictionary();

		public class ShapeSnapshot {
			public PowerPoint.Shape Shape;
			public readonly float Top;
			public readonly float Left;
			public readonly float Width;
			public readonly float Height;

			public ShapeSnapshot(PowerPoint.Shape shape) {
				this.Shape = shape;
				this.Top = shape.Top;
				this.Left = shape.Left;
				this.Width = shape.Width;
				this.Height = shape.Height;
			}
		}

		protected AbstractControllableArrangeShapeAction(string id) : base(id) { }

		protected abstract IEnumerable<PowerPoint.Shape> DoGetShapes();

		protected abstract void DoRun(ImmutableList<PowerPoint.Shape> shapes,
			int horizontal_gutter, int vertical_gutter, bool horizontal_resize, bool vertical_resize);

		public override bool Run(string arg = "") {
			var shapes = DoGetShapes().ToImmutableList();
			var horizontal_gutter = Properties.Settings.Default.arrange_horizontal_gutter;
			var vertical_gutter = Properties.Settings.Default.arrange_vertical_gutter;
			var horizontal_resize = Properties.Settings.Default.arrange_grid_resize_horizontal;
			var vertical_resize = Properties.Settings.Default.arrange_grid_resize_vertical;
			if (arg == "") {
				DoRun(shapes, horizontal_gutter, vertical_gutter, horizontal_resize, vertical_resize);
			} else if (arg == "interactive") {
				DoRun(shapes, horizontal_gutter, vertical_gutter, horizontal_resize, vertical_resize);
				var window = Utils.CreateModalWindow();
				var control = new ArrangeShapeInteractiveControl(window) {
					HorizontalGutter = horizontal_gutter,
					VerticalGutter = vertical_gutter,
					HorizontalResize = horizontal_resize,
					VerticalResize = vertical_resize
				};
				EventHandler changeHandler = (s, e) => {
					DoRun(shapes, control.HorizontalGutter, control.VerticalGutter, control.HorizontalResize, control.VerticalResize);
				};
				control.ValueChanged += changeHandler;
				try {
					window.Content = control;
					if (window.ShowDialog() == true) {
						Properties.Settings.Default.arrange_horizontal_gutter = control.HorizontalGutter;
						Properties.Settings.Default.arrange_vertical_gutter = control.VerticalGutter;
						Properties.Settings.Default.arrange_grid_resize_horizontal = control.HorizontalResize;
						Properties.Settings.Default.arrange_grid_resize_vertical = control.VerticalResize;
						Properties.Settings.Default.Save();
					}
				} finally {
					control.ValueChanged -= changeHandler;
				}
			} else if (GUTTERS.TryGetValue(arg, out int gutter)) {
				DoRun(shapes, gutter, gutter, false, false);
			}
			return false;
		}
	}

	public class ArrangeHorizontalAction : AbstractArrangeShapeAction {

		public ArrangeHorizontalAction() : base("arrange_horizontal") { }

		public override bool Run(string arg = "") {
			var shapes = (
				from shape in GetSelectedShapes()
				orderby shape.Left
				select shape
				).ToList();
			if (shapes.Count > 1) {
				var top = shapes[0].Top;
				var left = shapes[0].Left;
				foreach (var shape in shapes) {
					shape.Top = top;
					shape.Left = left;
					left += shape.Width + Properties.Settings.Default.arrange_horizontal_gutter;
				}
			}
			return false;
		}
	}

	public class ArrangeVerticalAction : AbstractArrangeShapeAction {

		public ArrangeVerticalAction() : base("arrange_vertical") { }

		public override bool Run(string arg = "") {
			var shapes = (
				from shape in GetSelectedShapes()
				orderby shape.Top
				select shape
				).ToList();
			if (shapes.Count > 1) {
				var reference = shapes[0];
				var top = shapes[0].Top;
				var left = shapes[0].Left;
				foreach (var shape in shapes) {
					shape.Top = top;
					shape.Left = left;
					top += shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
				}
			}
			return false;
		}

	}

	public class ArrangeGridAction : AbstractControllableArrangeShapeAction {

		public ArrangeGridAction() : base("arrange_grid") { }

		private List<List<PowerPoint.Shape>> MakeRows(ImmutableList<PowerPoint.Shape> shapes) {
			shapes = (
				from shape in shapes
				orderby shape.Top + 1.0 / 8.0 * shape.Height
				select shape
				).ToImmutableList();
			var result = new List<List<PowerPoint.Shape>>();
			var currentRow = new List<PowerPoint.Shape>();
			var scanline = shapes[0].Top + 7.0 / 8.0 * shapes[0].Height;
			foreach (var shape in shapes) {
				if (shape.Top + 1.0 / 8.0 * shape.Height > scanline) {
					result.Add(currentRow);
					currentRow = new List<PowerPoint.Shape>();
				}
				currentRow.Add(shape);
				scanline = Math.Max(scanline, shape.Top + 7.0 / 8.0 * shape.Height);
			}
			if (currentRow.Count > 0) {
				result.Add(currentRow);
			}
			return result;
		}

		private List<List<PowerPoint.Shape>> MakeColumns(ImmutableList<PowerPoint.Shape> shapes) {
			shapes = (
				from shape in shapes
				orderby shape.Left + 1.0 / 8.0 * shape.Width
				select shape
				).ToImmutableList();
			var result = new List<List<PowerPoint.Shape>>();
			var currentColumn = new List<PowerPoint.Shape>();
			var scanline = shapes[0].Left + 7.0 / 8.0 * shapes[0].Width;
			foreach (var shape in shapes) {
				if (shape.Left + 1.0 / 8.0 * shape.Width > scanline) {
					result.Add(currentColumn);
					currentColumn = new List<PowerPoint.Shape>();
				}
				currentColumn.Add(shape);
				scanline = Math.Max(scanline, shape.Left + 7.0 / 8.0 * shape.Width);
			}
			if (currentColumn.Count > 0) {
				result.Add(currentColumn);
			}
			return result;
		}

		private void AlignRows(ImmutableList<PowerPoint.Shape> shapes, int gutter, bool vertical_resize) {
			var rows = MakeRows(shapes);
			var top = rows[0].Min(it => it.Top);
			foreach (var row in rows) {
				var max_height = 0f;
				var next_top = top;
				foreach (var shape in row) {
					shape.Top = top;
					next_top = Math.Max(next_top, shape.Top + shape.Height);
					max_height = Math.Max(max_height, shape.Height);
				}
				if (vertical_resize) {
					foreach (var shape in row) {
						shape.Height = max_height;
					}
				}
				top = next_top + gutter;
			}
		}

		private void AlignColumns(ImmutableList<PowerPoint.Shape> shapes, int gutter, bool horizontal_resize) {
			var columns = MakeColumns(shapes);
			var left = columns[0].Min(it => it.Left);
			foreach (var col in columns) {
				var max_width = 0f;
				var next_left = left;
				foreach (var shape in col) {
					shape.Left = left;
					next_left = Math.Max(next_left, shape.Left + shape.Width);
					max_width = Math.Max(max_width, shape.Width);
				}
				if (horizontal_resize) {
					foreach (var shape in col) {
						shape.Width = max_width;
					}
				}
				left = next_left + gutter;
			}
		}

		protected override IEnumerable<PowerPoint.Shape> DoGetShapes() {
			return GetSelectedShapes();
		}

		protected override void DoRun(ImmutableList<PowerPoint.Shape> shapes, int horizontal_gutter, int vertical_gutter, bool horizontal_resize, bool vertical_resize) {
			AlignRows(shapes, vertical_gutter, vertical_resize);
			AlignColumns(shapes, horizontal_gutter, horizontal_resize);
		}

	}

	public class ShapeMagnet : AbstractArrangeShapeAction {

		public ShapeMagnet() : base("arrange_magnet") { }

		private (
			List<PowerPoint.Shape> right, List<PowerPoint.Shape> left,
			List<PowerPoint.Shape> top, List<PowerPoint.Shape> bottom
			) Partition(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes) {
			var right = new List<PowerPoint.Shape>();
			var bottom = new List<PowerPoint.Shape>();
			var left = new List<PowerPoint.Shape>();
			var top = new List<PowerPoint.Shape>();

			var ref_cx = reference.Left + 0.5 * reference.Width;
			var ref_cy = reference.Top + 0.5 * reference.Height;
			var norm_x = reference.Width > 0 ? 1.0 / (0.5 * reference.Width) : 1.0;
			var norm_y = reference.Height > 0 ? 1.0 / (0.5 * reference.Height) : 1.0;
			foreach (var shape in shapes) {
				if (shape == reference) {
					continue;
				}
				var cx = shape.Left + 0.5 * shape.Width;
				var cy = shape.Top + 0.5 * shape.Height;
				var sx = norm_x * (cx - ref_cx);
				var sy = norm_y * (cy - ref_cy);

				if (0.0 <= sx) {
					if (sy > sx) {
						bottom.Add(shape);
					} else if (Math.Abs(sy) <= sx) {
						right.Add(shape);
					} else {
						top.Add(shape);
					}
				} else {
					if (sy > -sx) {
						bottom.Add(shape);
					} else if (Math.Abs(sy) <= -sx) {
						left.Add(shape);
					} else {
						top.Add(shape);
					}
				}
			}
			return (right, left, top, bottom);
		}

		private void RightMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes) {
			shapes.Sort((a, b) => a.Top < b.Top ? -1 : a.Top == b.Top ? 0 : 1);
			var left = reference.Left + reference.Width + Properties.Settings.Default.arrange_horizontal_gutter;
			var top = reference.Top;
			foreach (var shape in shapes) {
				shape.Top = top;
				shape.Left = left;
				top += shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
			}
		}

		private void LeftMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes) {
			shapes.Sort((a, b) => a.Top < b.Top ? -1 : a.Top == b.Top ? 0 : 1);
			var right = reference.Left - Properties.Settings.Default.arrange_horizontal_gutter;
			var top = reference.Top;
			foreach (var shape in shapes) {
				shape.Top = top;
				shape.Left = right - shape.Width;
				top += shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
			}
		}

		private void TopMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes) {
			shapes.Sort((a, b) => a.Left < b.Left ? -1 : a.Left == b.Left ? 0 : 1);
			var bottom = reference.Top - Properties.Settings.Default.arrange_vertical_gutter;
			var left = reference.Left;
			foreach (var shape in shapes) {
				shape.Top = bottom - shape.Height;
				shape.Left = left;
				left += shape.Width + Properties.Settings.Default.arrange_horizontal_gutter;
			}
		}
		private void BottomMagnet(PowerPoint.Shape reference, List<PowerPoint.Shape> shapes) {
			shapes.Sort((a, b) => a.Left < b.Left ? -1 : a.Left == b.Left ? 0 : 1);
			var top = reference.Top + reference.Height + Properties.Settings.Default.arrange_vertical_gutter;
			var left = reference.Left;
			foreach (var shape in shapes) {
				shape.Top = top;
				shape.Left = left;
				left += shape.Width + Properties.Settings.Default.arrange_horizontal_gutter;
			}
		}
		public override bool Run(string arg = "") {
			var shapes = GetSelectedShapes().ToList();
			if (shapes.Count > 1) {
				var reference = shapes[0];
				var partitions = Partition(reference, shapes);
				RightMagnet(reference, partitions.right);
				BottomMagnet(reference, partitions.bottom);
				LeftMagnet(reference, partitions.left);
				TopMagnet(reference, partitions.top);
			}
			return false;
		}
	}

}