
#nullable enable

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using System.Drawing;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.RightsManagement;
using System;

namespace hztoolbar.actions {

	public abstract class AbstractThemeColorAction : ToolbarAction {
		private const string DARK1 = "dark1";
		private const string LIGHT1 = "light1";
		private const string DARK2 = "dark2";
		private const string LIGHT2 = "light2";
		private const string ACCENT1 = "accent1";
		private const string ACCENT2 = "accent2";
		private const string ACCENT3 = "accent3";
		private const string ACCENT4 = "accent4";
		private const string ACCENT5 = "accent5";
		private const string ACCENT6 = "accent6";

		public AbstractThemeColorAction(string id) : base(id) { }

		protected override IEnumerable<PowerPoint.Shape> GetSelectedShapes() {
			return from shape in base.GetSelectedShapes()
				   where shape.Type == Office.MsoShapeType.msoAutoShape
					|| shape.Type == Office.MsoShapeType.msoTextBox
					|| shape.Type == Office.MsoShapeType.msoPlaceholder
				   select shape;
		}

		protected Color? GetColor(string arg) {
			var slide = Utils.GetActiveSlide();
			if (slide == null) {
				return null;
			}

			var themeColors = slide.ThemeColorScheme;
			if (themeColors == null) {
				return null;
			}

			Office.MsoThemeColorSchemeIndex? themeColorIndex = arg switch {
				DARK1 => Office.MsoThemeColorSchemeIndex.msoThemeDark1,
				LIGHT1 => Office.MsoThemeColorSchemeIndex.msoThemeLight1,
				DARK2 => Office.MsoThemeColorSchemeIndex.msoThemeDark2,
				LIGHT2 => Office.MsoThemeColorSchemeIndex.msoThemeLight2,
				ACCENT1 => Office.MsoThemeColorSchemeIndex.msoThemeAccent1,
				ACCENT2 => Office.MsoThemeColorSchemeIndex.msoThemeAccent2,
				ACCENT3 => Office.MsoThemeColorSchemeIndex.msoThemeAccent3,
				ACCENT4 => Office.MsoThemeColorSchemeIndex.msoThemeAccent4,
				ACCENT5 => Office.MsoThemeColorSchemeIndex.msoThemeAccent5,
				ACCENT6 => Office.MsoThemeColorSchemeIndex.msoThemeAccent6,
				_ => Office.MsoThemeColorSchemeIndex.msoThemeDark1
			};

			if (themeColorIndex == null) {
				return null;
			}

			var result = themeColors.Colors(themeColorIndex.Value).RGB;
			return ColorTranslator.FromOle(result);
		}

		public override bool IsEnabled(string arg = "") {
			var shapes = GetSelectedShapes();
			return shapes.Take(1).Count() > 0 && GetColor(arg) != null;
		}


	}


	public class ApplyBackgroundThemeColorAction : AbstractThemeColorAction {

		public ApplyBackgroundThemeColorAction() : base("apply_background_theme_color") { }

		private Color GetTextColor(Color color) {
			return color.GetBrightness() > Math.Sqrt(0.5) ? Color.Black : Color.White;
		}

		public override Bitmap? GetImage(string controlId, string arg = "") {
			var result = base.GetImage(this.Id);
			var color = GetColor(arg);
			if (result != null && color != null) {
				result = Utils.ReplaceBitmapColor(
					result,
					new Dictionary<Color, Color>() {
						[Color.Red] = GetTextColor(color.Value)
					},
					color.Value
				);
			}
			return result;
		}


		public override bool Run(string arg = "") {
			var shapes = GetSelectedShapes();
			var color = GetColor(arg);

			if (color != null) {
				foreach (var shape in shapes) {
					if (shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue) {
						// TODO: validate whether it is true that only shapes with text frames can be filled
						shape.Fill.Solid();
						shape.Fill.BackColor.RGB = ColorTranslator.ToOle(color.Value);
						shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(color.Value);
						shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(GetTextColor(color.Value));
					}

				}
			}
			return false;
		}
	}

	public class ApplyLineThemeColorAction : AbstractThemeColorAction {

		public ApplyLineThemeColorAction() : base("apply_line_theme_color") { }

		public override Bitmap? GetImage(string controlId, string arg = "") {
			var result = base.GetImage(this.Id);
			var color = GetColor(arg);
			if (result != null && color != null) {
				result = Utils.ReplaceBitmapColor(
					result,
					new Dictionary<Color, Color>() {
						[Color.Red] = Color.LightGray,
						[Color.FromArgb(255, 0, 255, 0)] = Color.DarkGray
					},
					color.Value
				);
			}
			return result;
		}

		public override bool Run(string arg = "") {
			var shapes = GetSelectedShapes();
			var color = GetColor(arg);

			if (color != null) {
				foreach (var shape in shapes) {
					if (shape.Line.Style == Office.MsoLineStyle.msoLineStyleMixed) {
						shape.Line.Style = Office.MsoLineStyle.msoLineSingle;
						shape.Line.Weight = 3.6f;
					}
					if (shape.Line.DashStyle == Office.MsoLineDashStyle.msoLineDashStyleMixed) {
						shape.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
						shape.Line.Weight = 3.6f;
					}
					shape.Line.BackColor.RGB = ColorTranslator.ToOle(color.Value);
					shape.Line.ForeColor.RGB = ColorTranslator.ToOle(color.Value);
				}
			}
			return false;
		}
	}
}