#nullable enable

using hztoolbar.Properties;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Office = Microsoft.Office.Core;

namespace hztoolbar.actions {

	public abstract class AbstractTextAction : ToolbarAction {
		protected AbstractTextAction(string id) : base(id) {
		}

		protected override IEnumerable<Shape> GetSelectedShapes() {
			return (
				from shape in base.GetSelectedShapes()
				where shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue
				select shape
			   );
		}

		public override bool IsEnabled(string arg = "") {
			return GetSelectedShapes().Take(1).Count() > 0;
		}
	}

	public class ShapeSwapText : AbstractTextAction {
		public ShapeSwapText() : base("swap_text") { }

		public override bool IsEnabled(string arg = "") {
			return GetSelectedShapes().Take(2).Count() >= 2;
		}

		public override bool Run(string arg = "") {
			var textShapes = GetSelectedShapes().ToList();
			if (textShapes.Count > 1) {
				var snapshot = Utils.Capture(textShapes[textShapes.Count - 1].TextFrame2);
				for (var i = textShapes.Count - 1; i > 0; --i) {
					Utils.Copy(textShapes[i].TextFrame2, textShapes[i - 1].TextFrame2);
				}
				Utils.Apply(textShapes[0].TextFrame2, snapshot);
			}
			return false;
		}
	}

	public class ClearTextAction : AbstractTextAction {
		public ClearTextAction() : base("clear_text") { }

		public override bool Run(string arg = "") {
			foreach (var shape in GetSelectedShapes()) {
				var fontCapture = Utils.Capture(shape.TextFrame2.TextRange.Font);
				shape.TextFrame2.DeleteText();
				Utils.Apply(shape.TextFrame2.TextRange.Font, fontCapture);
			}
			return false;
		}
	}

	public class SplitText : ToolbarAction {
		public SplitText() : base("split_text") { }

		private Selection? GetSelection() {

			var activeWindow = Utils.GetActiveWindow();
			if (activeWindow == null) {
				return null;
			}

			var selection = activeWindow.Selection;

			return selection.Type == PpSelectionType.ppSelectionText ? selection : null;
		}

		public override bool IsEnabled(string arg = "") {
			return GetSelection() != null;
		}

		public override bool Run(string arg = "") {
			var selection = GetSelection();
			if (selection != null) {
				foreach (Shape shape in selection.ShapeRange) {
					var start = selection.TextRange2.Start;
					var shapeRange = shape.TextFrame2.TextRange;
					if (start < shapeRange.Length) {
						var duplicate = shape.Duplicate();
						var duplicateRange = duplicate.TextFrame2.TextRange;
						shapeRange.Characters[start, shapeRange.Length - start + 1].Delete();
						duplicateRange.Characters[1, start - 1].Delete();
						shape.Height = shape.Height / 2.0f - Properties.Settings.Default.arrange_vertical_gutter / 2.0f;
						duplicate.Left = shape.Left;
						duplicate.Top = shape.Top + shape.Height + Properties.Settings.Default.arrange_vertical_gutter;
						duplicate.Height = shape.Height;
					}
				}
			}
			return true;
		}
	}

	public class MergeTextAction : AbstractTextAction {

		public MergeTextAction() : base("merge_text") { }

		public override bool Run(string arg = "") {
			var selection = GetSelectedShapes().ToList();
			if (selection.Count > 0) {
				var pivot = selection[0];
				var pivotTextRange = pivot.TextFrame2.TextRange;
				foreach (var shape in selection.Skip(1)) {
					if (pivotTextRange.Text.Last() != '\n') {
						pivotTextRange.Text += "\n";
					}
					var shapeTextRange = shape.TextFrame2.TextRange;
					var offset = pivotTextRange.Length;
					pivotTextRange.InsertAfter(shapeTextRange.Text);
					Utils.TransferCharacters(pivotTextRange, shapeTextRange, offset);
					shape.Delete();
				}
			}
			return true;
		}
	}

	public class ChangeLanguageAction : AbstractTextAction {
		private readonly ImmutableDictionary<string, Office.MsoLanguageID> LANGUAGES = new Dictionary<string, Office.MsoLanguageID>() {
			["de"] = Office.MsoLanguageID.msoLanguageIDGerman,
			["de-DE"] = Office.MsoLanguageID.msoLanguageIDGerman,

			["en"] = Office.MsoLanguageID.msoLanguageIDEnglishUS,
			["en-UK"] = Office.MsoLanguageID.msoLanguageIDEnglishUK,
			["en-US"] = Office.MsoLanguageID.msoLanguageIDEnglishUS,

		}.ToImmutableDictionary();

		public ChangeLanguageAction() : base("change_language") { }

		public override bool IsEnabled(string arg = "") {
			return LANGUAGES.ContainsKey(arg) && (
				GetSelectedShapes().Take(1).Count() > 0
				|| Utils.GetActiveSlide() != null
			   );
		}

		public override bool Run(string arg = "") {
			if (LANGUAGES.TryGetValue(arg, out var language)) {
				var shapes = GetSelectedShapes().ToList();
				if (shapes.Count > 0) {
					foreach (var shape in shapes) {
						shape.TextFrame2.TextRange.LanguageID = language;
					}
				} else {
					var slide = Utils.GetActiveSlide();
					if (slide != null) {
						foreach (Shape shape in slide.Shapes) {
							if (shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue) {
								shape.TextFrame2.TextRange.LanguageID = language;
							}
						}
					}
				}
			}
			return false;
		}
	}

	public abstract class AbstractChangeTextMargin : AbstractTextAction {
		public AbstractChangeTextMargin(string id) : base(id) { }

		protected void ChangeMargins(IEnumerable<Shape> shapes, float top, float left, float bottom, float right) {
			foreach (var shape in shapes) {
				shape.TextFrame.MarginTop = top;
				shape.TextFrame.MarginLeft = left;
				shape.TextFrame.MarginBottom = bottom;
				shape.TextFrame.MarginRight = right;
			}
		}

	}

	public class DefaultTextMarginAction : AbstractChangeTextMargin {

		public DefaultTextMarginAction() : base("text_margin_default") { }

		public override bool Run(string arg = "") {
			ChangeMargins(
				GetSelectedShapes(),
				Properties.Settings.Default.textbox_margin_top,
				Properties.Settings.Default.textbox_margin_left,
				Properties.Settings.Default.textbox_margin_bottom,
				Properties.Settings.Default.textbox_margin_right
			);
			return false;
		}
	}

	public class ChangeTextMargin : AbstractChangeTextMargin {

		private readonly IImmutableSet<string> VALID_ARGUMENTS = new HashSet<string>() { "none", "small", "normal", "large" }.ToImmutableHashSet();

		private (float top, float left, float bottom, float right) GetMargin(string arg) {
			return arg switch {
				"small" => (
					Settings.Default.default_small_length, 2.0f * Settings.Default.default_small_length,
					Settings.Default.default_small_length, 2.0f * Settings.Default.default_small_length
				),
				"normal" => (
					Settings.Default.default_normal_length, 2.0f * Settings.Default.default_normal_length,
					Settings.Default.default_normal_length, 2.0f * Settings.Default.default_normal_length
				),
				"large" => (
					Settings.Default.default_large_length, 2.0f * Settings.Default.default_large_length,
					Settings.Default.default_large_length, 2.0f * Settings.Default.default_large_length
				),
				_ => (0.0f, 0.0f, 0.0f, 0.0f)
			};
		}

		public ChangeTextMargin() : base("text_margin") { }

		public override bool IsEnabled(string arg = "") {
			return VALID_ARGUMENTS.Contains(arg) && base.IsEnabled();
		}

		public override bool Run(string arg = "") {
			if (VALID_ARGUMENTS.Contains(arg)) {
				var margin = GetMargin(arg);
				ChangeMargins(GetSelectedShapes(), margin.top, margin.left, margin.bottom, margin.right);
			}
			return false;
		}
	}

	public class CustomTextMarginAction : AbstractChangeTextMargin {

		public CustomTextMarginAction() : base("text_margin_custom") { }

		public override bool Run(string arg = "") {
			var shapes = GetSelectedShapes().ToList();
			var top_margin = Properties.Settings.Default.textbox_margin_top;
			var left_margin = Properties.Settings.Default.textbox_margin_left;
			var bottom_margin = Properties.Settings.Default.textbox_margin_bottom;
			var right_margin = Properties.Settings.Default.textbox_margin_right;
			ChangeMargins(shapes, top_margin, left_margin, bottom_margin, right_margin);
			var window = Utils.CreateModalWindow();
			var control = new TextMarginControl(window) {
				TopMargin = top_margin,
				LeftMargin = left_margin,
				BottomMargin = bottom_margin,
				RightMargin = right_margin,
			};
			void changeHandler(object s, EventArgs e) {
				ChangeMargins(shapes, control.TopMargin, control.LeftMargin, control.BottomMargin, control.RightMargin);
			}
			control.ValueChanged += changeHandler;
			try {
				window.Content = control;
				if (window.ShowDialog() == true) {
					Properties.Settings.Default.textbox_margin_top = control.TopMargin;
					Properties.Settings.Default.textbox_margin_left = control.LeftMargin;
					Properties.Settings.Default.textbox_margin_bottom = control.BottomMargin;
					Properties.Settings.Default.textbox_margin_right = control.RightMargin;
					Properties.Settings.Default.Save();
				}
			} finally {
				control.ValueChanged -= changeHandler;
			}
			return false;
		}
	}
}

