﻿#nullable enable

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
				shape.TextFrame.DeleteText();
			}
			return false;
		}
	}

	public class ChangeLanguage : AbstractTextAction {
		private readonly ImmutableDictionary<string, Office.MsoLanguageID> LANGUAGES = new Dictionary<string, Office.MsoLanguageID>() {
			["de"] = Office.MsoLanguageID.msoLanguageIDGerman,
			["de-DE"] = Office.MsoLanguageID.msoLanguageIDGerman,

			["en"] = Office.MsoLanguageID.msoLanguageIDEnglishUS,
			["en-UK"] = Office.MsoLanguageID.msoLanguageIDEnglishUK,
			["en-US"] = Office.MsoLanguageID.msoLanguageIDEnglishUS,

		}.ToImmutableDictionary();

		public ChangeLanguage() : base("change_language") { }

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

	public class DefaultTextMargin : AbstractChangeTextMargin {

		public DefaultTextMargin() : base("text_margin_default") { }

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
		// H&Z Template 20223: x= 0.1"=7.2, y=0.05" = 3.6
		private readonly ImmutableDictionary<string, (float top, float left, float bottom, float right)> MARGINS = new Dictionary<string, (float top, float left, float bottom, float right)>() {
			["none"] = (0f, 0f, 0f, 0f),
			["small"] = (1.8f, 3.6f, 1.8f, 3.6f),
			["normal"] = (3.6f, 7.2f, 3.6f, 7.2f),
			["large"] = (7.2f, 14.4f, 7.2f, 14.4f),

		}.ToImmutableDictionary();

		public ChangeTextMargin() : base("text_margin") { }

		public override bool IsEnabled(string arg = "") {
			return MARGINS.ContainsKey(arg) && base.IsEnabled();
		}

		public override bool Run(string arg = "") {
			if (MARGINS.TryGetValue(arg, out var margin)) {
				ChangeMargins(GetSelectedShapes(), margin.top, margin.left, margin.bottom, margin.right);
			}
			return false;
		}
	}

	public class CustomTextMargin : AbstractChangeTextMargin {

		public CustomTextMargin() : base("text_margin_custom") { }

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
			EventHandler changeHandler = (s, e) => {
				ChangeMargins(shapes, control.TopMargin, control.LeftMargin, control.BottomMargin, control.RightMargin);
			};
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

