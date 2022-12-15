#nullable enable

using hztoolbar.Properties;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Interop;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace hztoolbar {
	using ColorReplaceImageCache = LinkedList<KeyValuePair<Tuple<Bitmap, ImmutableDictionary<int, int>, Color>, Bitmap>>;



	/// <summary>
	/// Class defining utility functions and tools.
	/// </summary>
	public static class Utils {

		public enum DefaultLength {
			Small, Normal, Large
		}

		public static bool IsTrue(Office.MsoTriState x) {
			return x == MsoTriState.msoTrue || x == MsoTriState.msoCTrue;
		}


		#region BulletFormat2
		public abstract class BulletFormatSnapshot {
			public abstract void Apply(Office.BulletFormat2 to);
		}

		public class UnnumberedBulletFormatSnapshot : BulletFormatSnapshot {
			public readonly int Character;
			public readonly float RelativeSize;
			public readonly FontSnapshot Font;
			public readonly Office.MsoTriState UseTextColor;
			public readonly Office.MsoTriState UseTextFont;
			public readonly Office.MsoTriState Visible;

			public UnnumberedBulletFormatSnapshot(int character, float relativeSize, FontSnapshot font, MsoTriState useTextColor, MsoTriState useTextFont, MsoTriState visible) {
				this.Character = character;
				this.RelativeSize = relativeSize;
				this.Font = font;
				this.UseTextColor = useTextColor;
				this.UseTextFont = useTextFont;
				this.UseTextColor = useTextColor;
				this.UseTextFont = useTextFont;
				this.Visible = visible;
			}

			public override void Apply(Office.BulletFormat2 to) {
				to.UseTextFont = this.UseTextFont;
				Utils.Apply(to.Font, this.Font);
				to.Character = this.Character;
				to.RelativeSize = this.RelativeSize;
				to.UseTextColor = this.UseTextColor;
				to.UseTextColor = this.UseTextColor;
				to.Visible = this.Visible;
				Debug.Assert(to.Type == MsoBulletType.msoBulletUnnumbered);
			}
		}

		public class NumberedBulletFormatSnapshot : BulletFormatSnapshot {
			public readonly int StartValue;
			public readonly Office.MsoNumberedBulletStyle Style;
			public readonly float RelativeSize;
			public readonly FontSnapshot Font;
			public readonly Office.MsoTriState UseTextColor;
			public readonly Office.MsoTriState UseTextFont;
			public readonly Office.MsoTriState Visible;

			public NumberedBulletFormatSnapshot(int startValue, Office.MsoNumberedBulletStyle style, float relativeSize, FontSnapshot font, MsoTriState useTextColor, MsoTriState useTextFont, MsoTriState visible) {
				this.StartValue = startValue;
				this.Style = style;
				this.RelativeSize = relativeSize;
				this.Font = font;
				this.UseTextColor = useTextColor;
				this.UseTextFont = useTextFont;
				this.UseTextColor = useTextColor;
				this.UseTextFont = useTextFont;
				this.Visible = visible;
			}
			public override void Apply(Office.BulletFormat2 to) {
				to.UseTextFont = this.UseTextFont;
				Utils.Apply(to.Font, this.Font);
				to.Style = this.Style;
				to.StartValue = this.StartValue;
				to.RelativeSize = this.RelativeSize;
				to.UseTextColor = this.UseTextColor;
				to.Visible = this.Visible;
				Debug.Assert(to.Type == MsoBulletType.msoBulletNumbered);
			}
		}

		public class NoneBulletFormatSnapshot : BulletFormatSnapshot {
			public override void Apply(Office.BulletFormat2 to) {
				to.Type = MsoBulletType.msoBulletNone;
				Debug.Assert(to.Type == MsoBulletType.msoBulletNone);
			}
		}

		public static BulletFormatSnapshot? Capture(Office.BulletFormat2 format) {
			return format.Type switch {
				MsoBulletType.msoBulletUnnumbered => new UnnumberedBulletFormatSnapshot(
										format.Character,
										format.RelativeSize,
										Capture(format.Font),
										format.UseTextColor,
										format.UseTextFont,
										format.Visible
										),
				MsoBulletType.msoBulletNumbered => new NumberedBulletFormatSnapshot(
										format.StartValue,
										format.Style,
										format.RelativeSize,
										Capture(format.Font),
										format.UseTextColor,
										format.UseTextFont,
										format.Visible
										),
				MsoBulletType.msoBulletNone => new NoneBulletFormatSnapshot(),
				_ => null,
			};
		}

		public static void Apply(Office.BulletFormat2 to, BulletFormatSnapshot? snapshot) {
			if (snapshot != null) {
				snapshot.Apply(to);
			}
		}

		public static void Copy(Office.BulletFormat2 to, Office.BulletFormat2 from) {
			Apply(to, Capture(from));
		}

		#endregion

		#region TabStops2

		public struct TabStopSnapshot {
			public readonly float Position;
			public readonly Office.MsoTabStopType Type;

			public TabStopSnapshot(float position, Office.MsoTabStopType type) {
				this.Position = position;
				this.Type = type;
			}
		}

		public class TabStopsSnapshot {
			public readonly ImmutableList<TabStopSnapshot> TabStops;

			public TabStopsSnapshot(IEnumerable<TabStopSnapshot> tabStops) {
				this.TabStops = tabStops.ToImmutableList();
			}
		}

		public static TabStopsSnapshot Capture(Office.TabStops2 tabStops) {
			return new TabStopsSnapshot(
				from Office.TabStop2 tabStop in tabStops
				select new TabStopSnapshot(tabStop.Position, tabStop.Type)
				);
		}

		public static void Clear(Office.TabStops2 tabStops) {
			List<Office.TabStop2> stops = new List<Office.TabStop2>();
			foreach (Office.TabStop2 tabStop in tabStops) {
				stops.Add(tabStop);
			}
			foreach (Office.TabStop2 tabStop in tabStops) {
				tabStop.Clear();
			}
		}

		public static void Apply(Office.TabStops2 to, TabStopsSnapshot snapshot) {
			Clear(to);
			foreach (var tabstop in snapshot.TabStops) {
				switch (tabstop.Type) {
					case Office.MsoTabStopType.msoTabStopLeft:
					case Office.MsoTabStopType.msoTabStopCenter:
					case Office.MsoTabStopType.msoTabStopRight:
					case Office.MsoTabStopType.msoTabStopDecimal:
						to.Add(tabstop.Type, tabstop.Position);
						break;
				};
			}
		}

		public static void Copy(Office.TabStops2 to, Office.TabStops2 from) {
			Apply(to, Capture(from));
		}
		#endregion

		#region ParagraphFormat2

		/// <summary>
		/// Snapshot of a paragraph format.
		/// </summary>
		public class ParagraphFormatSnapshot {
			public readonly Office.MsoParagraphAlignment Alignment;
			public readonly Office.MsoBaselineAlignment BaselineAlignment;
			public readonly BulletFormatSnapshot? Bullet;
			public readonly float FirstLineIndent;
			public readonly Office.MsoTriState HangingPunctuation;
			public readonly float LeftIndent;
			public readonly TabStopsSnapshot TabStops;

			public ParagraphFormatSnapshot(
				Office.MsoParagraphAlignment alignment, Office.MsoBaselineAlignment baselineAlignment,
				BulletFormatSnapshot? bullet, float firstLineIndent, Office.MsoTriState hangingPunctuation, float leftIndent,
				TabStopsSnapshot tabStops
			) {
				this.Alignment = alignment;
				this.BaselineAlignment = baselineAlignment;
				this.Bullet = bullet;
				this.FirstLineIndent = firstLineIndent;
				this.HangingPunctuation = hangingPunctuation;
				this.LeftIndent = leftIndent;
				this.TabStops = tabStops;
			}
		}

		/// <summary>
		/// Create snapshot of a pargaraph format.
		/// </summary>
		/// <param name="format">the format to capture</param>
		/// <returns>a snapshot of the paragraph format</returns>
		public static ParagraphFormatSnapshot Capture(Office.ParagraphFormat2 format) {
			return new ParagraphFormatSnapshot(
				format.Alignment,
				format.BaselineAlignment,
				Capture(format.Bullet),
				format.FirstLineIndent,
				format.HangingPunctuation,
				format.LeftIndent,
				Capture(format.TabStops)
				);
		}

		/// <summary>
		/// Copies a paragraph format snapshot onto a paragraph format.
		/// </summary>
		/// <param name="to">the paragraph format to modify</param>
		/// <param name="snapshot">the format snapshot to apply</param>
		public static void Apply(Office.ParagraphFormat2 to, ParagraphFormatSnapshot snapshot) {
			to.Alignment = snapshot.Alignment;
			to.BaselineAlignment = snapshot.BaselineAlignment;
			Apply(to.Bullet, snapshot.Bullet);
			try {
				to.FirstLineIndent = snapshot.FirstLineIndent;
			} catch { }
			to.HangingPunctuation = snapshot.HangingPunctuation;
			try {
				to.LeftIndent = snapshot.LeftIndent;
			} catch { }
			Apply(to.TabStops, snapshot.TabStops);
		}

		/// <summary>
		/// Copies a paragraph format.
		/// </summary>
		/// <param name="to">the target paragraph format</param>
		/// <param name="from">the source paragraph format to copy</param>
		public static void Copy(Office.ParagraphFormat2 to, Office.ParagraphFormat2 from) {
			Apply(to, Capture(from));
		}
		#endregion

		#region ColorFormat
		public abstract class ColorFormatSnapshot {
			public readonly float TintAndShade;

			public ColorFormatSnapshot(float tintAndShade) {
				this.TintAndShade = tintAndShade;
			}

			public abstract void Apply(Office.ColorFormat to);
		}

		public class RgbColorFormatSnapshot : ColorFormatSnapshot {
			public readonly int RGB;

			public RgbColorFormatSnapshot(int rgb, float tintAndShade) : base(tintAndShade) {
				this.RGB = rgb;
			}

			public override void Apply(ColorFormat to) {
				to.RGB = this.RGB;
				to.TintAndShade = this.TintAndShade;
				Debug.Assert(to.Type == MsoColorType.msoColorTypeRGB);
			}
		}

		public class SchemeColorFormatSnapshot : ColorFormatSnapshot {
			public readonly Office.MsoThemeColorIndex ObjectThemeColor;

			public SchemeColorFormatSnapshot( Office.MsoThemeColorIndex objectThemeColor, float tintAndShade) : base(tintAndShade) {
				this.ObjectThemeColor = objectThemeColor;
			}

			public override void Apply(Office.ColorFormat to) {
				to.ObjectThemeColor = this.ObjectThemeColor;
				to.TintAndShade = this.TintAndShade;
				Debug.Assert(to.Type == MsoColorType.msoColorTypeScheme);
			}
		}

		public static ColorFormatSnapshot? Capture(Office.ColorFormat from) {
			return from.Type switch {
				MsoColorType.msoColorTypeRGB => new RgbColorFormatSnapshot(from.RGB, from.TintAndShade),
				MsoColorType.msoColorTypeScheme => new SchemeColorFormatSnapshot(from.ObjectThemeColor, from.TintAndShade),
				_ => null,
			};
		}

		public static void Apply(Office.ColorFormat to, ColorFormatSnapshot? snapshot) {
			if (snapshot != null) {
				snapshot.Apply(to);
			}
		}

		public static void Copy(Office.ColorFormat to, Office.ColorFormat from) {
			Apply(to, Capture(from));
		}

		#endregion

		#region FillFormat

		public abstract class FillFormatSnapshot {

			public FillFormatSnapshot() { }

			public abstract void Apply(Office.FillFormat to);
		}

		public class SolidFillFormatSnapshot : FillFormatSnapshot {
			public readonly ColorFormatSnapshot? ForeColor;
			public readonly ColorFormatSnapshot? BackColor;

			public SolidFillFormatSnapshot(ColorFormatSnapshot? foreColor, ColorFormatSnapshot? backColor) {
				this.ForeColor = foreColor;
				this.BackColor = backColor;
			}

			public override void Apply(FillFormat to) {
				Utils.Apply(to.BackColor, this.BackColor);
				Utils.Apply(to.ForeColor, this.ForeColor);
				Debug.Assert(to.Type == Office.MsoFillType.msoFillSolid);
			}
		}

		public static FillFormatSnapshot? Capture(Office.FillFormat from) {
			return from.Type switch {
				Office.MsoFillType.msoFillSolid => new SolidFillFormatSnapshot(Capture(from.ForeColor), Capture(from.BackColor)),
				_ => null,
			};
		}

		public static void Apply(Office.FillFormat to, FillFormatSnapshot? snapshot) {
			if (snapshot != null) {
				snapshot.Apply(to);
			}
		}

		public static void Copy(Office.FillFormat to, Office.FillFormat from) {
			Apply(to, Capture(from));
		}

		#endregion

		#region Font2
		public class FontSnapshot {
			public readonly Office.MsoTriState Allcaps;
			public readonly float BaselineOffset;
			public readonly Office.MsoTriState Bold;
			public readonly Office.MsoTextCaps Caps;
			public readonly Office.MsoTriState DoubleStrikeThrough;
			public readonly Office.MsoTriState Equalize;
			public readonly FillFormatSnapshot? Fill;
			public readonly ColorFormatSnapshot? Highlight;
			public readonly Office.MsoTriState Italic;
			public readonly float Kerning;
			public readonly string Name;
			public readonly float Size;
			public readonly Office.MsoTriState Smallcaps;
			public readonly Office.MsoSoftEdgeType SoftEdgeFormat;
			public readonly float Spacing;
			public readonly Office.MsoTextStrike Strike;
			public readonly Office.MsoTriState StrikeThrough;
			public readonly Office.MsoTriState Subscript;
			public readonly Office.MsoTriState Superscript;
			public readonly Office.MsoTextUnderlineType UnderlineStyle;
			public readonly Office.MsoPresetTextEffect WordArtformat;
			// TODO highlight

			public FontSnapshot(
				string name, float size, Office.MsoTriState allcaps, float baselineOffset,
				Office.MsoTriState bold, Office.MsoTextCaps caps, Office.MsoTriState doubleStrikeThrough,
				Office.MsoTriState equalize, FillFormatSnapshot? fill, ColorFormatSnapshot? highlight,
				Office.MsoTriState italic, float kerning,
				Office.MsoTriState smallcaps, Office.MsoSoftEdgeType softEdgeFormat, float spacing,
				Office.MsoTextStrike strike, Office.MsoTriState strikeThrough, Office.MsoTriState subscript,
				Office.MsoTriState superscript, Office.MsoTextUnderlineType underlineStyle,
				Office.MsoPresetTextEffect wordArtformat
				) {
				this.Name = name;
				this.Size = size;
				this.Allcaps = allcaps;
				this.BaselineOffset = baselineOffset;
				this.Bold = bold;
				this.Caps = caps;
				this.DoubleStrikeThrough = doubleStrikeThrough;
				this.Equalize = equalize;
				this.Fill = fill;
				this.Highlight = highlight;
				this.Italic = italic;
				this.Kerning = kerning;
				this.Smallcaps = smallcaps;
				this.SoftEdgeFormat = softEdgeFormat;
				this.Spacing = spacing;
				this.Strike = strike;
				this.StrikeThrough = strikeThrough;
				this.Subscript = subscript;
				this.Superscript = superscript;
				this.UnderlineStyle = underlineStyle;
				this.WordArtformat = wordArtformat;
			}
		}

		public static FontSnapshot Capture(Office.Font2 font) {
		return new FontSnapshot(
				font.Name, font.Size, font.Allcaps, font.BaselineOffset,
				font.Bold, font.Caps, font.DoubleStrikeThrough, font.Equalize,
				Capture(font.Fill), Capture(font.Highlight), font.Italic, font.Kerning, font.Smallcaps,
				font.SoftEdgeFormat, font.Spacing, font.Strike, font.StrikeThrough,
				font.Subscript, font.Superscript, font.UnderlineStyle, font.WordArtformat
				);
		}

		public static void Apply(Office.Font2 to, FontSnapshot snapshot) {
			to.Name = snapshot.Name;
			to.Size = snapshot.Size;
			to.Allcaps = snapshot.Allcaps;
			to.BaselineOffset = snapshot.BaselineOffset;
			to.Bold = snapshot.Bold;
			to.Caps = snapshot.Caps;
			to.DoubleStrikeThrough = snapshot.DoubleStrikeThrough;
			to.Equalize = snapshot.Equalize;
			Apply(to.Fill, snapshot.Fill);
			Apply(to.Highlight, snapshot.Highlight);
			to.Italic = snapshot.Italic;
			to.Kerning = snapshot.Kerning;
			to.Smallcaps = snapshot.Smallcaps;
			to.SoftEdgeFormat = snapshot.SoftEdgeFormat;
			to.Spacing = snapshot.Spacing;
			to.Strike = snapshot.Strike;
			to.StrikeThrough = snapshot.StrikeThrough;
			to.Subscript = snapshot.Subscript;
			to.Superscript = snapshot.Superscript;
			to.UnderlineStyle = snapshot.UnderlineStyle;
			if (snapshot.WordArtformat != MsoPresetTextEffect.msoTextEffectMixed) {
				to.WordArtformat = snapshot.WordArtformat;
			}
		}

		public static void Copy(Office.Font2 to, Office.Font2 from) {
			Apply(to, Capture(from));
		}
		#endregion

		#region TextRange2
		public class CharacterRangeSnapshot {
			public readonly int Start;
			public readonly int Length;
			public readonly FontSnapshot Font;
			public readonly ImmutableList<CharacterRangeSnapshot> Runs;

			public CharacterRangeSnapshot(FontSnapshot font, IEnumerable<CharacterRangeSnapshot> runs) {
				this.Font = font;
				this.Runs = runs.ToImmutableList();
			}
		}

		public static CharacterRangeSnapshot CaptureCharacters(Office.TextRange2 range) {
			return new CharacterRangeSnapshot(
				Capture(range.Font),
				from Office.TextRange2 run in range
				where !(run.Start == range.Start && run.Length == range.Length)
				select CaptureCharacters(run)
				);
		}

		public static void ApplyCharacters(Office.TextRange2 to, CharacterRangeSnapshot snapshot) {
			Apply(to.Font, snapshot.Font);
			foreach (var run in snapshot.Runs) {
				ApplyCharacters(to.Characters[run.Start, run.Length], run);
			}
		}

		public static void CopyCharacters(Office.TextRange2 to, Office.TextRange2 from) {
			Copy(to.Font, from.Font);
			foreach (Office.TextRange2 run in from) {
				if (!(run.Start == from.Start && run.Length == from.Length)) {
					CopyCharacters(to.Characters[run.Start, run.Length], run);
				}
			}
		}

		public static void TransferCharacters(Office.TextRange2 to, Office.TextRange2 from, int toOffset) {
			foreach (Office.TextRange2 run in from) {
				Copy(to.Characters[run.Start + toOffset, run.Length].Font, run.Font);
			}
		}

		public static void ClearCharacters(Office.TextRange2 range) {
			var runs = new List<Office.TextRange2>();
			foreach (Office.TextRange2 run in range) {
				if (!(run.Start == range.Start && run.Length == range.Length)) {
					ClearCharacters(run);
					runs.Add(run);
				}
			}
			foreach (var run in runs) {
				run.Delete();
			}
			range.Font.Highlight.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorDark1;
		}

		public class ParagraphRangeSnapshot {
			public readonly ImmutableList<ParagraphFormatSnapshot> Formats;

			public ParagraphRangeSnapshot(IEnumerable<ParagraphFormatSnapshot> formats) {
				this.Formats = formats.ToImmutableList();
			}
		}

		public static ParagraphRangeSnapshot CaptureParagraph(Office.TextRange2 range) {
			if (range.Paragraphs.Count > 1) {
				return new ParagraphRangeSnapshot(
					from Office.TextRange2 para in range.Paragraphs
					select Capture(para.ParagraphFormat)
				);
			} else {
				return new ParagraphRangeSnapshot(Enumerable.Repeat(Capture(range.ParagraphFormat), 1));
			}
		}

		public static void ApplyParagraph(Office.TextRange2 to, ParagraphRangeSnapshot snapshot) {
			if (snapshot.Formats.Count > 1) {
				for (var i = 0; i < snapshot.Formats.Count; ++i) {
					Apply(to.Paragraphs[i + 1].ParagraphFormat, snapshot.Formats[i]);
				}
			} else {
				Apply(to.ParagraphFormat, snapshot.Formats[0]);
			}
		}

		public static void CopyParagraph(Office.TextRange2 to, Office.TextRange2 from) {
			ApplyParagraph(to, CaptureParagraph(from));
		}
		#endregion

		#region TextFrame2
		public class TextFrameSnapshot {
			public readonly string Text;
			public readonly ParagraphFormatSnapshot ParagraphFormat;
			public readonly CharacterRangeSnapshot CharacterRange;
			public readonly ParagraphRangeSnapshot ParagraphRange;

			public TextFrameSnapshot(string text, ParagraphFormatSnapshot paragraphFormat, CharacterRangeSnapshot characterRange, ParagraphRangeSnapshot paragraphRange) {
				this.Text = text;
				this.ParagraphFormat = paragraphFormat;
				this.CharacterRange = characterRange;
				this.ParagraphRange = paragraphRange;
			}

		}

		public static TextFrameSnapshot Capture(PowerPoint.TextFrame2 frame) {
			return new TextFrameSnapshot(
				frame.TextRange.Text,
				Capture(frame.TextRange.ParagraphFormat),
				CaptureCharacters(frame.TextRange),
				CaptureParagraph(frame.TextRange)
				);
		}

		public static void Apply(PowerPoint.TextFrame2 to, TextFrameSnapshot snapshot) {
			ClearCharacters(to.TextRange);
			to.TextRange.Text = snapshot.Text;
			Apply(to.TextRange.ParagraphFormat, snapshot.ParagraphFormat);
			ApplyCharacters(to.TextRange, snapshot.CharacterRange);
			ApplyParagraph(to.TextRange, snapshot.ParagraphRange);
		}

		public static void Copy(PowerPoint.TextFrame2 to, PowerPoint.TextFrame2 from) {
			Apply(to, Capture(from));
		}
		#endregion

		public static bool IsMetricRegion() {
			return System.Globalization.RegionInfo.CurrentRegion.IsMetric;
		}

		public static PowerPoint.DocumentWindow? GetActiveWindow() {
			var application = Globals.ThisAddIn.Application;
			if (application.Windows.Count == 0) { return null; }
			return application.ActiveWindow;
		}

		public static PowerPoint.Slide? GetActiveSlide() {
			var window = GetActiveWindow();
			if (window == null) { return null; }
			if (window.Presentation.Slides.Count == 0) { return null; }
			return (PowerPoint.Slide)window.View.Slide;
		}

		public static float? GetDefaultLength(string length) {
			return length switch {
				"small" => Settings.Default.default_small_length,
				"normal" => Settings.Default.default_normal_length,
				"large" => Settings.Default.default_large_length,
				_ => null
			};
		}

		public static float? GetDefaultLength(DefaultLength length) {
			return length switch {
				DefaultLength.Small => Settings.Default.default_small_length,
				DefaultLength.Normal => Settings.Default.default_normal_length,
				DefaultLength.Large => Settings.Default.default_large_length,
				_ => null
			};
		}

		public static string GetResourceString(string key, string defaultValue = "") {
			try {
				return Strings.ResourceManager.GetString(key, System.Globalization.CultureInfo.CurrentCulture);
			} catch (Exception ex) {
				Debug.WriteLine($"Missing resource string {key}: {ex.Message}");
			}
			return defaultValue;
		}

		private static readonly int MAX_CACHE_SIZE = 1024;
		private static readonly LinkedList<KeyValuePair<string, Bitmap?>> IMAGE_CACHE = new LinkedList<KeyValuePair<string, Bitmap?>>();

		private static Bitmap? DoLoadImageResource(string resourceName) {
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i) {
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
					try {
						return new Bitmap(Image.FromStream(asm.GetManifestResourceStream(resourceNames[i])));
					} catch (ArgumentException ex) {
						Debug.WriteLine($"Resource {resourceName} is not an image: {ex.Message}");
					}
				}
			}
			return null;
		}

		public static Bitmap? LoadImageResource(string resourceName) {
			var node = IMAGE_CACHE.First;
			while (node != null && (node.Value.Key != resourceName)) {
				node = node.Next;
			}
			if (node == null) {
				node = new LinkedListNode<KeyValuePair<string, Bitmap?>>(new KeyValuePair<string, Bitmap?>(
					resourceName, DoLoadImageResource(resourceName)
					));
				while (IMAGE_CACHE.Count > MAX_CACHE_SIZE - 1) {
					IMAGE_CACHE.RemoveLast();
				}
			} else {
				IMAGE_CACHE.Remove(node);
			}
			IMAGE_CACHE.AddFirst(node);
			return node.Value.Value;
		}

		/// <summary>
		/// Replaces all color in a image with the specified color and only preseving the alpha value.
		/// </summary>
		/// <param name="image">the image to modify</param>
		/// <param name="color">the new color value</param>
		/// <returns><c>image</c> where all rgb values are replaced with the rbg values from <c>color</c></returns>
		public static Bitmap ReplaceBitmapColor(Bitmap image, Color color) {
			return ReplaceBitmapColor(image, new Dictionary<Color, Color>(), color);
		}

		private static readonly ColorReplaceImageCache COLOR_REPLACEMENT_IMAGE_CACHE = new ColorReplaceImageCache();

		private static Bitmap DoReplaceBitmapColor(Bitmap image, ImmutableDictionary<int, int> lut, Color defaultColor) {
			Bitmap result = new Bitmap(image);

			var debug = new HashSet<Color>();

			for (var y = 0; y < result.Height; ++y) {
				for (var x = 0; x < result.Width; ++x) {
					var pixelColor = result.GetPixel(x, y);
					if (!debug.Contains(pixelColor)) {
						debug.Add(pixelColor);
					}
					var keyColor = Color.FromArgb(255, pixelColor.R, pixelColor.G, pixelColor.B);
					if (lut.TryGetValue(keyColor.ToArgb(), out int replaceRGB)) {
						var replaceColor = Color.FromArgb(replaceRGB);
						result.SetPixel(x, y, Color.FromArgb(pixelColor.A, replaceColor.R, replaceColor.G, replaceColor.B));
					} else {
						result.SetPixel(x, y, Color.FromArgb(pixelColor.A, defaultColor.R, defaultColor.G, defaultColor.B));
					}
				}
			}
			return result;
		}

		public static Bitmap ReplaceBitmapColor(Bitmap image, Dictionary<Color, Color> replacements, Color defaultColor) {
			var lut = (
				from entry in replacements
				select (Color.FromArgb(255, entry.Key.R, entry.Key.G, entry.Key.B).ToArgb(), entry.Value.ToArgb())
			).ToImmutableDictionary(entry => entry.Item1, entry => entry.Item2);
			var node = COLOR_REPLACEMENT_IMAGE_CACHE.First;
			while (node != null && (node.Value.Key.Item1 != image || !node.Value.Key.Item2.Equals(lut) || node.Value.Key.Item3 != defaultColor)) {
				node = node.Next;
			}
			if (node == null) {
				node = new LinkedListNode<KeyValuePair<Tuple<Bitmap, ImmutableDictionary<int, int>, Color>, Bitmap>>(
					new KeyValuePair<Tuple<Bitmap, ImmutableDictionary<int, int>, Color>, Bitmap>(
						Tuple.Create(image, lut, defaultColor), DoReplaceBitmapColor(image, lut, defaultColor)
				));
				while (COLOR_REPLACEMENT_IMAGE_CACHE.Count > MAX_CACHE_SIZE - 1) {
					COLOR_REPLACEMENT_IMAGE_CACHE.RemoveLast();
				}
			} else {
				COLOR_REPLACEMENT_IMAGE_CACHE.Remove(node);
			}
			COLOR_REPLACEMENT_IMAGE_CACHE.AddFirst(node);
			return node.Value.Value;
		}

		/// <summary>
		/// Returns the last item in a list that is below the query value.
		/// </summary>
		/// <typeparam name="T">list item type</typeparam>
		/// <param name="query">the query value</param>
		/// <param name="items">the list of items</param>
		/// <param name="mapper">maps a list item to a value</param>
		/// <param name="threshold">the epsilon threshold for the query value</param>
		/// <returns>the item with the highest mapping value smaller or equal to <c>query + threshold</c></returns>
		public static T? FloorItem<T>(double query, IEnumerable<T> items, Func<T, double> mapper, double threshold = 0.1)
			where T : class {
			var sorted = (
				from item in items
				where mapper(item) <= query + threshold
				orderby mapper(item) descending
				select item
			);

			return sorted.FirstOrDefault();
		}

		/// <summary>
		/// Returns the first item in a list that is above the query value.
		/// </summary>
		/// <typeparam name="T">list item type</typeparam>
		/// <param name="query">the query value</param>
		/// <param name="items">the list of items</param>
		/// <param name="mapper">maps a list item to a value</param>
		/// <param name="threshold">the epsilon threshold for the query value</param>
		/// <returns>the item with the lowest mapping value bigger or equal to <c>query - threshold</c></returns>
		public static T? CeilingItem<T>(double query, IEnumerable<T> items, Func<T, double> mapper, double threshold = 0.1)
		where T : class {
			var sorted = (
				from item in items
				where mapper(item) >= query - threshold
				orderby mapper(item) ascending
				select item
			);
			return sorted.FirstOrDefault();
		}

		/// <summary>
		/// Creates a modal window.
		/// </summary>
		/// <returns>the modal window</returns>
		public static Window CreateModalWindow() {
			var window = new Window() {
				SizeToContent = SizeToContent.WidthAndHeight,
				ResizeMode = ResizeMode.NoResize,
				ShowInTaskbar = false,
				WindowStartupLocation = WindowStartupLocation.CenterOwner,
			};
			new WindowInteropHelper(window) {
				Owner = Process.GetCurrentProcess().MainWindowHandle
			};
			return window;
		}

	}

}