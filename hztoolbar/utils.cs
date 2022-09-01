#nullable enable

using HandyControl.Data;
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
using System.Windows.Markup;
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


		#region BulletFormat2
		public class BulletFormatSnapshot {
			public readonly int Character;
			public readonly float RelativeSize;
			public readonly int StartValue;
			public readonly string FontName;
			public readonly float FontSize;
			public readonly Office.MsoNumberedBulletStyle Style;
			public readonly Office.MsoBulletType Type;

			public BulletFormatSnapshot(int character, float relativeSize, int startValue, string fontName, float fontSize,
				Office.MsoNumberedBulletStyle style, Office.MsoBulletType type) {
				this.Character = character;
				this.RelativeSize = relativeSize;
				this.StartValue = startValue;
				this.FontName = fontName;
				this.FontSize = fontSize;
				this.Style = style;
				this.Type = type;
			}
		}

		public static BulletFormatSnapshot Capture(Office.BulletFormat2 format) {
			return new BulletFormatSnapshot(
				format.Character,
				format.RelativeSize,
				format.StartValue,
				format.Font.Name,
				format.Font.Size,
				format.Style,
				format.Type
				);
		}

		public static void Apply(Office.BulletFormat2 to, BulletFormatSnapshot snapshot) {
			to.Character = snapshot.Character;
			to.RelativeSize = snapshot.RelativeSize;
			to.StartValue = snapshot.StartValue;
			to.Font.Name = snapshot.FontName;
			to.Font.Size = snapshot.FontSize;
			to.Style = snapshot.Style;
			switch (snapshot.Type) {
				case MsoBulletType.msoBulletNone:
				case MsoBulletType.msoBulletUnnumbered:
				case MsoBulletType.msoBulletNumbered:
					to.Type = snapshot.Type;
					break;
				default:
					to.Type = MsoBulletType.msoBulletUnnumbered;
					break;
			}
		}

		public static void Copy(Office.BulletFormat2 to, Office.BulletFormat2 from) {
			to.Character = from.Character;
			to.RelativeSize = from.RelativeSize;
			to.StartValue = from.StartValue;
			to.Font.Name = from.Font.Name;
			to.Font.Size = from.Font.Size;
			to.Style = from.Style;
			switch (from.Type) {
				case MsoBulletType.msoBulletNone:
				case MsoBulletType.msoBulletUnnumbered:
				case MsoBulletType.msoBulletNumbered:
					to.Type = from.Type;
					break;
				default:
					to.Type = MsoBulletType.msoBulletUnnumbered;
					break;
			}
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
			Clear(to);
			foreach (Office.TabStop2 tabStop in from) {
				switch (tabStop.Type) {
					case Office.MsoTabStopType.msoTabStopLeft:
					case Office.MsoTabStopType.msoTabStopCenter:
					case Office.MsoTabStopType.msoTabStopRight:
					case Office.MsoTabStopType.msoTabStopDecimal:
						to.Add(tabStop.Type, tabStop.Position);
						break;
				};
			}
		}
		#endregion

		#region ParagraphFormat2

		/// <summary>
		/// Snapshot of a paragraph format.
		/// </summary>
		public class ParagraphFormatSnapshot {
			public readonly Office.MsoParagraphAlignment Alignment;
			public readonly Office.MsoBaselineAlignment BaselineAlignment;
			public readonly BulletFormatSnapshot Bullet;
			public readonly float FirstLineIndent;
			public readonly Office.MsoTriState HangingPunctuation;
			public readonly int IndentLevel;
			public readonly float LeftIndent;
			public readonly TabStopsSnapshot TabStops;

			public ParagraphFormatSnapshot(Office.MsoParagraphAlignment alignment, Office.MsoBaselineAlignment baselineAlignment,
				BulletFormatSnapshot bullet, float firstLineIndent, Office.MsoTriState hangingPunctuation, int indentLevel, float leftIndent,
				TabStopsSnapshot tabStops) {
				this.Alignment = alignment;
				this.BaselineAlignment = baselineAlignment;
				this.Bullet = bullet;
				this.FirstLineIndent = firstLineIndent;
				this.HangingPunctuation = hangingPunctuation;
				this.IndentLevel = indentLevel;
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
				format.IndentLevel,
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
			to.FirstLineIndent = snapshot.FirstLineIndent;
			to.HangingPunctuation = snapshot.HangingPunctuation;
			to.IndentLevel = snapshot.IndentLevel;
			to.LeftIndent = snapshot.LeftIndent;
			Apply(to.TabStops, snapshot.TabStops);
		}

		/// <summary>
		/// Copies a paragraph format.
		/// </summary>
		/// <param name="to">the target paragraph format</param>
		/// <param name="from">the source paragraph format to copy</param>
		public static void Copy(Office.ParagraphFormat2 to, Office.ParagraphFormat2 from) {
			to.Alignment = from.Alignment;
			to.BaselineAlignment = from.BaselineAlignment;
			Copy(to.Bullet, from.Bullet);
			to.FirstLineIndent = from.FirstLineIndent;
			to.HangingPunctuation = from.HangingPunctuation;
			to.IndentLevel = from.IndentLevel;
			to.LeftIndent = from.LeftIndent;
			Copy(to.TabStops, from.TabStops);

		}
		#endregion

		#region Font2
		public class FontEmphasisSnapshot {
			public readonly Office.MsoTriState Bold;
			public readonly Office.MsoTriState Italic;
			public readonly Office.MsoTextStrike Strike;
			public readonly Office.MsoTextUnderlineType UnderlineStyle;
			public readonly Office.MsoTriState DoubleStrikeThrough;
			public readonly Office.MsoTriState Subscript;
			public readonly Office.MsoTriState Superscript;
			public readonly string Name;
			public readonly float Size;

			public FontEmphasisSnapshot(Office.MsoTriState bold, Office.MsoTriState italic, Office.MsoTextStrike strike,
				Office.MsoTextUnderlineType underlineStyle, Office.MsoTriState doubleStrikeThrough, Office.MsoTriState subscript, Office.MsoTriState superscript,
				string name, float size) {
				this.Bold = bold;
				this.Italic = italic;
				this.Strike = strike;
				this.UnderlineStyle = underlineStyle;
				this.DoubleStrikeThrough = doubleStrikeThrough;
				this.Subscript = subscript;
				this.Superscript = superscript;
				this.Name = name;
				this.Size = size;
			}
		}

		public static FontEmphasisSnapshot Capture(Office.Font2 font) {
			return new FontEmphasisSnapshot(
				font.Bold,
				font.Italic,
				font.Strike,
				font.UnderlineStyle,
				font.DoubleStrikeThrough,
				font.Subscript,
				font.Superscript,
				font.Name,
				font.Size
				);
		}

		public static void Apply(Office.Font2 to, FontEmphasisSnapshot snapshot) {
			to.Bold = snapshot.Bold;
			to.Italic = snapshot.Italic;
			to.Strike = snapshot.Strike;
			to.UnderlineStyle = snapshot.UnderlineStyle;
			to.DoubleStrikeThrough = snapshot.DoubleStrikeThrough;
			to.Subscript = snapshot.Subscript;
			to.Superscript = snapshot.Superscript;
			to.Name = snapshot.Name;
			to.Size = snapshot.Size;
		}

		public static void Copy(Office.Font2 to, Office.Font2 from) {
			to.Bold = from.Bold;
			to.Italic = from.Italic;
			to.Strike = from.Strike;
			to.UnderlineStyle = from.UnderlineStyle;
			to.DoubleStrikeThrough = from.DoubleStrikeThrough;
			to.Subscript = from.Subscript;
			to.Superscript = from.Superscript;
			to.Name = from.Name;
			to.Size = from.Size;
		}
		#endregion

		#region TextRange2
		public class CharacterRangeSnapshot {
			public readonly int Start;
			public readonly int Length;
			public readonly FontEmphasisSnapshot Emphasis;
			public readonly ImmutableList<CharacterRangeSnapshot> Runs;

			public CharacterRangeSnapshot(FontEmphasisSnapshot emphasis, IEnumerable<CharacterRangeSnapshot> runs) {
				this.Emphasis = emphasis;
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
			Apply(to.Font, snapshot.Emphasis);
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
			if (from.Paragraphs.Count > 1) {
				for (var i = 0; i < from.Paragraphs.Count; ++i) {
					CopyParagraph(to.Paragraphs[i + 1], from.Paragraphs[i + 1]);
				}
			} else {
				Copy(to.ParagraphFormat, from.ParagraphFormat);
			}
		}
		#endregion

		#region TextFrame2
		public class TextFrameSnapshot {
			public readonly string Text;
			public readonly CharacterRangeSnapshot CharacterRange;
			public readonly ParagraphRangeSnapshot ParagraphRange;

			public TextFrameSnapshot(string text, CharacterRangeSnapshot characterRange, ParagraphRangeSnapshot paragraphRange) {
				this.Text = text;
				this.CharacterRange = characterRange;
				this.ParagraphRange = paragraphRange;
			}

		}

		public static TextFrameSnapshot Capture(PowerPoint.TextFrame2 frame) {
			return new TextFrameSnapshot(
				frame.TextRange.Text,
				CaptureCharacters(frame.TextRange),
				CaptureParagraph(frame.TextRange)
				);
		}

		public static void Apply(PowerPoint.TextFrame2 to, TextFrameSnapshot snapshot) {
			ClearCharacters(to.TextRange);
			to.TextRange.Text = snapshot.Text;
			ApplyCharacters(to.TextRange, snapshot.CharacterRange);
			ApplyParagraph(to.TextRange, snapshot.ParagraphRange);
		}

		public static void Copy(PowerPoint.TextFrame2 to, PowerPoint.TextFrame2 from) {
			ClearCharacters(to.TextRange);
			to.TextRange.Text = from.TextRange.Text;
			CopyCharacters(to.TextRange, from.TextRange);
			CopyParagraph(to.TextRange, from.TextRange);
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