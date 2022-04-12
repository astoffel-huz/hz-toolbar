#nullable enable

using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Interop;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace hztoolbar {

	public static class Utils {

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

			public FontEmphasisSnapshot(Office.MsoTriState bold, Office.MsoTriState italic, Office.MsoTextStrike strike,
				Office.MsoTextUnderlineType underlineStyle, Office.MsoTriState doubleStrikeThrough, Office.MsoTriState subscript, Office.MsoTriState superscript) {
				this.Bold = bold;
				this.Italic = italic;
				this.Strike = strike;
				this.UnderlineStyle = underlineStyle;
				this.DoubleStrikeThrough = doubleStrikeThrough;
				this.Subscript = subscript;
				this.Superscript = superscript;
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
				font.Superscript
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
		}

		public static void Copy(Office.Font2 to, Office.Font2 from) {
			to.Bold = from.Bold;
			to.Italic = from.Italic;
			to.Strike = from.Strike;
			to.UnderlineStyle = from.UnderlineStyle;
			to.DoubleStrikeThrough = from.DoubleStrikeThrough;
			to.Subscript = from.Subscript;
			to.Superscript = from.Superscript;
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


		public static PowerPoint.Slide? GetActiveSlide() {
			var application = Globals.ThisAddIn.Application;
			try {
				return (PowerPoint.Slide)application.ActiveWindow.View.Slide;
			} catch {
				return null;
			}
		}

		public static string GetResourceString(string key, string defaultValue = "") {
			try {
				return Strings.ResourceManager.GetString(key, System.Globalization.CultureInfo.CurrentCulture);
			} catch (Exception ex) {
				Debug.WriteLine($"Missing resource string {key}: {ex.Message}");
			}
			return defaultValue;
		}

		private static T? ResourceCacheLookup<T>(LinkedList<KeyValuePair<string, T?>> cache, string key, Func<string, T?> generator) where T : class {
			const int MAX_CACHE_SIZE = 128;
			var node = cache.First;
			while (node != null && node.Value.Key != key) {
				node = node.Next;
			}
			if (node == null) {
				node = new LinkedListNode<KeyValuePair<string, T?>>(new KeyValuePair<string, T?>(key, generator(key)));
				while (cache.Count >= MAX_CACHE_SIZE - 1) {
					cache.RemoveLast();
				}
			} else {
				cache.Remove(node);
			}
			cache.AddFirst(node);
			return node.Value.Value;
		}

		private static LinkedList<KeyValuePair<string, Bitmap?>> IMAGE_CACHE = new LinkedList<KeyValuePair<string, Bitmap?>>();

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
			return ResourceCacheLookup(IMAGE_CACHE, resourceName, DoLoadImageResource);
		}

		private static LinkedList<KeyValuePair<Tuple<Bitmap, int>, Bitmap>> COLORED_BITMAP_CACHE = new LinkedList<KeyValuePair<Tuple<Bitmap, int>, Bitmap>>();

		private static Bitmap DoApplyColorToBitmap(Bitmap image, Color color) {
			Bitmap result = new Bitmap(image);
			for (var y = 0; y < result.Height; ++y) {
				for (var x = 0; x < result.Width; ++x) {
					var pixelColor = result.GetPixel(x, y);
					result.SetPixel(x, y, Color.FromArgb(pixelColor.A, color.R, color.G, color.B));
				}
			}
			return result;
		}


		public static Bitmap ReplaceBitmapColor(Bitmap image, Color color) {
			var node = COLORED_BITMAP_CACHE.First;
			while (node != null && !(node.Value.Key.Item1 == image && node.Value.Key.Item2 == color.ToArgb())) {
				node = node.Next;
			}

			if (node == null) {
				node = new LinkedListNode<KeyValuePair<Tuple<Bitmap, int>, Bitmap>>(new KeyValuePair<Tuple<Bitmap, int>, Bitmap>(
					Tuple.Create(image, color.ToArgb()), DoApplyColorToBitmap(image, color)
				));
			} else {
				COLORED_BITMAP_CACHE.Remove(node);
			}
			COLORED_BITMAP_CACHE.AddFirst(node);

			Bitmap result = new Bitmap(image);
			for (var y = 0; y < result.Height; ++y) {
				for (var x = 0; x < result.Width; ++x) {
					var pixelColor = result.GetPixel(x, y);
					result.SetPixel(x, y, Color.FromArgb(pixelColor.A, color.R, color.G, color.B));
				}
			}
			return result;
		}

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

		public static Window CreateModalWindow() {
			var window = new Window() {
				SizeToContent = SizeToContent.WidthAndHeight,
				ResizeMode = ResizeMode.NoResize,
				ShowInTaskbar = false,
				WindowStartupLocation = WindowStartupLocation.CenterOwner,
			};
			var interop = new WindowInteropHelper(window);
			interop.Owner = Process.GetCurrentProcess().MainWindowHandle;			
			return window;
		}
	}

}