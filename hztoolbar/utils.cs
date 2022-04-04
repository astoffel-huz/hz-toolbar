#nullable enable

using System;
using System.Linq;
using System.Diagnostics;
using System.Reflection;
using System.Collections.Generic;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace hztoolbar
{

    public static class Utils
    {

        public struct ShadowFormatSnapshot
        {
            //public float Blur;
            //public Office.MsoTriState Obscured;
            //public float OffsetX;
            //public float OffsetY;
            //public float Size;
            //public Office.MsoShadowStyle Style;
            //public float Transparency;
            //public Office.MsoShadowType Type;
            public Office.MsoTriState Visible;
        }

        public struct FontSnapshot
        {
            public Office.MsoTriState Bold;
            public Office.MsoTextCaps Caps;
            public Office.MsoTriState DoubleStrikeThrough;
            public Office.MsoTriState Italic;
            public ShadowFormatSnapshot Shadow;
            public Office.MsoTriState Smallcaps;
            public Office.MsoSoftEdgeType SoftEdgeFormat;
            public Office.MsoTextStrike Strike;
            public Office.MsoTriState StrikeThrough;
            public Office.MsoTriState Subscript;
            public Office.MsoTriState Superscript;
            public Office.MsoTextUnderlineType UnderlineStyle;
        }

        public struct TextRangeSnapshot
        {
            public FontSnapshot Font;
            public List<FontSnapshot> Characters;
        }

        public struct TextFrameSnapshot
        {
            public string Text;
            public TextRangeSnapshot Range;
        }


        public static ShadowFormatSnapshot Capture(Office.ShadowFormat shadow)
        {
            return new ShadowFormatSnapshot()
            {
                Visible = shadow.Visible,
            };
        }

        public static FontSnapshot Capture(Office.Font2 font)
        {
            return new FontSnapshot()
            {
                Bold = font.Bold,
                Caps = font.Caps,
                DoubleStrikeThrough = font.DoubleStrikeThrough,
                Italic = font.Italic,
                Shadow = Capture(font.Shadow),
                Smallcaps = font.Smallcaps,
                SoftEdgeFormat = font.SoftEdgeFormat,
                Strike = font.Strike,
                StrikeThrough = font.StrikeThrough,
                Subscript = font.Subscript,
                Superscript = font.Superscript,
                UnderlineStyle = font.UnderlineStyle,
            };
        }

        public static TextRangeSnapshot Capture(Office.TextRange2 range)
        {
            var characters = new List<FontSnapshot>();
            for (var idx = 0; idx < range.Length; ++idx)
            {
                characters.Add(Capture(range.Characters[range.Start + idx].Font));
            }
            return new TextRangeSnapshot()
            {
                Font = Capture(range.Font),
                Characters = characters
            };
        }

        public static TextFrameSnapshot Capture(PowerPoint.TextFrame2 frame)
        {
            return new TextFrameSnapshot()
            {
                Text = frame.TextRange.Text,
                Range = Capture(frame.TextRange),
            };
        }

        public static void Apply(Office.ShadowFormat shadow, ShadowFormatSnapshot snapshot)
        {
            shadow.Visible = snapshot.Visible;
        }

        public static void Apply(Office.Font2 font, FontSnapshot snapshot)
        {
            font.Bold = snapshot.Bold;
            font.Caps = snapshot.Caps;
            font.DoubleStrikeThrough = snapshot.DoubleStrikeThrough;
            font.Italic = snapshot.Italic;
            Apply(font.Shadow, snapshot.Shadow);
            font.Smallcaps = snapshot.Smallcaps;
            font.SoftEdgeFormat = snapshot.SoftEdgeFormat;
            font.Strike = snapshot.Strike;
            font.StrikeThrough = snapshot.StrikeThrough;
            font.Subscript = snapshot.Subscript;
            font.Superscript = snapshot.Superscript;
            font.UnderlineStyle = snapshot.UnderlineStyle;
        }

        public static void Apply(Office.TextRange2 range, TextRangeSnapshot snapshot)
        {
            Apply(range.Font, snapshot.Font);
            for (var idx = 0; idx < snapshot.Characters.Count; ++idx)
            {
                Apply(range.Characters[range.Start + idx].Font, snapshot.Characters[idx]);
            }
        }

        public static void Apply(PowerPoint.TextFrame2 frame, TextFrameSnapshot snapshot)
        {
            frame.DeleteText();
            frame.TextRange.Text = snapshot.Text;
            Apply(frame.TextRange, snapshot.Range);
        }

        public static void Copy(Office.ShadowFormat to, Office.ShadowFormat from)
        {
            to.Visible = from.Visible;
        }

        public static void Copy(Office.Font2 to, Office.Font2 from)
        {
            to.Bold = from.Bold;
            to.Caps = from.Caps;
            to.DoubleStrikeThrough = from.DoubleStrikeThrough;
            to.Italic = from.Italic;
            Copy(to.Shadow, from.Shadow);
            to.Smallcaps = from.Smallcaps;
            to.SoftEdgeFormat = from.SoftEdgeFormat;
            to.Strike = from.Strike;
            to.StrikeThrough = from.StrikeThrough;
            to.Subscript = from.Subscript;
            to.Superscript = from.Superscript;
            to.UnderlineStyle = from.UnderlineStyle;
        }

        public static void Copy(Office.TextRange2 to, Office.TextRange2 from)
        {
            Copy(to.Font, from.Font);
            for (var idx = 0; idx < from.Length; ++idx)
            {
                Copy(to.Characters[to.Start + idx].Font, from.Characters[from.Start + idx].Font);
            }
        }

        public static void Copy(PowerPoint.TextFrame2 to, PowerPoint.TextFrame2 from)
        {
            to.DeleteText();
            to.TextRange.Text = from.TextRange.Text;
            Copy(to.TextRange, from.TextRange);
        }


        public static PowerPoint.Slide? GetActiveSlide()
        {
            var application = Globals.ThisAddIn.Application;
            try
            {
                return (PowerPoint.Slide)application.ActiveWindow.View.Slide;
            }
            catch
            {
                return null;
            }
        }

        public static string GetResourceString(string key, string defaultValue = "")
        {
            try
            {
                return Strings.ResourceManager.GetString(key, System.Globalization.CultureInfo.CurrentCulture);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Missing resource string {key}: {ex.Message}");
            }
            return defaultValue;
        }

        private static T? ResourceCacheLookup<T>(LinkedList<KeyValuePair<string, T?>> cache, string key, Func<string, T?> generator) where T : class
        {
            const int MAX_CACHE_SIZE = 128;
            var node = cache.First;
            while (node != null && node.Value.Key != key)
            {
                node = node.Next;
            }
            if (node == null)
            {
                node = new LinkedListNode<KeyValuePair<string, T?>>(new KeyValuePair<string, T?>(key, generator(key)));
                while (cache.Count >= MAX_CACHE_SIZE - 1)
                {
                    cache.RemoveLast();
                }
            }
            else
            {
                cache.Remove(node);
            }
            cache.AddFirst(node);
            return node.Value.Value;
        }

        private static LinkedList<KeyValuePair<string, Image?>> IMAGE_CACHE = new LinkedList<KeyValuePair<string, Image?>>();

        private static Image? DoLoadImageResource(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    try
                    {
                        return Image.FromStream(asm.GetManifestResourceStream(resourceNames[i]));
                    }
                    catch (ArgumentException ex)
                    {
                        Debug.WriteLine($"Resource {resourceName} is not an image: {ex.Message}");
                    }
                }
            }
            return null;
        }

        public static Image? LoadImageResource(string resourceName)
        {
            return ResourceCacheLookup(IMAGE_CACHE, resourceName, DoLoadImageResource);
        }

        private static LinkedList<KeyValuePair<Tuple<Bitmap, int>, Bitmap>> COLORED_BITMAP_CACHE = new LinkedList<KeyValuePair<Tuple<Bitmap, int>, Bitmap>>();

        private static Bitmap DoApplyColorToBitmap(Bitmap image, Color color)
        {
            Bitmap result = new Bitmap(image);
            for (var y = 0; y < result.Height; ++y)
            {
                for (var x = 0; x < result.Width; ++x)
                {
                    var pixelColor = result.GetPixel(x, y);
                    result.SetPixel(x, y, Color.FromArgb(pixelColor.A, color.R, color.G, color.B));
                }
            }
            return result;
        }


        public static Bitmap ApplyColorToBitmap(Bitmap image, Color color)
        {
            var node = COLORED_BITMAP_CACHE.First;
            while (node != null && !(node.Value.Key.Item1 == image && node.Value.Key.Item2 == color.ToArgb()))
            {
                node = node.Next;
            }

            if (node == null)
            {
                node = new LinkedListNode<KeyValuePair<Tuple<Bitmap, int>, Bitmap>>(new KeyValuePair<Tuple<Bitmap, int>, Bitmap>(
                    Tuple.Create(image, color.ToArgb()), DoApplyColorToBitmap(image, color)
                ));
            }
            else
            {
                COLORED_BITMAP_CACHE.Remove(node);
            }
            COLORED_BITMAP_CACHE.AddFirst(node);

            Bitmap result = new Bitmap(image);
            for (var y = 0; y < result.Height; ++y)
            {
                for (var x = 0; x < result.Width; ++x)
                {
                    var pixelColor = result.GetPixel(x, y);
                    result.SetPixel(x, y, Color.FromArgb(pixelColor.A, color.R, color.G, color.B));
                }
            }
            return result;
        }
    }

}