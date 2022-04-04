
#nullable enable

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using System.Drawing;
using System.Collections.Generic;

namespace hztoolbar.actions
{


    public abstract class AbstractThemeColorAction : ToolbarAction
    {
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

        protected override IEnumerable<PowerPoint.Shape> GetSelectedShapes()
        {
            return from shape in base.GetSelectedShapes()
                   where shape.Type == Office.MsoShapeType.msoAutoShape
                   select shape;
        }

        public override Bitmap UpdateIcon(Bitmap image, string arg = "")
        {
            var color = GetColor(arg);
            if (color != null)
            {

                return Utils.ApplyColorToBitmap(image, Color.FromArgb(ColorTranslator.ToWin32(Color.FromArgb(color.Value))));
            }
            return image;
        }



        public override bool IsEnabled(string arg = "")
        {
            var shapes = GetSelectedShapes();
            return shapes.Take(1).Count() > 0;
        }

        protected int? GetColor(string arg)
        {
            var slide = Utils.GetActiveSlide();
            if (slide == null)
            {
                return null;
            }

            var themeColors = slide.ThemeColorScheme;
            if (themeColors == null)
            {
                return null;
            }

            Office.MsoThemeColorSchemeIndex? themeColorIndex = arg switch
            {
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

            if (themeColorIndex == null)
            {
                return null;
            }

            return themeColors.Colors(themeColorIndex.Value).RGB;
        }

    }


    public class ApplyBackgroundThemeColorAction : AbstractThemeColorAction
    {

        public ApplyBackgroundThemeColorAction() : base("apply_background_theme_color") { }

        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes();
            var color = GetColor(arg);

            if (color != null)
            {
                var textColor = Color.FromArgb(color.Value).GetBrightness() > 0.5
                    ? Color.Black.ToArgb()
                    : Color.White.ToArgb();

                foreach (var shape in shapes)
                {
                    shape.Fill.Solid();
                    shape.Fill.BackColor.RGB = color.Value;
                    shape.Fill.ForeColor.RGB = color.Value;

                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue)
                    {
                        shape.TextFrame.TextRange.Font.Color.RGB = textColor;
                    }

                }
            }
        }
    }

    public class ApplyStrokeThemeColorAction : AbstractThemeColorAction
    {

        public ApplyStrokeThemeColorAction() : base("apply_line_theme_color") { }

        public override void Run(string arg = "")
        {
            var shapes = GetSelectedShapes();
            var color = GetColor(arg);

            if (color != null)
            {
                var textColor = Color.FromArgb(color.Value).GetBrightness() > 0.5
                    ? Color.Black.ToArgb()
                    : Color.White.ToArgb();

                foreach (var shape in shapes)
                {
                    shape.Line.BackColor.RGB = color.Value;
                    shape.Line.ForeColor.RGB = color.Value;

                }
            }
        }
    }
}