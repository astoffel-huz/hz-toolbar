#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Immutable;

namespace hztoolbar.actions
{

    public abstract class AbstractTextAction : ToolbarAction
    {
        protected AbstractTextAction(string id) : base(id)
        {
        }

        protected override IEnumerable<Shape> GetSelectedShapes()
        {
            return (
                from shape in base.GetSelectedShapes()
                where shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue
                select shape
               );
        }

        public override bool IsEnabled(string arg = "")
        {
            return GetSelectedShapes().Take(1).Count() > 0;
        }

    }


    public class ShapeText : AbstractTextAction
    {
        public ShapeText() : base("swap_text") { }

        public override bool IsEnabled(string arg = "")
        {
            return GetSelectedShapes().Take(2).Count() >= 2;
        }

        public override void Run(string arg = "")
        {
            var textShapes = GetSelectedShapes().ToList();
            if (textShapes.Count > 1)
            {
                var snapshot = Utils.Capture(textShapes[textShapes.Count - 1].TextFrame2);
                for (var i = textShapes.Count - 1; i > 0; --i)
                {
                    Utils.Copy(textShapes[i].TextFrame2, textShapes[i - 1].TextFrame2);
                }
                Utils.Apply(textShapes[0].TextFrame2, snapshot);
            }
        }
    }

    public class ClearTextAction : AbstractTextAction
    {
        public ClearTextAction() : base("clear_text") { }

        public override void Run(string arg = "")
        {
            foreach (var shape in GetSelectedShapes())
            {
                shape.TextFrame.DeleteText();
            }
        }
    }

    public class ChangeLanguage : AbstractTextAction
    {
        private readonly ImmutableDictionary<string, Office.MsoLanguageID> LANGUAGES = new Dictionary<string, Office.MsoLanguageID>()
        {
            ["de"] = Office.MsoLanguageID.msoLanguageIDGerman,
            ["de-DE"] = Office.MsoLanguageID.msoLanguageIDGerman,

            ["en"] = Office.MsoLanguageID.msoLanguageIDEnglishUS,
            ["en-UK"] = Office.MsoLanguageID.msoLanguageIDEnglishUK,
            ["en-US"] = Office.MsoLanguageID.msoLanguageIDEnglishUS,

        }.ToImmutableDictionary();

        public ChangeLanguage() : base("change_language") { }

        public override bool IsEnabled(string arg = "")
        {
            return LANGUAGES.ContainsKey(arg) && (
                GetSelectedShapes().Take(1).Count() > 0
                || Utils.GetActiveSlide() != null
               );
        }

        public override void Run(string arg = "")
        {
            if (LANGUAGES.TryGetValue(arg, out var language))
            {
                var shapes = GetSelectedShapes().ToList();
                if (shapes.Count == 0)
                {
                    var slide = Utils.GetActiveSlide();
                    if (slide != null)
                    {
                        shapes = (
                            from PowerPoint.Shape shape in slide.Shapes
                            where shape.HasTextFrame == Office.MsoTriState.msoTrue || shape.HasTextFrame == Office.MsoTriState.msoCTrue
                            select shape
                        ).ToList();
                    }
                }
                foreach (var shape in shapes)
                {
                    shape.TextFrame2.TextRange.LanguageID = language;
                }
            }
        }
    }
}

