#nullable enable

using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Linq;
using System.Drawing;

namespace hztoolbar
{
    public abstract class ToolbarAction
    {

        public readonly string Id;

        protected ToolbarAction(string id)
        {
            this.Id = id;
        }

        protected virtual IEnumerable<PowerPoint.Shape> GetSelectedShapes()
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return Enumerable.Empty<PowerPoint.Shape>();
            }

            return (
                from PowerPoint.Shape shape in selection.ShapeRange
                select shape
            );
        }

        public virtual Bitmap? GetImage(string controlId, string arg = "")
        {
            var result = Utils.LoadImageResource($"hztoolbar.icons.{controlId}.png");
            if (result == null)
            {
                result = Utils.LoadImageResource("hztoolbar.icons.question.png");
            }
            return result;
        }

        public virtual string? GetMsoImage(string controlId, string arg = "")
        {
            return null;
        }

        public virtual string GetLabel(string controlId, string arg = "")
        {
            return Utils.GetResourceString($"{controlId}_label");
        }

        public virtual string GetSupertip(string controlId, string arg = "")
        {
            return Utils.GetResourceString($"{controlId}_supertip");
        }

        public virtual bool IsEnabled(string arg = "")
        {
            return true;
        }

        public abstract bool Run(string arg = "");
    }
}
