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

        public virtual Bitmap UpdateIcon(Bitmap image, string arg = "")
        {
            return image;
        }

        public virtual bool IsEnabled(string arg = "")
        {
            return true;
        }

        public abstract void Run(string arg = "");
    }
}
