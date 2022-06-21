#nullable enable

using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace hztoolbar {
	/// <summary>
	/// Base class for H&amp;Z toolbar actions.
	/// </summary>
	public abstract class ToolbarAction {

		public readonly string Id;

		protected ToolbarAction(string id) {
			this.Id = id;
		}

		protected virtual IEnumerable<PowerPoint.Shape> GetSelectedShapes() {
			var activeWindow = Utils.GetActiveWindow();
			if (activeWindow == null) {
				return Enumerable.Empty<PowerPoint.Shape>();
			}

			var selection = activeWindow.Selection;

			if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) {
				return Enumerable.Empty<PowerPoint.Shape>();
			}

			return (
				from PowerPoint.Shape shape in selection.ShapeRange
				select shape
			);
		}

		public virtual Bitmap? GetImage(string controlId, string arg = "") {
			var result = Utils.LoadImageResource($"hztoolbar.icons.{controlId}.png");
			if (result == null) {
				result = Utils.LoadImageResource("hztoolbar.icons.question.png");
			}
			return result;
		}

		public virtual string? GetMsoImage(string controlId, string arg = "") {
			return null;
		}

		public virtual string GetLabel(string controlId, string arg = "") {
			return Utils.GetResourceString($"{controlId}_label");
		}

		public virtual string GetSupertip(string controlId, string arg = "") {
			return Utils.GetResourceString($"{controlId}_supertip");
		}

		public virtual bool IsEnabled(string arg = "") {
			return true;
		}

		public abstract bool Run(string arg = "");
	}
}
