#nullable enable

using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace hztoolbar {
	/// <summary>
	/// Base class for H&amp;Z toolbar actions.
	/// 
	/// Toolbar actions are registered automatically with the toolbar, if placed as public class with a default constructor 
	/// in the namespace <c>hztoolbar.actions</c>. The id of a toolbar is used to identify the action class. The <c>arg</c> 
	/// argument is used to differientiate the handling code.
	/// </summary>
	public abstract class ToolbarAction {

		/// <summary>
		/// Unique identifier of this action.
		/// </summary>
		public readonly string Id;

		/// <summary>
		/// Initializes the action.
		/// </summary>
		/// <param name="id">the unique id of this action</param>
		protected ToolbarAction(string id) {
			this.Id = id;
		}

		/// <summary>
		/// Returns an enumerable of all selected shapes.
		/// </summary>
		/// <returns>the enumeration of all selected shapes</returns>
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

		/// <summary>
		/// Returns the item bitmap icon of this action
		/// </summary>
		/// 
		/// The default implementation returns the resource png named after the controlId  in the <c>hztoolbar.icons</c> namespace.
		/// 
		/// <param name="controlId">the id of the control</param>
		/// <param name="arg">the action argument</param>
		/// <returns>the bitmap of the icon</returns>
		public virtual Bitmap? GetImage(string controlId, string arg = "") {
			var result = Utils.LoadImageResource($"hztoolbar.icons.{controlId}.png");
			if (result == null) {
				result = Utils.LoadImageResource("hztoolbar.icons.question.png");
			}
			return result;
		}


		/// <summary>
		/// Returns the mso image id for this action.
		/// </summary>
		/// <param name="controlId">the id of the control</param>
		/// <param name="arg">the action argument</param>
		/// <returns>the id of the mso image</returns>
		public virtual string? GetMsoImage(string controlId, string arg = "") {
			return null;
		}

		/// <summary>
		/// Returns the label for this action.
		/// </summary>
		/// 
		/// By default the string resource named <c>{controlId}_label</c> is returned.
		/// 
		/// <param name="controlId">the control id</param>
		/// <param name="arg">the action argument</param>
		/// <returns>the label of this action</returns>
		public virtual string GetLabel(string controlId, string arg = "") {
			return Utils.GetResourceString($"{controlId}_label");
		}

		/// <summary>
		/// Returns the supertip for this action.
		/// </summary>
		/// 
		/// By default the string resource named <c>{controlId}_supertip</c> is returned.
		/// 
		/// <param name="controlId">the control id</param>
		/// <param name="arg">the action argument</param>
		/// <returns>the supertip of this action</returns>
		public virtual string GetSupertip(string controlId, string arg = "") {
			return Utils.GetResourceString($"{controlId}_supertip");
		}

		protected (string, string) SplitArg(string arg, (string, string) defaultValues) {
			var result = arg.Split('+');
			if (result.Length != 2) {
				return defaultValues;
			}
			return (result[0], result[1]);
		}

		protected (string, string, string) SplitArg(string arg, (string, string, string) defaultValues) {
			var result = arg.Split('+');
			if (result.Length != 3) {
				return defaultValues;
			}
			return (result[0], result[1], result[2]);
		}

		/// <summary>
		/// Indicates whether this action is enabled.
		/// </summary>
		/// 
		/// By default all actions are enabled.
		/// 
		/// <param name="arg">the action argument</param>
		/// <returns>whether this action is enabled</returns>
		public virtual bool IsEnabled(string arg = "") {
			return true;
		}

		/// <summary>
		/// Execute the action.
		/// </summary>
		/// <param name="arg">action argument</param>
		/// <returns><c>true</c> if the toolbar state should be revalidated</returns>
		public abstract bool Run(string arg = "");
	}
}
