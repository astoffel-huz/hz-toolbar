
#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using hztoolbar.actions;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Collections.Immutable;
using System.Windows;
using System.Windows.Interop;
using hztoolbar.Properties;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace hztoolbar
{
    [ComVisible(true)]
    public class ToolbarRibbon : Office.IRibbonExtensibility
    {

        private static readonly Regex ACTION_TAG_PATTERN = new Regex(@"(?<action_id>[a-z_]+)(?::(?<parameter>.+))?");

        private static IEnumerable<ToolbarAction> EnumerateActions()
        {
            return from ctor in (
                    from t in Assembly.GetExecutingAssembly().GetTypes()
                    where t.Namespace == "hztoolbar.actions"
                    where !t.IsAbstract
                    where typeof(ToolbarAction).IsAssignableFrom(t)
                    select t.GetConstructor(Type.EmptyTypes)
                )
                   where ctor != null
                   select (ToolbarAction)ctor.Invoke(new object[] { });
        }

        private Office.IRibbonUI? ribbon;

        private readonly ImmutableDictionary<string, ToolbarAction> actions;

        public ToolbarRibbon()
        {
            this.actions = EnumerateActions().ToImmutableDictionary(it => it.Id);
        }

        private R WithActionDo<R>(Office.IRibbonControl control, Func<ToolbarAction, string, R> action, R defaultResult)
        {
            var actionId = control.Id;
            var parameter = "";
            if (control.Tag != null && control.Tag.Length > 0)
            {
                var match = ACTION_TAG_PATTERN.Match(control.Tag);
                if (match.Success)
                {
                    actionId = match.Groups["action_id"].Value;
                    var parameter_group = match.Groups["parameter"];
                    if (parameter_group != null)
                    {
                        parameter = parameter_group.Value;
                    }
                }
            }
            if (this.actions.TryGetValue(actionId, out ToolbarAction? toolbarAction))
            {
                return action(toolbarAction, parameter);
            }
            else
            {
                Debug.WriteLine($"Unknown toolbar action {actionId}");
            }
            return defaultResult;
        }

        public void OnAction(Office.IRibbonControl control)
        {
            var invalidate = WithActionDo<bool>(control, (action, param) => action.Run(param), false); ;
            if (invalidate && this.ribbon != null)
            {
                this.ribbon.Invalidate();
            }
        }

        public bool IsEnabled(Office.IRibbonControl control)
        {
            return WithActionDo(control, (action, param) => action.IsEnabled(param), false);
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            return WithActionDo<string>(control, (action, param) => action.GetLabel(control.Id, param), Utils.GetResourceString($"{control.Id}_label"));
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            return WithActionDo<string>(control, (action, param) => action.GetSupertip(control.Id, param), "");
        }


        public object? GetImage(Office.IRibbonControl control)
        {
            var bitmap = WithActionDo<Bitmap?>(control, (action, param) => action.GetImage(control.Id, param), null);
            if (bitmap != null)
            {
                return bitmap;
            }
            var msoImage = WithActionDo<string?>(control, (action, param) => action.GetMsoImage(control.Id, param), null);
            return msoImage;
        }

        public void OnArrangeSettingsOpen(Office.IRibbonControl control)
        {
            var dialog = new Window()
            {
                Title = Strings.arrange_shape_settings_dialog_title,
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
            };
            var settings = new ArrangeShapeControl(dialog);
            dialog.Content = settings;
            var interop = new WindowInteropHelper(dialog);
            interop.Owner = Process.GetCurrentProcess().MainWindowHandle;
			if (dialog.ShowDialog() == true) {
				Properties.Settings.Default.Save();
			} else { 
				Properties.Settings.Default.Reload();
			}                            
        }


        #region IRibbonExtensibility Members

        public string? GetCustomUI(string ribbonID)
        {
            return GetResourceText("hztoolbar.ToolbarRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            var application = Globals.ThisAddIn.Application;
            application.WindowSelectionChange += (_) =>
            {
                this.ribbon.Invalidate();
            };
        }

        #endregion

        #region Helpers

        private static string? GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}
