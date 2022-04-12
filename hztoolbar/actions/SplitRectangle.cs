using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using System.Diagnostics;
using System;

namespace hztoolbar.actions
{
    public class SplitRectangle : ToolbarAction
    {
        public SplitRectangle() : base("split_rectangle") { }

        public override bool IsEnabled(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();

            if (shapes.Count != 1)
            {
                return false;
            }

            var shape = shapes[0];

            return shape.Type == Office.MsoShapeType.msoAutoShape
                && shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle;
        }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();

            if (shapes.Count != 1)
            {
                return false;
            }

            var shape = shapes[0];

            var dialog = new Window()
            {
                Title = Strings.split_rectangle_dialog_title,
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
            };
            var settings = new SplitRectangleControl(dialog)
            {
                NumColumns = 2,
                NumRows = 2,
                ColumnGutter = Properties.Settings.Default.arrange_horizontal_gutter,
                RowGutter = Properties.Settings.Default.arrange_vertical_gutter,
            };
            dialog.Content = settings;
            var interop = new WindowInteropHelper(dialog);
            interop.Owner = Process.GetCurrentProcess().MainWindowHandle;
            if (dialog.ShowDialog() == true)
            {
                if (settings.NumRows == 1 && settings.NumColumns == 1)
                {
                    return false;
                }
                var rowHeight = (shape.Height - (settings.NumRows - 1) * settings.RowGutter) / settings.NumRows;
                var columnWidth = (shape.Width - (settings.NumColumns - 1) * settings.ColumnGutter) / settings.NumColumns;
                if (rowHeight <= 0||columnWidth <= 0)
                {
                    return false;
                }
                var slide = Utils.GetActiveSlide();
                if (slide == null)
                {
                    return false;
                }
                var top = shape.Top;
                for (var row = 0; row < settings.NumRows; ++row)
                {
                    var left = shape.Left;
                    for (var col = 0; col < settings.NumColumns; ++col)
                    {
                        var cell = slide.Shapes.AddShape(shape.AutoShapeType, left, top, columnWidth, rowHeight);
                        shape.PickUp();
                        cell.Apply();
                        left += columnWidth + settings.ColumnGutter;
                    }
                    top += rowHeight + settings.RowGutter;
                }
                shape.Delete();
            }
            return false;
        }
    }
}
