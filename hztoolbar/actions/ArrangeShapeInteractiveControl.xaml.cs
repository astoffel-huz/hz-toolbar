#nullable enable

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace hztoolbar.actions {
	/// <summary>
	/// Interaction logic for ArrangeShapeInteractiveControl.xaml
	/// </summary>
	public partial class ArrangeShapeInteractiveControl : UserControl {
		private static void OnPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e) {
			if (sender is ArrangeShapeInteractiveControl control) {
				control.OnValueChange();
			}
		}

		private readonly Window window;

		public int HorizontalGutter {
			get => (int)GetValue(HorizontalGutterProperty);
			set => SetValue(HorizontalGutterProperty, value);
		}

		public static readonly DependencyProperty HorizontalGutterProperty =
			DependencyProperty.Register("HorizontalGutter", typeof(int), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

		public bool HorizontalResize {
			get => (bool)GetValue(HorizontalResizeProperty);
			set => SetValue(HorizontalResizeProperty, value);
		}

		public static readonly DependencyProperty HorizontalResizeProperty =
			DependencyProperty.Register("HorizontalResize", typeof(bool), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

		public int VerticalGutter {
			get => (int)GetValue(VerticalGutterProperty);
			set => SetValue(VerticalGutterProperty, value);
		}

		public static readonly DependencyProperty VerticalGutterProperty =
			DependencyProperty.Register("VerticalGutter", typeof(int), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

		public bool VerticalResize {
			get => (bool)GetValue(VerticalResizeProperty);
			set => SetValue(VerticalResizeProperty, value);
		}

		public static readonly DependencyProperty VerticalResizeProperty =
			DependencyProperty.Register("VerticalResize", typeof(bool), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

		public event EventHandler? ValueChanged;

		public ArrangeShapeInteractiveControl(Window window) {
			InitializeComponent();
			this.window = window;
		}

		protected virtual void OnValueChange() {
			ValueChanged?.Invoke(this, EventArgs.Empty);
		}

		public void OnCloseClickHandler(object sender, RoutedEventArgs e) {
			this.window.DialogResult = false;
		}

		public void OnSaveClickHandler(object sender, RoutedEventArgs e) {
			this.window.DialogResult = true;
		}
	}
}
