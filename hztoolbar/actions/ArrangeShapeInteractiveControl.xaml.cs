#nullable enable

using System;
using System.Windows;
using System.Windows.Controls;

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

		public float HorizontalGutter {
			get => Math.Max(0.0f, (float)GetValue(HorizontalGutterProperty));
			set => SetValue(HorizontalGutterProperty, value);
		}

		public static readonly DependencyProperty HorizontalGutterProperty =
			DependencyProperty.Register("HorizontalGutter", typeof(float), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

		public bool HorizontalResize {
			get => (bool)GetValue(HorizontalResizeProperty);
			set => SetValue(HorizontalResizeProperty, value);
		}

		public static readonly DependencyProperty HorizontalResizeProperty =
			DependencyProperty.Register("HorizontalResize", typeof(bool), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

		public float VerticalGutter {
			get => Math.Max(0.0f, (float)GetValue(VerticalGutterProperty));
			set => SetValue(VerticalGutterProperty, value);
		}

		public static readonly DependencyProperty VerticalGutterProperty =
			DependencyProperty.Register("VerticalGutter", typeof(float), typeof(ArrangeShapeInteractiveControl), new PropertyMetadata(OnPropertyChanged));

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
