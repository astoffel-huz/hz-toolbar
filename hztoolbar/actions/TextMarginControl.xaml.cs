#nullable enable


using System;
using System.Collections.Generic;
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
	/// Interaction logic for UserControl1.xaml
	/// </summary>
	public partial class TextMarginControl : UserControl {

		private static void OnPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e) {
			if (sender is TextMarginControl control) {
				control.OnValueChange();
			}
		}

		private readonly Window window;

		public float TopMargin {
			get => Math.Max(0.0f, (float)GetValue(TopMarginProperty));
			set => SetValue(TopMarginProperty, value);
		}

		public static readonly DependencyProperty TopMarginProperty =
			DependencyProperty.Register("TopMargin", typeof(float), typeof(TextMarginControl), new PropertyMetadata(OnPropertyChanged));

		public float LeftMargin {
			get => Math.Max(0.0f, (float)GetValue(LeftMarginProperty));
			set => SetValue(LeftMarginProperty, value);
		}

		public static readonly DependencyProperty LeftMarginProperty =
			DependencyProperty.Register("LeftMargin", typeof(float), typeof(TextMarginControl), new PropertyMetadata(OnPropertyChanged));

		public float BottomMargin {
			get => Math.Max(0.0f, (float)GetValue(BottomMarginProperty));
			set => SetValue(BottomMarginProperty, value);
		}

		public static readonly DependencyProperty BottomMarginProperty =
			DependencyProperty.Register("BottomMargin", typeof(float), typeof(TextMarginControl), new PropertyMetadata(OnPropertyChanged));

		public float RightMargin {
			get => Math.Max(0.0f, (float)GetValue(RightMarginProperty));
			set => SetValue(RightMarginProperty, value);
		}

		public static readonly DependencyProperty RightMarginProperty =
			DependencyProperty.Register("RightMargin", typeof(float), typeof(TextMarginControl), new PropertyMetadata(OnPropertyChanged));

		public event EventHandler? ValueChanged;

		public TextMarginControl(Window window) {
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
