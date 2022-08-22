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
	/// Interaction logic for ArrangeShapeControl.xaml
	/// </summary>
	public partial class ArrangeShapeControl : UserControl {
		private readonly Window window;

		public ArrangeShapeControl(Window window) {
			InitializeComponent();
			this.window = window;
		}

		private void OnOkClickHandler(object sender, RoutedEventArgs e) {
			this.window.DialogResult = true;
		}

	}
}
