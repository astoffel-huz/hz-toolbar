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

namespace hztoolbar.actions
{
    /// <summary>
    /// Interaction logic for ArrangeShapeControl.xaml
    /// </summary>
    public partial class ArrangeShapeControl : UserControl
    {
        private readonly Window window;

		public event EventHandler ValueChanged;

        public ArrangeShapeControl(Window window)
        {
            InitializeComponent();
            this.window = window;
        }

		protected virtual void OnValueChange() {
			ValueChanged?.Invoke(this, EventArgs.Empty);
		}

        private void OnOkClickHandler(object sender, RoutedEventArgs e)
        {
            this.window.DialogResult = true;
        }

		private void OnValueChangeHandler(object sender, HandyControl.Data.FunctionEventArgs<double> e) {
			OnValueChange();
		}
	}
}
