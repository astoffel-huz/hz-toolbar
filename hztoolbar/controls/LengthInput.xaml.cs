using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

using System.Globalization;
using System;

namespace hztoolbar.controls {

	/// <summary>
	/// Interaction logic for LengthInputControl.xaml
	/// </summary>
	/// 

	public partial class LengthInputControl : UserControl {

		public float Value {
			get { return Math.Max((float)Minimum, (float)GetValue(ValueProperty)); }
			set { SetValue(ValueProperty, Math.Max(Minimum, value)); }
		}

		public static readonly DependencyProperty ValueProperty =
			DependencyProperty.Register(
				"Value", typeof(float),
				typeof(LengthInputControl),
				new PropertyMetadata(0.0f));

		public LengthUnit.Unit Unit {
			get { return (LengthUnit.Unit)GetValue(UnitProperty); }
			set { SetValue(UnitProperty, value); }
		}

		public static readonly DependencyProperty UnitProperty =
			DependencyProperty.Register(
				"Unit", typeof(LengthUnit.Unit),
				typeof(LengthInputControl),
				new PropertyMetadata(LengthUnit.Unit.INCH));

		public double Minimum {
			get { return (double)GetValue(MinimumProperty); }
			set { SetValue(MinimumProperty, value); }
		}

		public static readonly DependencyProperty MinimumProperty =
			DependencyProperty.Register(
				"Minimum", typeof(double),
				typeof(LengthInputControl),
				new PropertyMetadata(Double.MinValue));

		public string ValueFormat {
			get { return "0.00 " + LengthUnit.Symbol(this.Unit).Replace("\\", "\\\\").Replace("\"", "\\\""); }
		}

		public double Increment {
			get { return LengthUnit.Increment(this.Unit); }
		}

		public LengthInputControl() {
			InitializeComponent();
			if (Utils.IsMetricRegion()) {
				this.Unit = LengthUnit.Unit.CM;
			}
			this.inputControl.DataContext = this;
		}


	}
}
