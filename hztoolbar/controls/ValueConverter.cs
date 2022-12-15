#nullable enable

using System;
using System.Diagnostics;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows;
using System.Globalization;

namespace hztoolbar.controls {

	public static class LengthUnit {
		public enum Unit {
			PT,
			INCH,
			CM
		}

		public static double Increment(Unit unit) {
			switch (unit) {
				case Unit.PT:
					return 1.0;
				case Unit.INCH:
				case Unit.CM:
					return 0.1;
				default:
					return 1.0;
			}
		}

		public static double Factor(Unit unit) {
			return unit switch {
				Unit.PT => 1.0,
				Unit.INCH => 1.0 / 72.0,
				Unit.CM => 0.393701 / 72.0,
				_ => 1.0,
			};
		}


		public static string Symbol(Unit unit) {
			return unit switch {
				Unit.PT => "pt",
				Unit.INCH => "\"",
				Unit.CM => "cm",
				_ => "",
			};
		}
	}

	public class LengthUnitValueConverter : MarkupExtension, IMultiValueConverter {

		private double factor = 1.0;

		public override object ProvideValue(IServiceProvider serviceProvider) {
			return this;
		}

		public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
			if (values.Length == 2 && values[1] is LengthUnit.Unit unit) {
				this.factor = LengthUnit.Factor(unit);
			} else {
				this.factor = 1.0;
			}
			var value = System.Convert.ToDouble(values[0]);
			return System.Convert.ChangeType(value * this.factor, targetType);
		}

		public object[] ConvertBack(object value, Type[] targetType, object parameter, System.Globalization.CultureInfo culture) {
			return new object[] {
				System.Convert.ChangeType(System.Convert.ToDouble(value) / this.factor, targetType[0]),
				Activator.CreateInstance(targetType[1])
			};
		}
	}
}
