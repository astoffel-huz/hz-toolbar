﻿using System;
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
using hztoolbar;

namespace hztoolbar.actions
{
    /// <summary>
    /// Interaction logic for SplitRectangleControl.xaml
    /// </summary>
    public partial class SplitRectangleControl : UserControl
    {

        public int NumRows
        {
            get => (int)GetValue(NumRowsProperty);
            set => SetValue(NumRowsProperty, value);
        }

        public static readonly DependencyProperty NumRowsProperty =
            DependencyProperty.Register("NumRows", typeof(int), typeof(SplitRectangleControl));

        public int RowGutter
        {
            get => (int)GetValue(RowGutterProperty);
            set => SetValue(RowGutterProperty, value);
        }

        public static readonly DependencyProperty RowGutterProperty =
            DependencyProperty.Register("RowGutter", typeof(int), typeof(SplitRectangleControl));

        public int NumColumns
        {
            get => (int)GetValue(NumColumnsProperty);
            set => SetValue(NumColumnsProperty, value);
        }

        public static readonly DependencyProperty NumColumnsProperty =
            DependencyProperty.Register("NumColumns", typeof(int), typeof(SplitRectangleControl));

        public int ColumnGutter
        {
            get => (int)GetValue(ColumnGutterProperty);
            set => SetValue(ColumnGutterProperty, value);
        }

        public static readonly DependencyProperty ColumnGutterProperty =
            DependencyProperty.Register("ColumnGutter", typeof(int), typeof(SplitRectangleControl));

        private readonly Window window;

        public SplitRectangleControl(Window window)
        {
            this.window = window;
            InitializeComponent();
            this.DataContext = this;
        }

        private void onOkClick(object sender, RoutedEventArgs e)
        {
            this.window.DialogResult = true;
        }
    }
}
