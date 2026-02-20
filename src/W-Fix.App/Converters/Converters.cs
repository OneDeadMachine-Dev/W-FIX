using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using WFix.Core.Models;

namespace WFix.App.Converters;

/// <summary>LogLevel → text color for the live log console.</summary>
[ValueConversion(typeof(LogLevel), typeof(SolidColorBrush))]
public class LogLevelToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is LogLevel level ? level switch
        {
            LogLevel.Success => new SolidColorBrush(Color.FromRgb(0x2E, 0xCC, 0x71)),
            LogLevel.Warning => new SolidColorBrush(Color.FromRgb(0xFF, 0xA5, 0x00)),
            LogLevel.Error   => new SolidColorBrush(Color.FromRgb(0xFF, 0x47, 0x57)),
            _                => new SolidColorBrush(Color.FromRgb(0xC9, 0xD1, 0xD9)),
        } : Brushes.White;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>PrinterStatus → status dot color.</summary>
[ValueConversion(typeof(PrinterStatus), typeof(SolidColorBrush))]
public class PrinterStatusToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is PrinterStatus status ? status switch
        {
            PrinterStatus.Ready    => new SolidColorBrush(Color.FromRgb(0x2E, 0xCC, 0x71)),
            PrinterStatus.Printing => new SolidColorBrush(Color.FromRgb(0x2D, 0x7F, 0xF9)),
            PrinterStatus.Offline  or
            PrinterStatus.NotAvailable => new SolidColorBrush(Color.FromRgb(0x6E, 0x76, 0x81)),
            PrinterStatus.Error    or
            PrinterStatus.UserIntervention => new SolidColorBrush(Color.FromRgb(0xFF, 0x47, 0x57)),
            PrinterStatus.PaperJam or
            PrinterStatus.PaperOut => new SolidColorBrush(Color.FromRgb(0xFF, 0xA5, 0x00)),
            PrinterStatus.TonerLow or
            PrinterStatus.NoToner  => new SolidColorBrush(Color.FromRgb(0xFF, 0xD7, 0x00)),
            _ => new SolidColorBrush(Color.FromRgb(0x6E, 0x76, 0x81)),
        } : Brushes.Gray;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>bool/null → Visibility (true/not-null=Visible).</summary>
[ValueConversion(typeof(object), typeof(Visibility))]
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b) return b ? Visibility.Visible : Visibility.Collapsed;
        if (value is int i) return i > 0 ? Visibility.Visible : Visibility.Collapsed;
        return value != null ? Visibility.Visible : Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>Inverse of BoolToVisibilityConverter (null/false → Visible).</summary>
[ValueConversion(typeof(object), typeof(Visibility))]
public class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b) return b ? Visibility.Collapsed : Visibility.Visible;
        return value == null ? Visibility.Visible : Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>Checks if bound value equals converter parameter — for RadioButton IsChecked binding.</summary>
public class IsEqualToBoolConverter : IValueConverter
{
    public static readonly IsEqualToBoolConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        Equals(value, parameter);

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is true ? parameter : Binding.DoNothing;
}

/// <summary>Inverts a bool — for disabling Apply button while fix runs.</summary>
public class InverseBoolConverter : IValueConverter
{
    public static readonly InverseBoolConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is bool b && !b;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is bool b && !b;
}
