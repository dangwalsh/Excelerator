namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Globalization;
    using System.Windows.Data;

    using Microsoft.Office.Interop.Excel;

    [ValueConversion(typeof(Range), typeof(string))]
    class WorksheetConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var range = (Range)value;
            return range?.Worksheet.Name;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
