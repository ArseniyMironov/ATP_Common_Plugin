using ATP_Common_Plugin.Services;
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ATP_Common_Plugin.Views
{
    class LogLevelToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is LogLevel level)
            {
                switch (level)
                {
                    case LogLevel.Error:
                        return Brushes.Red;
                    case LogLevel.Warning:
                        return Brushes.Orange;
                    default:
                        return Brushes.Black;
                }
            }
            return Brushes.Black;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
