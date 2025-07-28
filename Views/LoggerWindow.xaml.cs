using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;
using ATP_Common_Plugin.Services;

namespace ATP_Common_Plugin.Views
{
    /// <summary>
    /// Логика взаимодействия для LoggerWindow.xaml
    /// </summary>
    public partial class LoggerWindow : Window
    {
        private readonly ILoggerService _loggerService;
        private readonly ObservableCollection<LogEntry> _logEntries = new ObservableCollection<LogEntry>();
        public static bool IsWindowOpen { get; private set; }

        public LoggerWindow(ILoggerService loggerService)
        {
            InitializeComponent();
            // Начальное состояние - скрыто
            Visibility = Visibility.Collapsed;
            _loggerService = loggerService;

            LogListView.ItemsSource = _logEntries;

            LoadExistingLogs();
            _loggerService.LogAdded += OnLogAdded;
        }

        private void LoadExistingLogs()
        {
            foreach (var entry in _loggerService.GetLogEntries())
            {
                _logEntries.Add(entry);
            }
        }

        private void OnLogAdded(LogEntry entry)
        {
            // Обновляем UI в потоке UI
            Dispatcher.Invoke(() =>
            {
                _logEntries.Add(entry);

                // Автопрокрутка к последнему элементу
                if (LogListView.Items.Count > 0)
                {
                    LogListView.ScrollIntoView(LogListView.Items[LogListView.Items.Count - 1]);
                }
            });
            _loggerService.ShowWindow();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            _logEntries.Clear();
            _loggerService.ClearLogs(); // Реализуем позже
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Hide(); // Скрываем окно вместо закрытия
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _loggerService.LogAdded -= OnLogAdded;
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            e.Cancel = true; // Отменяем реальное закрытие
            Hide(); // Скрываем окно
            base.OnClosing(e);

            base.OnClosing(e);
        }
        public new void Show()
        {
            base.Show();
            Visibility = Visibility.Visible;
            Activate();
        }

        public new void Hide()
        {
            Visibility = Visibility.Collapsed;
        }
    }
}