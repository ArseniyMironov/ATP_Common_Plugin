using ATP_Common_Plugin.Views;
using Autodesk.Revit.DB;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Services
{
    class LoggerService : ILoggerService
    {
        private readonly ConcurrentQueue<LogEntry> _logEntries = new ConcurrentQueue<LogEntry>();
        private readonly object _eventLock = new object();
        private LoggerWindow _window;
        private readonly object _windowLock = new object();
        private bool _isWindowCreated = false;

        public event Action<LogEntry> LogAdded;
        public IEnumerable<LogEntry> GetLogEntries()
            => _logEntries.ToArray(); // Возвращаем копию для безопасности потоков

        public void ShowWindow()
        {
            lock (_windowLock)
            {
                try
                {
                    // Создаем окно только при первом вызове
                    if (_window == null)
                    {
                        _window = new LoggerWindow(this);
                        _window.Closed += (s, e) =>
                        {
                            // При закрытии просто обнуляем ссылку
                            _window = null;
                        };
                        _isWindowCreated = true;
                    }

                    // Всегда показываем окно при вызове ShowWindow()
                    _window.Show();
                    _window.Activate();

                    // Поднимаем окно на передний план
                    _window.Topmost = true;
                    _window.Topmost = false;
                    _window.Focus();
                }
                catch (Exception ex)
                {
                    LogError($"Failed to show logger window: {ex.Message}");
                }
            }
        }

        public bool IsWindowVisible
        {
            get
            {
                lock (_windowLock)
                {
                    return _window != null && _window.IsVisible;
                }
            }
        }

        public void HideWindow()
        {
            lock (_windowLock)
            {
                _window?.Hide();
            }
        }
        public void ClearLogs()
        {
            while (_logEntries.TryDequeue(out _)) { }
            LogAdded = null; // Очищаем подписчиков
        }

        private void RaiseLogAdded(LogEntry entry)
        {
            // Потокобезопасный вызов события
            lock (_eventLock)
            {
                LogAdded?.Invoke(entry);
            }
        }

        public void Log(LogLevel level, string message, string doc = null)
        {
            var entry = new LogEntry(level, message, doc);
            _logEntries.Enqueue(entry);
            RaiseLogAdded(entry);
        }

        public void LogError(string message, string doc = null)
            => Log(LogLevel.Error, message, doc);

        public void LogInfo(string message, string doc = null)
            => Log(LogLevel.Info, message, doc);

        public void LogWarning(string message, string doc = null)
            => Log(LogLevel.Info, message, doc);
    }
}
