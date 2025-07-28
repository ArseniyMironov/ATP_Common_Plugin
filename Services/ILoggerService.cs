using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Services
{
    /// <summary>
    /// Уровни логирования для сообщений
    /// </summary>
    public enum LogLevel
    {
        Info,       // Информационные сообщения
        Warning,    // Некритические проблемы
        Error       // Критические ошибки
    }

    public class LogEntry
    {
        public DateTime Timestamp { get; }
        public LogLevel Level { get; }
        public string Message { get; }
        public string DocumentName { get; }

        public LogEntry(LogLevel level, string message, string doc = null)
        {
            Timestamp = DateTime.Now;
            Level = level;
            Message = message ?? throw new ArgumentNullException(nameof(message));
            DocumentName = doc ?? "Global";
        }
    }

    /// <summary>
    /// Интерфейс сервиса логирования
    /// </summary>
    public interface ILoggerService
    {
        /// <summary>
        /// Основной метод логирования
        /// </summary>
        void Log(LogLevel level, string message, string doc = null);

        // Специализированные методы для удобства
        void LogInfo(string message, string doc = null);
        void LogWarning(string message, string doc = null);
        void LogError(string message, string doc = null);
        bool IsWindowVisible { get; }

        /// <summary>
        /// Управление окном логгера
        /// </summary>
        void ShowWindow();
        void HideWindow();
        void ClearLogs();
        /// <summary>
        /// Событие для уведомления о новых записях
        /// </summary>
        event Action<LogEntry> LogAdded;

        /// <summary>
        /// Получение всех записей (для привязки в UI)
        /// </summary>
        IEnumerable<LogEntry> GetLogEntries();
    }
}
