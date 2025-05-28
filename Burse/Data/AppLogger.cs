using System;

using Serilog;
using Serilog.Core;
using Serilog.Events;


namespace Burse.Data
{
    public class AppLogger
    {
        private readonly Serilog.ILogger _errorLogger;
        private readonly Serilog.ILogger _studentLogger;
        private readonly Serilog.ILogger _formatiiLogger;
        private readonly Serilog.ILogger _excelLogger;
        private readonly Serilog.ILogger _fisiereStudenti;

        public AppLogger()
        {
            _errorLogger = new LoggerConfiguration()
                .MinimumLevel.Error()
                .WriteTo.File("logs/errors-.txt", rollingInterval: RollingInterval.Day, shared: true)
                .CreateLogger();

            _studentLogger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File("logs/students-.txt", rollingInterval: RollingInterval.Day, shared: true)
                .CreateLogger();

            _formatiiLogger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File("logs/formatii-.txt", rollingInterval: RollingInterval.Day, shared: true)
                .CreateLogger();

            _excelLogger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File("logs/excel-import-.txt", rollingInterval: RollingInterval.Day, shared: true)
                .CreateLogger();

            _fisiereStudenti = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File("logs/students-excels-.txt", rollingInterval: RollingInterval.Day, shared: true)
                .CreateLogger();
        }

        public void LogError(string message, Exception ex = null) =>
            _errorLogger.Error(ex, message);

        public void LogStudentInfo(string message) =>
            _studentLogger.Information(message);

        public void LogFormatiiInfo(string message) =>
            _formatiiLogger.Information(message);

        public void LogExcelImport(string message) =>
            _excelLogger.Information(message);

        public void LogStudentsExcels(string message) =>
            _fisiereStudenti.Information(message);

    }

}
