using Microsoft.Extensions.Logging;

namespace ExcelToCSV.Utilities;

internal static class LoggingUtility
{
    internal static ILogger<T> GetLogger<T>()
    {
        using ILoggerFactory factory = LoggerFactory.Create(
            builder =>
                builder
                    .AddFilter("Microsoft", LogLevel.Warning)
                    .AddFilter("System", LogLevel.Warning)
                    .AddSimpleConsole(
                        options =>
                        {
                            options.IncludeScopes = false;
                            options.SingleLine = true;
                            options.TimestampFormat = "HH:mm:ss ";
                        }
                    )
        ); ;

        return factory.CreateLogger<T>();
    }
}
