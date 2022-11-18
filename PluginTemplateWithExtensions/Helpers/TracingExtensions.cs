using System;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace PluginTemplateWithExtensions.Helpers;

public static class TracingExtensions
{
    /// <summary>
    /// Writes out all the properties from an POCO object.
    /// </summary>
    /// <typeparam name="TValue"></typeparam>
    /// <param name="value"></param>
    /// <param name="trace"></param>
    /// <param name="headerMessage"></param>
    public static void LogValues<TValue>(this TValue value, ITracingService trace, string headerMessage = null)
    {
        var valueType = value.GetType();
        headerMessage ??= $"-----------  {valueType.Name} -------------";
        trace.Trace(headerMessage);
        foreach (var prop in valueType.GetProperties().OrderBy(property => property.Name))
        {
            trace.Trace($"{prop.Name}={prop.GetValue(value, null)}");
        }
    }
    /// <summary>
    /// This will add a time stamp to the trace along with the message
    /// </summary>
    /// <param name="tracingService"></param>
    /// <param name="message"></param>
    public static void Log(this ITracingService tracingService, string message)
    {
        tracingService?.Trace($"{DateTime.Now:hh:mm:ss.fff} {message}");
    }
}