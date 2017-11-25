using System;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Repository.Hierarchy;

namespace toolkit.excel.data
{
    public static class LogFactory
    {
        public const string Log4NetConfig = "Log4Net.config";

        public static ILog GetLogger()
        {
            var hierarchy = LogManager.GetRepository() as Hierarchy;

            var uri = new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase),
                Log4NetConfig));
            var configFile = new FileInfo(Path.GetFullPath(uri.LocalPath));
            XmlConfigurator.ConfigureAndWatch(configFile);

            if (hierarchy != null)
            {
                var adoNetAppenders = hierarchy.GetAppenders().OfType<AdoNetAppender>();
                foreach (var adoNetAppender in adoNetAppenders)
                {
                    adoNetAppender.ConnectionString =
                        ConfigurationManager.AppSettings["ExcelDataContextConnectionString"];
                    adoNetAppender.ActivateOptions();
                }
            }
            var log = LogManager.GetLogger(typeof(LogFactory));
            return log;
        }
    }

    public class Log
    {
        [Key]
        public long LogId { get; set; }

        public DateTime Date { get; set; }
        public string Thread { get; set; }
        public string Level { get; set; }
        public string Logger { get; set; }
        public string Message { get; set; }
        public string Exception { get; set; }
    }
}