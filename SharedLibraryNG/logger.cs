using log4net;
using log4net.Repository;
using log4net.Appender;
using System.IO;
using System.Windows.Forms;

public class logger
{
    public static readonly ILog Log = LogManager.GetLogger(typeof(logger));

    public logger()
	{
       CreateLogger();
    }


    public static void CreateLogger()
    {
        if(logger.Log.Logger.Name.ToLower().Equals("logger"))
        {
            // We did this already, so return.
            return;
        }

        log4net.Config.XmlConfigurator.Configure();

        string logFilePath = null;
#if !PORTABLE
	    logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Application.ProductName);
#else
        logFilePath = Application.StartupPath;
#endif
        string logFileName = Path.ChangeExtension(Application.ProductName, ".log");
        string logFile = Path.Combine(logFilePath, logFileName);

        ILoggerRepository repository = LogManager.GetRepository();
        IAppender[] appenders = repository.GetAppenders();
        FileAppender fileAppender = default(FileAppender);
        foreach (IAppender appender in appenders)
        {
            fileAppender = appender as FileAppender;
            if (!(fileAppender == null || !(fileAppender.Name == "LogFileAppender")))
            {
                fileAppender.File = logFile;
                fileAppender.ActivateOptions();
            }
        }
    }
}
