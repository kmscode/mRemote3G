using System;
using log4net;
using System.Reflection;

public class logger
{
    public static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

    public logger()
	{
        
    }
}
