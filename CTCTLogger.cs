using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using CTCT;
using CTCT.Components.Contacts;
using CTCT.Components;
using System.Globalization;
using CTCT.Exceptions;
using System.Configuration;
using CTCT.Services;
using System.IO;

namespace SageCRMConnector
{
   
    
    static class CTCTLogger
    {
       

    public static void LogFile(string sExceptionName, string sEventName, string sControlName, int       nErrorLineNo, string sFormName)

        {

            StreamWriter log;

            string logFile = DateTime.Now.ToString("yyyyMMdd") + ".log";


            if (!File.Exists("c:\\ctctlog\\"+ logFile))

            {

                log = new StreamWriter("c:\\ctctlog\\" + logFile);

            }

            else

            {

                log = File.AppendText("c:\\ctctlog\\" + logFile);

            }

            // Write to the file:

            log.WriteLine("Data Time:" + DateTime.Now);

            log.WriteLine("Exception Name:" + sExceptionName);

            log.WriteLine("Event Name:" + sEventName);

            log.WriteLine("Control Name:" + sControlName);

            log.WriteLine("Error Line No.:" + nErrorLineNo);

            log.WriteLine("Form Name:" + sFormName);

            // Close the stream:

            log.Close();

        }


   
    }

    public static class ExceptionHelper
    {

        public static int LineNumber(this Exception e)
        {

            int linenum = 0;

            try
            {

                linenum = Convert.ToInt32(e.StackTrace.Substring(e.StackTrace.LastIndexOf(":line") + 5));

            }

            catch
            {

                //Stack trace is not available!

            }

            return linenum;

        }

    }

    }

    

