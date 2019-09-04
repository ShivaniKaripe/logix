using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CMS.AMS.Models;
using CMS.AMS.Contract;
using CMS.Contract;
namespace CMS.AMS
{
  /// <summary>
  /// Summary description for AMSResultExtension
  /// </summary>
  /// 
  public static class AMSResultExtension
  {
    private static IPhraseLib phraseLib;
    private static ILogger logger;
    private static IErrorHandler errorHandler;
    //public static AMSResult<P> Convert<T,P>(this AMSResult<T> source) where P : T
    //{
    //  AMSResult<P> result = new AMSResult<P>();
    //  result.MessageString = source.MessageString;
    //  result.ParameterList = source.ParameterList;
    //  result.PhraseString = source.PhraseString;
    //  result.ResultType = source.ResultType;
    //  result.Result = source.Result as P;
    //}
    public static string GetLocalizedMessage<T>(this AMSResult<T> source, int LanguageID )
    {
      ResolveDependencies();
      // default error message
      string localizedError = phraseLib.Lookup("term.UnknownError", LanguageID);
      
      if (source.ResultType != AMSResultType.Success  &&
          !string.IsNullOrEmpty(source.PhraseString))
      {
       

        // TODO eventually move this into a custom resource manager
        // See if there is a terminal specific resource
        string resourceString = phraseLib.Lookup(source.PhraseString, LanguageID);
       

        // If the resource key isn't present in the resources then it returns an empty string
        if (!string.IsNullOrEmpty(resourceString))
        {
          try
          {
            if (source.ParameterList == null || source.ParameterList.Count() == 0)
            {
              localizedError = resourceString;
            }
            else
            {
              localizedError = string.Format(resourceString, source.ParameterList.ToArray());
            }
          }
          catch (FormatException ex)
          {
            logger.WriteCritical(string.Format("{0} - {1} - {2}",
                System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName,
                System.Reflection.MethodBase.GetCurrentMethod().Name, ex));
          }
        }
        else
        {
          logger.WriteError(string.Format("Unable to find resource named '{0}' in Phrase Lib", source.PhraseString));
        }
      }
      string err = LogMessage<T>(source);
      if (!string.IsNullOrEmpty(err))
        localizedError = err;
      return localizedError;
    }

	
    private static void ResolveDependencies()
    {
      phraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
      logger = CurrentRequest.Resolver.Resolve<ILogger>();
      errorHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
    }

	
    private  static string  LogMessage<T>(this AMSResult<T> source)
    {
      string strerror = string.Empty;
      switch (source.ResultType)
      {

        case AMSResultType.Warning:
          logger.WriteWarn(source.MessageString);
          break;
        case AMSResultType.Unknown:
          logger.WriteError(source.MessageString);
          break;
        case AMSResultType.ValidationError:
          logger.WriteError(source.MessageString);
          break;
        case AMSResultType.Exception:
          strerror=errorHandler.ProcessError(source.MessageString);
          break;
        case AMSResultType.SQLException:
          strerror=errorHandler.ProcessError(source.MessageString);
          break;
      }
      return strerror;
    }
  }

}