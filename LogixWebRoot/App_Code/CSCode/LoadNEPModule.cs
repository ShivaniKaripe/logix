using Microsoft.Web.Infrastructure.DynamicModuleHelper;

[assembly: System.Web.PreApplicationStartMethod(typeof(Copient.LoadNEPModule), "LoadNEPModules")]
namespace Copient
{
    public class LoadNEPModule
  {
    public static void LoadNEPModules()
    {
      DynamicModuleUtility.RegisterModule(typeof(Ncr.Nep.Web.Security.BasicAuthenticationModule));
      DynamicModuleUtility.RegisterModule(typeof(Ncr.Nep.Web.Security.AccessTokenAuthenticationModule));
    }
  }
}