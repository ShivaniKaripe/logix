using System.Web;
using CMS.AMS;
using Microsoft.Practices.Unity;

///<summary>
/// Summary description for WebRequestServiceResolver
/// </summary>
public class WebRequestServiceResolver : ServiceResolver
{

  public WebRequestServiceResolver(IUnityContainer container)
    : base(container)
  {

  }

  public override T Resolve<T>()
  {
    return base.Resolve<T>(HttpContext.Current.Request.Url.Host);
  }

}
