using Microsoft.Practices.Unity;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Security;
using System.Net;
using System.Net.Http;

/// <summary>
/// Summary description for WebRequestResolverBuilder
/// </summary>
public class WebRequestResolverBuilder : ResolverBuilder
{

  //
  // TODO: Add constructor logic here
  //
  public WebRequestResolverBuilder()
  {
  }

  public override IServiceResolver GetResolver()
  {
   return  new WebRequestServiceResolver(ChildContainer);
  }

  protected override void SetUpRegisters()
  {
    base.SetUpRegisters();
   
    //Register specific dependency
    Container.RegisterType<IOffer, OfferService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICustomerGroups,CustomerGroups>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICustomerGroupCondition, CustomerGroupConditionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<Copient.StoredValue>(new HierarchicalLifetimeManager(), new InjectionConstructor());
    Container.RegisterType<Copient.Points>(new HierarchicalLifetimeManager(), new InjectionConstructor());
    Container.RegisterType<Copient.ImportXMLUE>(new HierarchicalLifetimeManager(), new InjectionConstructor());
    Container.RegisterType<Copient.ExportXmlUE>(new HierarchicalLifetimeManager(), new InjectionConstructor());
    Container.RegisterType<IAppMenuService, AppMenuService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IPointsCondition, PointsConditionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IStoredValueCondition, StoredValueConditionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IStoredValueProgramService, StoredValueProgramService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IPointsProgramService, PointsProgramService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IPassThroughRewards, PassThroughRewardService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ITrackableCouponProgramService, TrackableCouponProgramService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ITrackableCouponService, TrackableCouponProgramService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IActivityLogService, ActivityLogService>(new HierarchicalLifetimeManager());
    Container.RegisterType<INotesService, NotesService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICustomerService, CustomerService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICardTypeService, CardTypeService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ILocationsService, LocationsService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ITrackableCouponConditionService, TrackableCouponConditionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IGiftCardRewardService, GiftCardRewardService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IPreferenceRewardService, PreferenceRewardService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IProximityMessageRewardService, ProximityMessageRewardService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IAdminUserData, AdminUserDataService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IBuyerRoleData, BuyerRoleDataService>(new HierarchicalLifetimeManager());    
    Container.RegisterType<IImportValidator, OfferValidator>(new HierarchicalLifetimeManager());
    Container.RegisterType<ILocalizationService, LocalizationService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IProductService, ProductService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IProductGroupService, ProductGroupService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IAttributeService, AttributeService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IMessagingService, MessagingService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ISystemSettings, SystemSettings>(new HierarchicalLifetimeManager());
    Container.RegisterType<IPreferenceService, PreferenceService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IOfferDeploymentValidator, OfferDeploymentValidator>(new HierarchicalLifetimeManager());
    Container.RegisterType<IDepartment, DepartmentService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IHealth, Health>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICollisionDetectionService, CollisionDetectionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IInstantWinConditionService, InstantWinConditionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IXssEncoding, XssEncoding>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICouponRewardService, CouponRewardService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IProductConditionService, ProductConditionService>(new HierarchicalLifetimeManager());
    Container.RegisterType<ICouponPatternService, CouponPatternService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IDiscountRewardService, DiscountRewardService>(new HierarchicalLifetimeManager());
    Container.RegisterType<IAnalyticsCustomerGroups, AnalyticsCustomerGroups>(new HierarchicalLifetimeManager());
    Container.RegisterType<IOTPHelper, OTPHelper>(new HierarchicalLifetimeManager());
    Container.RegisterType<ILoginWithNEP, LoginWithNEP>(new HierarchicalLifetimeManager());
	    Container.RegisterInstance(typeof(WebClient), new WebClient());
    Container.RegisterType<IRestServiceHelper, RESTServiceHelper>(new HierarchicalLifetimeManager());
	    Container.RegisterType<IOfferApprovalWorkflowService, OfferApprovalWorkflowService>(new HierarchicalLifetimeManager());
        Container.RegisterType<INotificationService, NotificationService>(new HierarchicalLifetimeManager());
  }
}
