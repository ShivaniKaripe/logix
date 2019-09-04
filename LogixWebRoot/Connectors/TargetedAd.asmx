<%@ WebService Language="VB" Class="Service" %>

Imports System.Web.Services
Imports System.Web.Services.Protocols

<WebService(Namespace:="http://www.copienttech.com/TargetedAd/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService
  ' version:5.99.1.39532.Official Build (CARBON)
  
  <WebMethod()> _
  Public Function AcceptCouponOffer(ByVal CardID As String, ByVal CardTypeID As Integer, ByVal ExtLocationCode As String, ByVal OfferID As Integer) As String
    'Redeems the coupon offer presented to the member.
    Dim se As New SoapException("This web service is no longer used. The methods in this web service are now found in the Kiosk Ad connector.", SoapException.ClientFaultCode, Context.Request.Url.AbsoluteUri)
    Throw se
  End Function
  
  <WebMethod()> _
  Public Function GetOfferCoupon(ByVal TargetedOffer As Integer) As String
    'Provides the coupon information needed to present a member with a targeted coupon.
    Dim se As New SoapException("This web service is no longer used. The methods in this web service are now found in the Kiosk Ad connector.", SoapException.ClientFaultCode, Context.Request.Url.AbsoluteUri)
    Throw se
  End Function
  
  <WebMethod()> _
  Public Function GetMemberOffers(ByVal CardID As String, ByVal CardTypeID As Integer) As String
    Dim se As New SoapException("This web service is no longer used. The methods in this web service are now found in the Kiosk Ad connector.", SoapException.ClientFaultCode, Context.Request.Url.AbsoluteUri)
    Throw se
  End Function

  <WebMethod()> _
  Public Function DeclineCouponOffer(ByVal CardID As String, ByVal CardTypeID As Integer, ByVal OfferID As Integer) As String
    'Adds the customer to an excluded group for the specified offer.
    Dim se As New SoapException("This web service is no longer used. The methods in this web service are now found in the Kiosk Ad connector.", SoapException.ClientFaultCode, Context.Request.Url.AbsoluteUri)
    Throw se
  End Function
  
End Class