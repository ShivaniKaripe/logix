<?xml version="1.0" encoding="utf-8"?>
<!-- version:7.3.1.138972.Official Build (SUSDAY10202) -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html"/>
<!-- This is the overall template for the page; subsequent templates for each section are called from within it -->
  <xsl:template match="CustWeb">
    <xsl:variable name="CustomerPK">
      <xsl:value-of select="@CustomerPK"/>
    </xsl:variable>
    <xsl:variable name="Mode">
      <xsl:value-of select="@Mode"/>
    </xsl:variable>
    <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
      <head>
        <title>Customer Website:
          <xsl:choose>
            <xsl:when test="$Mode='Home'">
              Home
            </xsl:when>
            <xsl:when test="$Mode='Stores'">
              Stores
            </xsl:when>
          </xsl:choose>
        </title>
        <meta http-equiv="cache-control" content="no-cache" />
        <meta http-equiv="pragma" content="no-cache" />
        <link rel="icon" href="images/favicon.ico" type="image/x-icon" />
        <link rel="shortcut icon" href="image/favicon.ico" type="image/x-icon" />
        <link rel="stylesheet" href="css/cw-handheld.css" type="text/css" media="handheld" />
        <link rel="stylesheet" href="css/cw-screen.css" type="text/css" media="screen" />
        <link rel="stylesheet" href="css/cw-print.css" type="text/css" media="braille, embossed, print, projection, tty" />
        <script src="javascript/cw.js" type="text/javascript">
          <xsl:comment> Comment inserted for Internet Explorer </xsl:comment>
        </script>
      </head>
      <body class="framed">
        <div id="wrap" style="border:0;padding:0;">
          <xsl:apply-templates select="Customer"/>
          <xsl:choose>
            <xsl:when test="$Mode='Home'">
              <xsl:apply-templates select="Offers"/>
              <xsl:apply-templates select="Groups"/>
            </xsl:when>
            <xsl:when test="$Mode='Stores'">
              <xsl:apply-templates select="Stores"/>
            </xsl:when>
          </xsl:choose>
        </div>
      </body>
    </html>
  </xsl:template>
  
  
  
  
  <!-- Below are the templates that generate each of the major sections, called from within the main template -->
  
  
  <!-- Customer section template -->
  <xsl:template match="Customer">
    <xsl:for-each select="../Customer">
      <div id="sidebar">
        <div id="sidebarinset">
          <h1 class="title">
            <span class="FirstName">
              <xsl:value-of select="FirstName"/>
            </span>
            <xsl:text>&#xA0;</xsl:text>
            <span class="LastName">
              <xsl:value-of select="LastName"/>
            </span>
          </h1>
          <p id="Contact">
            <span class="Address">
              <xsl:value-of select="Address"/>
            </span>
            <br />
            <span class="City">
              <xsl:value-of select="City"/>
            </span>
            <xsl:text>,&#xA0;</xsl:text>
            <span class="State">
              <xsl:value-of select="State"/>
            </span>
            <xsl:text>&#xA0;</xsl:text>
            <span class="Zip">
              <xsl:value-of select="Zip"/>
            </span>
            <br />
            <span class="Phone">
              <xsl:choose>
                <xsl:when test="Phone=''">
                  <i>No phone number on file.</i>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="Phone"/>
                </xsl:otherwise>
              </xsl:choose>
            </span>
            <br />
            <span class="Email">
              <xsl:choose>
                <xsl:when test="Email=''">
                  <i>No email address on file.</i>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="EmailAddress">
                    <xsl:value-of select="Email"/>
                  </xsl:variable>
                  <a href="mailto:{$EmailAddress}">
                    <xsl:value-of select="Email"/>
                  </a>
                </xsl:otherwise>
              </xsl:choose>
            </span>
          </p>
          <p id="Identifiers">
            Your customer ID is
            <span class="PrimaryExtID">
              <xsl:value-of select="PrimaryExtID"/>
            </span>
            <br />
            <xsl:choose>
              <xsl:when test="HHPK=0">
                Not in a household.
              </xsl:when>
              <xsl:otherwise>
                You are part of household
                <span class="">
                  <xsl:value-of select="HHPK"/>
                </span>.
              </xsl:otherwise>
            </xsl:choose>
            <br />
            <xsl:choose>
              <xsl:when test="Employee='true'">
                You are an employee.
              </xsl:when>
              <xsl:otherwise>
                Not an employee.
              </xsl:otherwise>
            </xsl:choose>
          </p>
          <p id="Savings">
            This year you've saved<br />
            <span class="CurrYearSTD">
              $<xsl:value-of select="CurrYearSTD"/>
            </span>.
            <br />
            Last year you saved<br />
            <span class="LastYearSTD">
              $<xsl:value-of select="LastYearSTD"/>
            </span>.
          </p>
          <xsl:if test="@Editable='true'">
            <form action="nhome.aspx?Edit=edit" id="editform" name="editform"  target="_top">
              <input id="editbutton" name="editbutton" type="submit" value="Edit" />
            </form>
            <br />
          </xsl:if>
          <hr />
          <p>Protect your privacy by logging out when done.</p>
          <form action="nhome.aspx?Logout=logout" id="logoutform" name="logoutform" target="_top">
            <input id="logoutbutton" name="editbutton" type="submit" value="Logout" />
          </form>
        </div>
      </div>
    </xsl:for-each>
  </xsl:template>
  
  
  <!-- Offers section template -->
  <xsl:template match="Offers">
    <div id="Offers">
      <h1 class="title">Offers</h1>
      <p>You currently have <xsl:value-of select="count(Offer)"/> offers:</p>
      <xsl:for-each select="Offer">
        <xsl:variable name="OfferID">
          <xsl:value-of select="@ID"/>
        </xsl:variable>
        <div class="Offer" id="Offer{$OfferID}">
          <span class="Name">
            <xsl:value-of select="Name"/>
          </span>
          <br />
          <span class="Description">
            <xsl:value-of select="Description"/>
          </span>
          <br />
          <xsl:choose>
            <xsl:when test="Points>0">
              <span class="Points">
                Points: <xsl:value-of select="Points"/>
              </span>
              <br />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="Accumulation>0">
              <span class="Accumulation">
                Accumulation: <xsl:value-of select="Accumulation"/>
              </span>
              <br />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="BodyText!=''">
              <span class="BodyText">
                <xsl:value-of select="BodyText"/>
              </span>
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="Graphic!=''">
              <xsl:variable name="ImgSrc">
                <xsl:value-of select="Graphic"/>
              </xsl:variable>
              <span class="Graphic">
                <a href="{$ImgSrc}" target="_blank">
                  <img src="{$ImgSrc}"/>
                </a>
              </span>
            </xsl:when>
          </xsl:choose>
          <span class="Valid">
            Valid from
            <span class="StartDate">
              <xsl:value-of select="StartDate"/>
            </span>
            to
            <span class="EndDate">
              <xsl:value-of select="EndDate"/>
            </span>
            <xsl:choose>
              <xsl:when test="EmployeesOnly='true'">
                This offer is for employees only.<br />
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="EmployeesExcluded='true'">
                Employees are excluded from the offer.<br />
              </xsl:when>
            </xsl:choose>
          </span>
          <xsl:choose>
            <xsl:when test="AllowOptOut='true'">
              <form action="#" name="OptOut">
                <input class="optout" type="submit" value="Opt out of this offer" />
              </form>
            </xsl:when>
          </xsl:choose>
        </div>
      </xsl:for-each>
    </div>
  </xsl:template>


  <!-- Groups section template -->
  <xsl:template match="Groups">
    <div id="Groups">
      <h1 class="title">Groups</h1>
      <p>You're eligible to join <xsl:value-of select="count(Offer)"/> groups:</p>
      <xsl:for-each select="Offer">
        <xsl:variable name="OfferID">
          <xsl:value-of select="@ID"/>
        </xsl:variable>
        <div class="Offer" id="Offer{$OfferID}">
          <span class="Name">
            <xsl:value-of select="Name"/>
          </span>
          <br />
          <span class="Description">
            <xsl:value-of select="Description"/>
          </span>
          <br />
          <xsl:choose>
            <xsl:when test="Points>0">
              <span class="Points">
                Points: <xsl:value-of select="Points"/>
              </span>
              <br />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="Accumulation>0">
              <span class="Accumulation">
                Accumulation: <xsl:value-of select="Accumulation"/>
              </span>
              <br />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="BodyText!=''">
              <span class="BodyText">
                <xsl:value-of select="BodyText"/>
              </span>
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="Graphic!=''">
              <xsl:variable name="ImgSrc">
                <xsl:value-of select="Graphic"/>
              </xsl:variable>
              <span class="Graphic">
                <a href="{$ImgSrc}" target="_blank">
                  <img src="{$ImgSrc}"/>
                </a>
              </span>
            </xsl:when>
          </xsl:choose>
          <span class="Valid">
            Valid from
            <span class="StartDate">
              <xsl:value-of select="StartDate"/>
            </span>
            to
            <span class="EndDate">
              <xsl:value-of select="EndDate"/>
            </span>
            <xsl:choose>
              <xsl:when test="EmployeesOnly='true'">
                This offer is for employees only.<br />
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="EmployeesExcluded='true'">
                Employees are excluded from this offer.<br />
              </xsl:when>
            </xsl:choose>
          </span>
          <form action="#" name="OptIn">
            <input class="optin" type="submit" value="Join this group!" />
          </form>
        </div>
      </xsl:for-each>
    </div>
  </xsl:template>


  <!-- Stores section template -->
  <xsl:template match="Stores">
    <div id="Stores">
      <h1 class="title">Stores</h1>
      <p>There are <xsl:value-of select="count(Store)"/> stores:</p>
      <xsl:for-each select="Store">
        <xsl:variable name="StoreID">
          <xsl:value-of select="@ID"/>
        </xsl:variable>
        <div class="Store" id="Store{$StoreID}">
          <span class="LocationName">
            <xsl:value-of select="LocationName"/>
          </span>
          <span class="ExtLocationCode">
            (<xsl:value-of select="ExtLocationCode"/>)
          </span>
          <br />
          <span class="Address">
            <xsl:value-of select="Address1"/>
          </span>
          <br />
          <xsl:choose>
            <xsl:when test="Address2!=''">
              <span class="Address">
                <xsl:value-of select="Address2"/>
              </span>
              <br />
            </xsl:when>
          </xsl:choose>
          <span class="City">
            <xsl:value-of select="City"/>
          </span>
          <xsl:text>,&#xA0;</xsl:text>
          <span class="State">
            <xsl:value-of select="State"/>
          </span>
          <xsl:text>&#xA0;</xsl:text>
          <span class="Zip">
            <xsl:value-of select="Zip"/>
          </span>
          <br />
          <span class="Phone">
            <xsl:value-of select="Phone"/>
          </span>
          <br />
          <span class="Map">
            <a href="http://maps.google.com" target="Google Maps">Map...</a>
          </span>
          <br />
        </div>
      </xsl:for-each>
    </div>
  </xsl:template>
  
  
  
  
</xsl:stylesheet>