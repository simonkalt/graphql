Ban Man Pro Release Notes
--------------------------------------------------------------


Version 1.1.0
-------------
Added an option to view all expired campaigns.

Option to be emailed when a campaign expires.

Sizes of banners are now shown under campaigns and zones.  (Remember, all banners 
of the same size should be placed in any given campaign or zone).

Highly advanced ad serving code for the HTML mode.  This new code uses JavaScript 
and significantly decreases the memory requirements on the server.

Added reporting by Zones.

Removed most session variables to significantly cut down on the server resources 
utilized by the application. 

Fixed a few bugs in the reporting options. 


Version 1.1.3
-------------
Optimized queries to only pull required parameters from database to enhance 
performance.


Version 1.1.4
-------------
Fixed a bug whereby if a campaign expired and a default was not included in that 
zone nothing was returned to the browser and this would crash Netscaped 4.X.


Version 1.1.5
-------------
Added Randomization option for third party ad code.  The user simply inserts 
[RandomNumber] including the brackets in any location where a random number 
is to be placed.


	'Replace [RandomNumber] with a random number
	If InStr(strBMPAdCode,"[RandomNumber]") >0 Then
		Randomize
		lngRandom2=Int(Rnd*100000)
		strBMPAdCode=Replace(strBMPAdCode,"[RandomNumber]",lngRandom2)
	End If

Version 1.1.6
-------------
Altered Netscape 3 ad serving code to include border=0 and target="_new".  
This affects HTML code only.

Version 1.1.7
-------------
Added Executive Report.

Version 1.1.8
-------------
Added Chart to Reports which summarizes impressions in the past 7 days.

Version 1.1.9
-------------
Updated internal query which validated campaigns.  Program was serving one 
less than the quantity that was sold.  Program now serves the full quantity.

Version 1.2.0
-------------
Added option to view banners by advertiser.

Version 1.2.1
-------------
Added Error Trap to prevent user from adding two campaigns with the same name.

Version 1.2.2--4/14/2000
-------------
Added full support for BurstMedia and FlyCast code.  Changes took placed on the 
"Advanced Banner" screen.  Originally this code only worked in IE but now also 
supports Netscape.

Version 1.2.3--4/16/2000
-------------
Added parameter called [BanManProURL] which can be used for rich media ads and 
other ads added through the advanced code option.  Ban Man Pro will dynamically 
place this with the target URL allowing you to track clicks on rich media ads.

Version 1.2.4--4/19/2000
-------------
Expanded [RandomNumber] code to also work with regular banners in Addition to 
the Advanced banners.  This required changes to the getcode function in 
banman.asp as well as this code in the Case "Click" section of banman.asp...

		'Replace Random Number
		If InStr(strBMPTargetURL, "[RandomNumber]") >0 And Request.QueryString("Mode")<>"HTML" Then
			strBMPTargetURL = Replace(strBMPTargetURL,"[RandomNumber]", Request.QueryString("RandomNumber"))
		End If

Version 1.2.5--4/20/2000
-------------
Added a additional types of ad serving code.  Also added a parameter called 
PageID to help defeat the cache in non-javascript browsers.

View the following support article for more information on the different 
types of ad code.
http://www.banmanpro.com/support/codesummary.asp

Added a Delete confirmation to Banners, Campaigns and Zones.

Added Trim function to to the FIXBLANK function in addadvertiser.asp, 
addbanner.asp, addcampaign.asp and addzone.asp to prevent a space from 
appearing in front of some fields when updating.


Version 1.2.6--4/27/2000
-------------
Added a Billing Summary report.  For this function to work properly you 
need to enter a cost for your campaigns.
Files affected: detreports.asp, createreport.asp

Fixed an important bug relating to the newly added IFRAME/JavaScript ad 
code.  This bug counted two impressions for Internet Explorer browsers.  
This only affects people who have begun using the newly added Non-Cache 
Defeating Rich Media code.

Version 1.2.7--5/22/2000
-------------
Mastered International Dates for the SQL version.  We now use the SQL 
Convert function so that dates are always sent in the same format 
and then converted internally to match the regional date settings.

Version 2.0--6/15/2000
-------------
1) Added a smoothing algorithm.  The user specifies the number of 
minutes to base the smoothing algorithm on and then Ban Man Pro
continually updates the percentages.  For sites with a steady stream
of traffic, low values on the order of 5-15 minutes work great.  For
low traffic sites a larger value may be necessary such as 30-60 minutes.
This value is set in the Preferences.

2) Added a stop feature to completely stop serving ads in the event of
a database or some other failure.  This is useful in the situation that
the database server is temporarily down.

3) Added an option to select which reports your advertisers will see. This
information is set on the preferences screen.

4) Login information can now be stored in a cookie.

5) Added a logout button.  

6) Added an option to purge the database of old statistics.  When

7) Added an option to email campaigns which are approaching expiration.
This is in addition to the notification that is sent when one expires.
A notice is sent when a campaign is within 95% of expiration based on 
impressions.

8) Added a date setting option to support both US and international dates.

9) Added an email test option so you can test which email component your
server supports.

10) Added an option to determine how often multiple clicks from the same
user are counted.  For example, setting the value to 1 requires that
clicks be unique each hour.

11) Converted all queries in main script to stored procedures.  This
provides approximately a 15% performance gain.  [SQL Only]

12) Added an option to export all reports to Excel in addition to 
viewing them in a browser.

13) Introduced a multi-site version for managing multiple sites.

14) Added option to view both Banners and Campaigns by advertiser.  Sites
with a large number of banners and campaigns complained because it took
too long to load all banners.  When first visiting the banners screen
only the first 10 are loaded by default.  Also, on the advertisers screen,
advertisers can now be viewed by clicking on the appropriate letter.

15) Added the ability to call Zones by name.

16) Converted the main script into functions and now ASP users can call
banners by function.

17) Added Error trap for invalid end dates.

18) Added support for static text links.

19) Added support for multiple size defaults.  For this feature to work
you must specify a zone size for all your zones.  Default banners can
then be pulled based on matching sizes.

19) Added Keywords to campaigns so that they can be called by keyword.

20) Added email reports option for both advertisers and administrators.
Reports are available both daily and weekly.

21) Updated documentation

23) Added a slot option.


Version 2.01--6/25/2000
-------------
Note: 2.0--->2.01 Multisite users must run storedprocedures.sql to upgrade.
******Multi-Site Version Only******************************************
1) List of sites is now sorted alphabetically
2) Added "Run of Network" option so advertisers/banners/campaigns can 
be used across all sites
3) Added two additional reports for reporting across all sites
4) Added option to call site by site name (Function call only)

******Both Versions****************************************************
1) When viewing by advertiser, the advertiser now remains selected in
the drop-down list.
2) Added the company name in addition to the campaign name when 
creating/editing a zone.
3) Added Campaign Expiration report which displays all campaigns that
expire between the selected dates.  Note that campaigns can always
expire sooner if they are CPM or Perclick and are being distributed
based on weightings within a zone.
4) When visiting the reports screen, stats are now only displayed for
active campaigns.

Version 2.03--7/28/2000
-------------
1) Altered banmanfunc.asp to prevent the case where the server went down 
while the even algorithm was being computed.  If this occurred the even
algorithm would not be executed again.