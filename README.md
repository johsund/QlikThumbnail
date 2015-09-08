# QlikThumbnail
Generate thumbnails for your Qlik Sense apps on Qlik Sense Server

NOTE: Read the Guide document included in the project.

When building new Qlik Sense apps in the Hub the sheets are assigned generic icons. Creating a new thumbnail can be a bit cumbersome, especially with multiple sheets.
Qlik Thumbnail allows you to get a screenshot of the sheet contents and automatically generates thumbnails and assigns them to the sheet.

Installation:
Unzip entire package into a folder.

Configuration:
1.	Check that you can access the target Qlik Sense hub properly and have rights to open all apps and update the sheet thumbnails.
(You may need to update the security role that applies to your user or create a temporary rule granting you RootAdmin access).

2.	Due to Qlik Thumbnail launching multiple sessions to different sheets and apps it seems to hit a limit in Qlik Sense. Assign a Login Access Rule and set aside some tokens for extra sessions. 


Using Qlik Thumbnail:
After you have created or imported apps into your Qlik Sense site and published them to streams on the Hub you’ll notice that there will be lots of sheets with the default “No Thumbnail” logo.
 

Let’s launch Qlik Thumbnail. The first thing to set is the hostname. Use https:// and the fully qualified domain name (FQDN) as per the Qlik Sense site certificate. 
 
Click on Test Connection to ensure that the connection to the APIs works as per the below screenshot:
 
If the connection fails, you will be shown the error message. Common issues are around using another host path than what is specified in the certificate (i.e. using IP, localhost, hostname). Ensure you can access the hub with this FQDN without getting warned about the certificate not matching the URL.

Once connection has been verified, click on Fetch Apps to pull a list of available Apps.
 

Select the apps whose sheets you want to generate thumbnails for and click “Assign thumbnails”.
 
Qlik Thumbnail will now launch an IE instance in full screen kiosk mode and open each sheet one by one. It waits for 25 seconds to give a sheet time to render before snapping a screenshot.
The screenshot is then cropped, resized, fed into the default Content Repository and assigned as a thumbnail to the sheet.
NOTE: If you pop open any windows in front of the dashboard sheet – these will be captured as part of the thumbnail, so stay clear of interrupting the image creation.
The IE process is killed and a new one is launched for the next sheet until all images have been generated and assigned. 

After finishing you will be able to see the activities that have taken place in the Progress view. 
 
Now if you check your Hub, the thumbnails for the sheets should have been assigned and images available in the Content Library. For the App itself, choose an appropriate picture or select one of your sheet thumbnails to use.
