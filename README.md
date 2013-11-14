OneToManyLabels
===============

This VBScript will allow you to label features with a one to many relationship in ArcMap.

Directions:

In order for this script to work, you must have two feature classes that when joined, have a one to many relationship. In my script these feature classes are ZONINGHISTORYPARCELS and VW_ZONINGHISTORYPARCELS. (I am joining the view,
VW_ZoningHistoryParcels to ZoningHistoryParcel. (This creates a one to many relationship since there are several polygons in the view that relate to the one record in ZoningHistoryParcel) You would paste this script into the feature class that has a Unique Value field. In my case, this is the ZoningHistoryParcel feature class, and the Unique Value field is the [PARCELID] field.

You also will have to modify the database connection to match your configuration.

You must change the [CASENUMCOMBINED] field to the field name you wish to display. This is the strInfo2 variable.

After you make the changes, paste the code into the feature class. Go to properties, then the label tab, click expression and then click the advanced box. Paste the code over anything that may already exist in the box.

I have included a screen shot that shows the end result. The format of the labels will depend on how you set up your labels appearance settings.

![Alt text](/example.gif)
