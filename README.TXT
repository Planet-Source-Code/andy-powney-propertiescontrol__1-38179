
Name:    PropertiesControl.ctl
Version: 1.1
Date:    21st August 2002
Author:  Andy Powney, The Perplexity Project


Introduction

  There are many occasions where the VB developer needs to allow the user to modify the properties of something, much in the same
way that the developer modifies properties of visual objects during the development of a project. This VB UserControl enables the
developer to incorporate this functionality into their applications.

  This control is a wrapper for the MSFlexGrid control, so that must be installed on your development system, and must also be
included in your release bundle for your application. It provides a simple interface for the developer to manipulate the data held
within the control, and to access the data at any time.

  The functionality is not as complete as the properties sheets found in Developer Studio etc., but there is definitely enough for
the developer to achieve most results. Future versions may incorporate more advanced features, but this will be driven by demand.


Features

  This version of the control can manipulate TEXT, BOOLEAN, LIST and BUTTON properties.
  All properties displayed in the control are sorted according to their name.
  Addition, updating and removal of properties and their values can be done at any time.


Supported Property Types

  TEXT         This is the most basic of property types. When the user clicks on a named
               property, he is allowed to change the text value according to his wishes.
               This property can also handle numeric values.

  BOOLEAN      This property type will display a drop-down list containing the words "True"
               and "False", allowing the user to select the boolean value to assign to the
               named property.

  LIST         In much the same was as the boolean property type, the user is presented
               with a drop-down list, which is populated in advance by the developer. The
               user has no choice but to select a value from this list. A list is represented
               as a single string, where each item is separated with the vbCRLF character pair.

  BUTTON       The most common use for a button property type is for the selection of a file
               or a filename. The filename is displayed as text in the property sheet (though
               it is not editable directly by the user), and a button appears to the right.
               When the user clicks on the button, he is presented with a dialog box to select
               the filename.


Using The Control In Your Application

  1) Copy both PropertiesControl.ctl and PropertiesControl.ctx into your source code directory
  2) Select Add > UserControl
  3) Select the "Existing" tab, and click the PropertiesControl.ctl file


More Information

  This is freeware, no warranties etc. You are free to use or modify this code, but you can only
distribute modifications to the source code with my consent. Use it in any application you like,
and you don't have to mention it.


Bugs And Comments

  Either fix them yourself and report back to me, or tell me via email: andy.powney@gotomy.com
