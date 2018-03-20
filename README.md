# XlsMISP

XlsMISP is an Excel based tool to assist visualizing from and sharing data to a MISP instance.  XlsMISP is highly customizable to your needs and proxy friendly.

Its current version can:
- Provide you a customizable list of the events of the last X days published on a given MISP instance. The number of days to retrieve is configurable (tab "About").
- Retrieve the details of these events automatically and display them in a user configurable template (1 Excel tab per event).
- Create new MISP events based on templates (you can define several creation templates by cloning tab "Create new MISP" and customizing its layout).
- Add attributes to existing events, either directly if it's your own event or by submitting a proposal for events created by other organizations.
- View and accept pending proposals on your own events.
- Publish (or alert) updated events.

The current version is certainly not bug free and has the current known limitations:
- It's not yet possible to propose files to a MISP event which isn't yours
- Modifications of event properties or attributes are not possible (yet)
- Attribute category and type isn't yet a dropdown but just free text

Several other ideas or enhancements are pending and documented in tab "About".


## How to get started
- Download the XLSM file
- Configure your Excel to allow running the macro (consider signing the VBA code or add the folder where XlsMISP is located to the trusted path)
- Open XlsMISP and fill out your MISP URL, API key, as well as your MISP organization name in tab "About"
- Review the other settings in the same tab (typically the web settings for proxy and user agent parameters)
- Consider loading the tags of your MISP instance, by clicking button "Refresh tags" in tab "About"
- In tab "Overview", retrieve the list of current MISP events from the configured MISP instance


## FAQ

### Why XlsMISP?
The idea of XlsMISP is to bridge a bit the gap between the MISP web GUI and tools such as PyMisp. XlsMISP allows easy customization so that analysts can define exactly what they want to see and ease their contributions. Being a tool based on Excel means as well that there's no strong technical prerequisite (aside having Office) to be able to use the tool. Note that the VBA code should mostly work on MacOSX as well but hasn't been tested.


### So, XlsMISP is a XLSM with macro?
Indeed, sounds like fighting fire with fire! On the other hand, it provides a ground for easy customization and everyone has a bit of Excel knowledge. It's clear that running untrusted macros on your system isn't recommended. The idea behind XlsMISP is that you can (in the future, unless you want to be the first to test this experimental aspect) rebuild the tool by:
- Downloading yourself your trusted version of VBA-Web from Tim Hall, which is currently the only dependency and used for all web service calls
- Import the few XlsMISP VBA modules after having reviewed their code
- Create your own templates

Aside "binary" (XLSM) releases, we also provide the dump of the VBA content for each commit.


### Do you have tricks to create a new MISP event or add new attributes to an existing MISP event?
Sure! Don't worry if not all lines are complete or if there are a small numbers of blank lines in your attribute list, as XlsMISP will only submit to the server attribute lines which are complete. An attribute line is considered as complete if its "Attrib type" is not "Attribute" and the following attributes are not empty:
- category
- type
- to_ids

As you can see in the initial creation template (tab "Create new MISP"), you can leave blank lines in your template (number of maximal blank lines can be configured in tab "About" - setting "Max blank lines to parse for new attribs").


### The display layout is ugly as hell!
Probably - but almost everything is configurable as you just get a template shell and you can enhance it yourself!

You can choose the relevant event properties you can to see in tab "Overview" by simply adding or removing the column headers. The column header must either match a MISP event property name or be a pre-processing formula. Column names starting with xl. indicate that the MISP property in argument gets pre-processed by XlsMISP according to the pseudo-formula. For example xl.ts(timestamp) converts the Unix timestamp to an Excel compatible date while xl.listrec(EventTag.Tag) converts an array of tags to a string by extracting the Tag name.

Tab "Template" is just a template used when you retrieve MISP event details. You can customize most of the template, by 
- Deleting or adding event properties (e.g. event uuid)
- Adding, removing or rearranging columns for the event attributes (you need to keep columns "Attrib type", category, type and to_ids as they are mandatory, but their order shouldn't matter)

Tab "Create new MISP" is another template allowing the creation of new MISP events. Unlike tab "Template" which is unique and cannot be renamed, you can clone as many of those tabs and define the default properties and attribute names for different types of event creations.

Contributions to enhance the default templates / tabs are also welcome!


### Help! It doesn't work as expected / I cannot connect / there's an error somewhere
Consider turning on logging in the "About" tab and you will see the web request and responses written to the VBA immediate window (in the VBA Editor, press Ctrl+G or menu "View" - "Immediate Window"). This should allow troubleshooting most connectivity issues.
