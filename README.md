# DT_split

This macro that I made for work splits my WorldWide Deal Tracker files into separate files, by countries. 

Any other countries or country arrays can be added and removed at any time with minimal code change by editing modules 4 and/or 5.

What also needs to be done is edit the first few lines of the first three modules, in order to help the macro find the correct WW file.

Working with MacOS, I encountered many many issues along the way (Thank you /u/ITFuture for the help!!).

Ultimately, I chose to have the macro save the new files to /Users/JohnDoe/Library/Containers/com.microsoft.Excel/Data/Documents/

On Windows, a couple of important improvements can be made:
1. VBA Scripting Dictionary for saving the contact info in a smarter way...
2. Being able to save wherever you'd like...

