**This is fork of http://github.com/whut/AttachTo**

**I've added in my fork better support of attaching to IIS.**
If at the attachment moment there are more than one w3wp.exe processes alive 
(for example there are running WCF and ASP.NET applications as separate sites, having own application pools) 
there's a possibility, that debugger will attach to wrong IIS process.
So what I'm doing differenty from original project - I'm scanning solution projects and when find first Web project, take its IIS url and use it to find right application pool.
**As long as I didn't wished any user interface to have, it does not handle more complex cases (like several applications in site, multiple web projects in solution).**

Adds "Attach to IIS", "Attach to IIS Express" and "Attach to NUnit" commands to Tools menu.

Now you can start debugging web site hosted in local IIS server or failing NUnit test quicker than before:) 

AttachTo extension provides options to hide commands not relevant to you. You can also assign shortcut using Tools -> Options -> Keyboard.