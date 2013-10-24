Powershell.MWA
==============

Powershell.MWA is a PowerShell module making use of the Microsoft.Web.Administration API.
The goal is to provide a quick and easy way to automate and configure IIS application pools and websites.

One of the main motivations behind this was for batch processing of new sites and application pools. I personally find the current IIS Provider a bit hard to use and not exactly the ideal candidate for performance. I have tested up to 1,000 sites and application pools, and this module finished in < 10s.  This was executing locally, not using any remote capabilities, however. While doing this remotely will have an impact on the performance, it will still be faster than the IIS Provider.

Requirements
============

IIS v7.0+. IIS Management Tools and Scripts Feature must be enabled on any remote servers. For example, if you want to interact remotely with IIS servers from your admin workstation, IIS 7 is required on the admin workstation and IIS Management Tools and Scripts Feature must be enabled on any remote servers.
