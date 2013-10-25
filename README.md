Powershell.MWA
==============

Powershell.MWA is a PowerShell module making use of the Microsoft.Web.Administration API.
The goal is to provide a quick and easy way to automate and configure IIS application pools and websites.

One of the main motivations behind this was for batch processing of new sites and application pools. I personally find the current IIS Provider a bit hard to use and not exactly the ideal candidate for performance. I have tested up to 1,000 sites and application pools, and this module finished in < 10s.  This was executing locally, not using any remote capabilities, however. While doing this remotely will have an impact on the performance, it will still be faster than the IIS Provider.

The PSCustomObjects that are returned by Get-IisApplicationPool and Get-IisSite were designed based on what is useful for me in our environment.  We don't have any FTP sites, we use an SSL Proxy, and I haven't figured out how to deal with redirects beyond appcmd commands.  Feel free to extend the functionality for Certificate management, FTP, and redirects, as those are either non-existent issues for us or don't have enough value in figuring out beyond our current automation means at the moment.

For usuage information about the cmdlets, please use the built-in Get-Help system.  I have provided fairly basic help documentation that should be adequate in figuring out cmdlet usage.  Please use the issue tracker to report any bugs and I'll do my best to fix them.


Requirements
============

IIS v7.0+. IIS Management Tools and Scripts Feature must be enabled on any remote servers. For example, if you want to interact remotely with IIS servers from your admin workstation, IIS 7 is required on the admin workstation and IIS Management Tools and Scripts Feature must be enabled on any remote servers.

Available Cmdlets
=================

Name
----
Add-IisSiteBinding
Get-IisApplicationPool
Get-IisSite
New-IisApplicationPool
New-IisSite
Restart-IisApplicationPool
Set-IisApplicationPoolUser
Set-IisSiteCodePath
Start-IisSite
Stop-IisSite

Usage
=====

NAME
    Add-IisSiteBinding
    
SYNOPSIS
    Adds a new binding to a Microsoft IIS7.0+ website.
    
    -------------------------- EXAMPLE 1 --------------------------
    
    C:\PS>Add-IisSiteBinding -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -HostHeader www.contoso.com
    
    
 
    
    -------------------------- EXAMPLE 2 --------------------------
    
    C:\PS>Add-IisSiteBinding -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -IP 8.8.5.5 -Port 443 -Protocol https 
    -HostHeader www.contoso.com
    
    

NAME
    New-IisApplicationPool
    
SYNOPSIS
    Creates new IIS application pools.
    
    -------------------------- EXAMPLE 1 --------------------------
    
    C:\PS>Get-IisApplicationPool -ComputerName 8.8.8.8 | New-IisApplicationPool -Username "foo" -Password "bar"
    

    
    
    -------------------------- EXAMPLE 2 --------------------------
    
    C:\PS>New-IisApplicationPool -ComputerName 8.8.8.8, 8.8.4.4 -ApplicationPoolName contoso -AutoStart $true 
    -ManagedPipelineMode Integrated -ManagedRuntimeVersion "v4.0" -Enable32bit $false -IdentityType SpecificUser 
    -Username foo -Password bar
    
    
    

NAME
    New-IisSite
    
SYNOPSIS
    Creates new IIS websites.
    
    -------------------------- EXAMPLE 1 --------------------------
    
    C:\PS>Get-IisSite -ComputerName 8.8.8.8 | New-IisSite
    
    

    
    -------------------------- EXAMPLE 2 --------------------------
    
    C:\PS>New-IisSite -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -CodePath "D:\inetpub\wwwroot\code" 
    -ApplicationPoolName ContosoAppPool -Bindings "*:80:www.consoso.com", "*:443:www.contoso.com"
    
    
    


