[System.Reflection.Assembly]::LoadFrom("$($Env:windir)\system32\inetsrv\Microsoft.Web.Administration.dll") | Out-Null

<#
.Synopsis
   Gets information about Microsoft IIS7.0+ website configuration.
.DESCRIPTION
   Returns a PSCustomObject containg informaton about the configuration of IIS websites.
.EXAMPLE
   Get-IisSite -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso
.EXAMPLE
   Get-IisSite -ComputerName 8.8.8.8, 8.8.4.4
.NOTES
#>
function Get-IisSite
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$false,
                   Position=1,
                   ParameterSetName="SingleSite")]
        [string]$SiteName
    )

    Begin
    {
        Write-Verbose "Connecting to $ComputerName"
        $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$ComputerName")
    }
    Process
    {
        if ($SiteName)
        {
            try
            {
                Write-Verbose "Getting site $SiteName"
                $Site = $ServerManager.Sites[$SiteName]
                $SiteObject = New-Object PSObject -Property @{

                    Name = $null
                    Id   = $null
                    ApplicationPool = $null
                    Bindings = @()
                    VirtualDirectories = $null

                }

                $SiteObject.Name = $Site.Name
                $SiteObject.Id = $Site.Id
                $SiteObject.ApplicationPool = $Site.Applications["/"].ApplicationPoolName
                $SiteObject.Bindings = $Site.Bindings | Select -ExpandProperty BindingInformation
                $SiteObject.VirtualDirectories = $Site.Applications.VirtualDirectories | Select Path, PhysicalPath

                Write-Output $SiteObject
            }
            
            catch
            {
                Write-Warning -Message $Error[0]
            }
        }

        else
        {
            try
            {
                $Sites = $ServerManager.Sites
                ForEach ($Site in $Sites)
                {
                    Write-Verbose "Getting site $site"
                    $SiteObject = New-Object PSObject -Property @{

                        Name = $null
                        Id   = $null
                        ApplicationPool = $null
                        Bindings = @()
                        VirtualDirectories = $null

                    }

                    $SiteObject.Name = $Site.Name
                    $SiteObject.Id = $Site.Id
                    $SiteObject.ApplicationPool = $Site.Applications["/"].ApplicationPoolName
                    $SiteObject.Bindings = $Site.Bindings | Select -ExpandProperty BindingInformation
                    $SiteObject.VirtualDirectories = $Site.Applications.VirtualDirectories | Select Path, PhysicalPath

                    Write-Output $SiteObject
                }
            }

            catch
            {
                Write-Warning -Message $Error[0]
            }
        }
    }

    End
    {
        $ServerManager.Dispose()
    }
}

<#
.Synopsis
   Gets information about Microsoft IIS7.0+ application pool configuration.
.DESCRIPTION
   Returns a PSCustomObject containg informaton about the configuration of IIS application pools.
.EXAMPLE
   Get-IisApplicationPool -ComputerName 8.8.8.8, 8.8.4.4 -ApplicationPoolName contoso
.EXAMPLE
   Get-IisSite -ComputerName 8.8.8.8, 8.8.4.4
.NOTES
#>
function Get-IisApplicationPool
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$false,
                   Position=1,
                   ParameterSetName="singleApplicationPool")]
        [string]$ApplicationPoolName
    )

    Begin
    {
        Write-Verbose "Connecting to $ComputerName"
        $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$ComputerName")
    }
    Process
    {
        if ($ApplicationPoolName)
        {
            try
            {
                Write-Verbose "Getting application pool $ApplicationPoolName"
                $ApplicationPool = $ServerManager.ApplicationPools[$ApplicationPoolName]
                $ApplicationPoolObject = New-Object PSObject -Property @{

                    Name = $null
                    Version   = $null
                    PipelineMode = $null
                    Enable32Bit = $null
                    AutoStart = $null
                    IdentityType = $null

                }

                $ApplicationPoolObject.Name = $ApplicationPool.Name
                $ApplicationPoolObject.Version = $ApplicationPool.ManagedRuntimeVersion
                $ApplicationPoolObject.PipelineMode = $ApplicationPool.ManagedPipelineMode
                $ApplicationPoolObject.Enable32Bit = $ApplicationPool.Enable32BitAppOnWin64
                $ApplicationPoolObject.AutoStart = $ApplicationPool.AutoStart
                $ApplicationPoolObject.IdentityType = $ApplicationPool.ProcessModel.IdentityType

                Write-Output $ApplicationPoolObject
            }

            catch
            {

            }
        }

        else
        {
            try
            {
                $ApplicationPools = $ServerManager.ApplicationPools
                ForEach ($ApplicationPool in $ApplicationPools)
                {
                    $ApplicationPool = $ServerManager.ApplicationPools["$($ApplicationPool.Name)"]
                    $ApplicationPoolObject = New-Object PSObject -Property @{

                        Name = $null
                        Version   = $null
                        PipelineMode = $null
                        Enable32Bit = $null
                        AutoStart = $null

                    }

                    $ApplicationPoolObject.Name = $ApplicationPool.Name
                    $ApplicationPoolObject.Version = $ApplicationPool.ManagedRuntimeVersion
                    $ApplicationPoolObject.PipelineMode = $ApplicationPool.ManagedPipelineMode
                    $ApplicationPoolObject.Enable32Bit = $ApplicationPool.Enable32BitAppOnWin64
                    $ApplicationPoolObject.AutoStart = $ApplicationPool.AutoStart

                    Write-Output $ApplicationPoolObject
                }
            }

            catch
            {
                Write-Warning -Message $Error[0]
            }
        }
    }

    End
    {
        $ServerManager.Dispose()
    }
}

<#
.Synopsis
   Adds a new binding to a Microsoft IIS7.0+ website.
.DESCRIPTION
   Adds a new binding to the existing collection of bindings for a website.
.EXAMPLE
   Add-IisSiteBinding -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -HostHeader www.contoso.com
.EXAMPLE
   Add-IisSiteBinding -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -IP 8.8.5.5 -Port 443 -Protocol https -HostHeader www.contoso.com
.NOTES
   IP, Port, and Protocol paremters default to *, 80, and http respectively.
   Protocol parameter sets the Microsoft.Web.Administration.Binding.Protocol property.
   Binding format is ultimately built into `'IP:Port:HostHeader`' format.
   Example: "*:80:www.contoso.com" 
#>
function Add-IisSiteBinding
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string[]]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1)]
        [string[]]$SiteName,

        # Param3 help description
        [Parameter(Mandatory=$false,
                   Position=2)]
        [string[]]$Ip = "*",

        # Param4 help description
        [Parameter(Mandatory=$false,
                   Position=3)]
        [string[]]$Port = "80",

        # Param5 help description
        [Parameter(Mandatory=$true,
                   Position=4)]
        [string[]]$HostHeader,
        
        # Param6 help description
        [Parameter(Mandatory=$false,
                   Position=5)]
        [string]$Protocol = "http"
    )

    Begin
    {
    }
    Process
    {
        ForEach ($Server in $ComputerName)
        {
            Write-Verbose "Connecting to $Server"
            $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

            try
            {
                foreach ($site in $SiteName)
                {
                    try
                    {
                        Write-Verbose "Getting site $site"
                        $Site = $ServerManager.Sites["$site"]

                        ForEach ($Header in $HostHeader)
                        {
                            $NewBinding = "$($Ip):$($Port):$($Header)"
                            $binding = $Site.Bindings.CreateElement()
                            $binding.Protocol = $Protocol
                            $binding.BindingInformation = $NewBinding

                            Write-Verbose "Adding $NewBinding to site"
                            $Site.Bindings.AddAt(0, $binding) | Out-Null
                        }
                    }

                    catch
                    {
                        Write-Warning -Message $Error[0]
                    }
                }
            }

            catch
            {
                Write-Warning -Message $Error[0]
            }

            finally
            {
                Write-Verbose "Committing and cleaing up . . ."
                $ServerManager.CommitChanges()
                $ServerManager.Dispose()
            }
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Updates the physical path of an IIS website.
.DESCRIPTION
   Sets the physical path of an IIS website.  Does not return any value.
.EXAMPLE
   Set-IisSiteCodePath -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -CodePath "C:\inetpub\wwwroot\codepath"
.EXAMPLE
   Set-IisSiteCodePath -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -CodePath "\\Fileshare\path\to\new\code"
.NOTES
   Can be run against multiple computers simultaneously.  Accepts both UNC and Literal Paths.
#>
function Set-IisSiteCodePath
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string[]]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1)]
        [string[]]$SiteName,

        # Param3 help description
        [Parameter(Mandatory=$true,
                   Position=2)]
        [string]$CodePath
    )

    Begin
    {
    }
    Process
    {
        ForEach ($Server in $ComputerName)
        {
            Write-Verbose "Connectiong to $Server"
            $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

            try
            {
                ForEach ($WebSite in $SiteName)
                {
                    Write-Verbose "Getting site $Website"
                    $Site = $ServerManager.Sites["$WebSite"]
    
                    try
                    {
                        Write-Verbose "Updating root path $codepath for site"
                        $Site.Applications[0].VirtualDirectories["/"].PhysicalPath = $Codepath
                    }

                    catch
                    {
                        Write-Warning $Error[0]
                    }
                }
            }

            catch
            {
                Write-Warning -Message $Error[0]
            }

            finally
            {
                Write-Verbose "Committing and cleaing up . . ."                
                $ServerManager.CommitChanges()
                $ServerManager.Dispose()
            }
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Restarts an IIS application pool.
.DESCRIPTION
   Restarts an IIS application pool.
.EXAMPLE
   Restart-IisApplicationPool -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso
.EXAMPLE
   Restart-IisApplicationPool -ComputerName 8.8.8.8, 8.8.4.4 -ApplicationPoolName contoso
.Example
   Get-IisSite -ComputerName 8.8.8.8, 8.8.4.4 -Sitename contoso | Restart-IisApplicationPool -ComputerName 8.8.8.8, 8.8.4.4
.Example
   PS> $site = Get-IisSite -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso
   PS> Restart-IisApplicationPool -InputObject $site -ComputerName 8.8.8.8, 8.8.4.4
.NOTES
   Cmdlet will recycle the application pool if it is in a running state.  If the application pool is not in a running state,
   the cmdlet will stop the application pool, wait 3 seconds, and then start the application pool.
#>
function Restart-IisApplicationPool
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [String[]]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName="InputObject",
                   ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject,

        # Param3 help description
        [Parameter(Mandatory=$true,
                   Position=2,
                   ParameterSetName="SiteName")]
        [string[]]$SiteName,

        [Parameter(Mandatory=$true,
                   Position=3,
                   ParameterSetName="ApplicationPoolname")]
        [string[]]$ApplicationPoolName
    )

    Begin
    {
    }
    Process
    {
        ForEach ($Server in $ComputerName)
        {
            Write-Verbose "Connecting to $Server"
            $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")
            
            if ($SiteName)
            {
                try
                {
                    ForEach ($site in $SiteName)
                    {
                        Write-Verbose "Getting application pool for site $site"
                        $ApplicationPool = $ServerManager.Sites["$SiteName"].Applications["/"].ApplicationPoolName
                        $state = $ServerManager.ApplicationPools[$ApplicationPool].State

                        if ($state -eq "Started")
                        {
                            Write-Verbose "Application pool is started, recycling . . ."
                            $ServerManager.ApplicationPools[$ApplicationPool].Recycle()
                        }

                        else
                        {
                            Write-Verbose "Application pool not in a `'started`' state, starting and stopping . . ."
                            $ServerManager.ApplicationPools[$ApplicationPool].Stop()
                            Start-Sleep -Seconds 3
                            $ServerManager.ApplicationPools[$ApplicationPool].Start()
                        }
                    }
                }

                catch
                {
                    Write-Error -Message "$Error[0]"
                }

                finally
                {
                    Write-Verbose "Committing and cleaing up . . ."
                    $ServerManager.CommitChanges()
                    $ServerManager.Dispose()
                }
            }

            if ($InputObject)
            {
                try
                {
                    ForEach ($SiteObject in $InputObject)
                    {
                        Write-Verbose "Getting application pool for site $site"
                        $ApplicationPool = $SiteObject.ApplicationPool
                        $state = $ServerManager.ApplicationPools[$ApplicationPool].State

                        if ($state -eq "Started")
                        {
                            Write-Verbose "Application pool is started, recycling . . ."
                            $ServerManager.ApplicationPools[$ApplicationPool].Recycle()
                        }

                        else
                        {
                            Write-Verbose "Application pool not in a `'started`' state, starting and stopping . . ."
                            $ServerManager.ApplicationPools[$ApplicationPool].Stop()
                            Start-Sleep -Seconds 3
                            $ServerManager.ApplicationPools[$ApplicationPool].Start()
                        }
                    }
                }

                catch
                {
                    Write-Warning -Message $Error[0]
                }

                finally
                {
                    Write-Verbose "Committing and cleaing up . . ."                   
                    $ServerManager.CommitChanges()
                    $ServerManager.Dispose()
                }
            }

            if ($ApplicationPoolName)
            {
                try
                {
                    ForEach ($ApplicationPool in $ApplicationPoolName)
                    {
                        Write-Verbose "Getting application pool $ApplicationPool"
                        $state = $ServerManager.ApplicationPools[$ApplicationPool].State

                        if ($state -eq "Started")
                        {
                            Write-Verbose "Application pool started, recycling . . ."
                            $ServerManager.ApplicationPools[$ApplicationPool].Recycle()
                        }

                        else
                        {
                            Write-Verbose "Application pool not in a `'started`' state, starting and stopping . . ."
                            $ServerManager.ApplicationPools[$ApplicationPool].Stop()
                            Start-Sleep -Seconds 3
                            $ServerManager.ApplicationPools[$ApplicationPool].Start()
                        }
                    }
                }

                catch
                {
                    Write-Warning "$Error[0]"
                }

                finally
                {
                    Write-Verbose "Committing and cleaing up . . ."
                    $ServerManager.CommitChanges()
                    $ServerManager.Dispose()
                }
            }
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Starts an IIS website.
.DESCRIPTION
   Starts an IIS website.
.EXAMPLE
   Start-IisSite -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso
.NOTES
   Cmdlet starts an IIS website.
#>
function Start-IisSite
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        $ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1)]
        $SiteName
    )

    Begin
    {
    }
    Process
    {

        ForEach ($Server in $ComputerName)
        {
            
            Write-Verbose "Connectiong to $Server"
            $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

            try
            {
                Write-Verbose "Getting site $SiteName"
                $Site = $ServerManager.Sites[$SiteName]
           
                Write-Verbose "Starting site $SiteName"
                $Site.Start() | Out-Null
                $ServerManager.Dispose()
            }

            catch
            {
                Write-Warning "$Error[0]"
            }

            finally
            {
                $ServerManager.Dispose()
            }
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Restarts an IIS application pool.
.DESCRIPTION
   Restarts an IIS application pool.
.EXAMPLE
   StopIisSite -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso
.NOTES
   Cmdlet will stop a website.
#>
function Stop-IisSite
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        $ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1)]
        $SiteName
    )

    Begin
    {
    }
    Process
    {
        ForEach ($Server in $ComputerName)
        {
            Write-Verbose "Connectiong to $Server"
            $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

            try
            {
                Write-Verbose "Getting site $sitename"
                $Site = $ServerManager.Sites[$SiteName]
            
                Write-Verbose "Stopping site $SiteName"
                $Site.Stop() | Out-Null
            }

            catch
            {
                Write-Warning -Message $Error[0]
            }

            finally
            {
                $ServerManager.Dispose()
            }
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Creates new IIS websites.
.DESCRIPTION
   Creates new IIS websites from a PSCustomObject or named parameter set. The input object matches the design of the object(s) returned from Get-IisSite.
.EXAMPLE
   Get-IisSite -ComputerName 8.8.8.8 | New-IisSite -ComputerName 8.8.4.4
.Example
    New-IisSite -ComputerName 8.8.8.8, 8.8.4.4 -SiteName contoso -CodePath "D:\inetpub\wwwroot\code" -ApplicationPoolName ContosoAppPool -Bindings "*:80:www.consoso.com", "*:443:www.contoso.com"
.NOTES
   Instead of using a continually incrementing site ID starting with 1, the website ID is generated using the same methodology from IIS6.
   A hash is taken from the name of the website, and the absolute value of the hash is set as the site ID.  If the site ID already exists,
   it will be incremented until the ID is unique.
#>
function New-IisSite
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [String[]]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1,
                   ValueFromPipeline=$true,
                   ParameterSetName="InputObject")]
        [PSCustomObject[]]$InputObject,

        # Param3 help description
        [Parameter(Mandatory=$true,
                   Position=2,
                   ParameterSetName="SingleSite")]
        [String]$SiteName,

        # Param4 help description
        [Parameter(Mandatory=$true,
                   Position=3,
                   ParameterSetName="SingleSite")]
        [String]$CodePath,

        # Param5 help description
        [Parameter(Mandatory=$false,
                   Position=4,
                   ParameterSetName="SingleSite")]
        [String]$ApplicationPoolName,

        # Param6 help description
        [Parameter(Mandatory=$true,
                   Position=5,
                   ParameterSetName="SingleSite")]
        [String[]]$Bindings
    )

    Begin
    {
        function create_site
        {
            Param
            (
                [Parameter(Mandatory=$true,
                           Position=0)]
                $NewSite,

                [Parameter(Mandatory=$true,
                           Position=1)]
                $SiteName,
                
                [Parameter(Mandatory=$true,
                           Position=2)]
                $CodePath

            )

            $NewSite.Name = $SiteName
            $NewSite.Id = [System.Math]::Abs($NewSite.Name.GetHashCode())
            $NewSite.Applications.Add("/", "$CodePath") | Out-Null
            Return $NewSite
        }

        function add_virtual_directories
        {
            Param
            (
                [Parameter(Mandatory=$true,
                           Position=0)]
                $Site,

                [Parameter(Mandatory=$true,
                           Position=1)]
                $VirtualDirectory
            )
           
            $NewVirtualDirectory = $Site.Applications["/"].VirtualDirectories.CreateElement()
            $NewVirtualDirectory.Path = "$($VirtualDirectory.Path)"
            $NewVirtualDirectory.PhysicalPath = "$($VirtualDirectory.PhysicalPath)"
            $NewSite.Applications["/"].VirtualDirectories.Add($NewVirtualDirectory) | Out-Null
            Return $NewSite
        }

        function add_bindings
        {
            Param
            (
                [Parameter(Mandatory=$true,
                           Position=0)]
                $Site,

                [Parameter(Mandatory=$true,
                           Position=1)]
                $Binding
            )
           
            $NewBinding = $Site.Bindings.CreateElement()

            if ($Binding.split(":")[1] -eq "80")
            {
                $NewBinding.Protocol = "http"
            }
            else
            {
                $NewBinding.Protocol = "http"
            }
                        
            $NewBinding.BindingInformation = $Binding
            $NewSite.Bindings.AddAt(0, $NewBinding) | Out-Null
            Return $NewSite
        }
    }
    Process
    {
        ForEach ($Server in $ComputerName)
        {
            if ($InputObject)
            {
                Write-Verbose "Connecting to $Server"
                $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

                ForEach ($Site in $InputObject)
                {
                    try
                    {
                        $NewSite = $ServerManager.Sites.CreateElement()
                        $NewSite = create_site -NewSite $NewSite -SiteName $Site.Name -CodePath ($Site.VirtualDirectories | Where-Object { $_.Path -eq "/" } | Select -ExpandProperty PhysicalPath)

                        Foreach ($VirtualDirectory in $Site.VirtualDirectories | Where-Object { $_.Path -ne "/" })
                        {
                            Write-Verbose "Creating Virtual Directory with name `"$($VirtualDirectory.Path)`" and Physical Path `"$($VirtualDirectory.PhysicalPath)`""
                            $NewSite = add_virtual_directories -Site $NewSite -VirtualDirectory $VirtualDirectory
                        }

                        Write-Verbose "Adding host headers for site `"$($site.Name)`""
                        Foreach ($Binding in $Site.Bindings)
                        {
                            $NewSite = add_bindings -Site $NewSite -Binding $Binding
                        }

                        if ( ($ServerManager.ApplicationPools["$($Site.ApplicationPool)"]) -ne $null )
                        {
                            Write-Verbose "Setting site `"$($Site.Name)`" to use application pool `"$($Site.ApplicationPool)`""
                            $NewSite.Applications["/"].ApplicationPoolName = $Site.ApplicationPool
                        }

                        $ServerManager.Sites.Add($NewSite) | Out-Null
                    }

                    catch
                    {
                        Write-Warning -Message "$($Error[0])"
                    }

                    finally
                    {
                        Write-Verbose "Committing and cleaning up . . ."
                        $ServerManager.CommitChanges()
                        $ServerManager.Dispose()
                    }
                }
            }

            else
            {
                Write-Verbose "Connecting to $Server"
                $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

                try
                {
                    $NewSite = $ServerManager.Sites.CreateElement()
                    $NewSite = create_site -NewSite $NewSite -SiteName $SiteName -CodePath $CodePath

                    Write-Verbose "Adding host headers for site `"$SiteName`""
                    Foreach ($Binding in $Bindings)
                    {
                        $NewSite = add_bindings -Site $NewSite -Binding $Binding
                    }

                    if ($ApplicationPoolName)
                    {
                        Write-Verbose "Setting site `"$SiteName`" to use application pool `"$ApplicationPoolName`""
                        $NewSite.Applications["/"].ApplicationPoolName = $ApplicationPoolName
                    }

                    $ServerManager.Sites.Add($NewSite) | Out-Null
                }

                catch
                {
                    Write-Warning -Message "$($Error[0])"
                }

                finally
                {
                    Write-Verbose "Committing and cleaing up . . ."
                    $ServerManager.CommitChanges()
                    $ServerManager.Dispose()
                }
            }
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Creates new IIS application pools.
.DESCRIPTION
   Creates new IIS application pools from a PSCustomObject or a named parameter set. The input object matches the design of the object(s) returned from Get-IisApplicationPool.
.EXAMPLE
   Get-IisApplicationPool -ComputerName 8.8.8.8 | New-IisApplicationPool -ComputerName -8.8.4.4 -Username "foo" -Password "bar"
.Example
    New-IisApplicationPool -ComputerName 8.8.8.8, 8.8.4.4 -ApplicationPoolName contoso -AutoStart $true -ManagedPipelineMode Integrated -ManagedRuntimeVersion "v4.0" -Enable32bit $false -IdentityType SpecificUser -Username foo -Password bar
.NOTES
   
#>
 function New-IisApplicationPool
 {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [String[]]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=1,
                   ValueFromPipeline=$true,
                   ParameterSetName="InputObject")]
        [PSCustomObject[]]$InputObject,

        # Param3 help description
        [Parameter(Mandatory=$true,
                   Position=2,
                   ParameterSetName="SingleSite")]
        [String]$ApplicationPoolName,

        [Parameter(Mandatory=$true,
                   Position=3,
                   ParameterSetName="SingleSite")]
        [bool]$AutoStart,

        # Param4 help description
        [ValidateSet("Integrated", "Classic")]
        [Parameter(Mandatory=$true,
                   Position=4,
                   ParameterSetName="SingleSite")]
        [String]$ManagedPipelineMode,

        [ValidateSet("v1.1", "v2.0", "v4.0")]
        [Parameter(Mandatory=$true,
                   Position=5,
                   ParameterSetName="SingleSite")]
        [String]$ManagedRuntimeVersion,

        # Param5 help description
        [Parameter(Mandatory=$false,
                   Position=6,
                   ParameterSetName="SingleSite")]
        [bool]$Enable32bit = $false
    )
 
    Begin
    {
    }
    Process
     {
        if ($InputObject)
        {
            ForEach ($Server in $ComputerName)
            {
                Write-Verbose "Connecting to $Server"
                $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

                ForEach ($ApplicationPool in $InputObject)
                {
                    try
                    {
                        Write-Verbose "Creating Application Pool with name `"$($ApplicationPool.Name)`""
                        $NewApplicationPool = $ServerManager.ApplicationPools.CreateElement()
                        $NewApplicationPool.Enable32BitAppOnWin64 = $ApplicationPool.Enable32bit
                        $NewApplicationPool.AutoStart = $ApplicationPool.Autostart
                        $NewApplicationPool.Name = $ApplicationPool.Name
                        $NewApplicationPool.ManagedPipelineMode = New-Object Microsoft.Web.Administration.ManagedPipelineMode

                        if ($ApplicationPool.ManagedPipelineMode -eq "Classic")
                        {
                            $NewApplicationPool.ManagedPipelineMode.Value__ = 0
                        }

                        $NewApplicationPool.ManagedRuntimeVersion = "$ApplicationPool.ManagedRuntimeVersion"
                        $NewApplicationPool.ProcessModel.IdentityType = $ApplicationPool.IdentityType	

                        $ServerManager.ApplicationPools.Add($NewApplicationPool) | Out-Null
                    }

                    catch
                    {
                        Write-Warning -Message "$($Error[0])"
                    }

                    finally
                    {
                        Write-Verbose "Committing changes . . ."
                        $ServerManager.CommitChanges()
                        $ServerManager.Dispose()
                    }
                }
            }
        }

        else
        {
            ForEach ($Server in $ComputerName)
            {
                try
                {
                    Write-Verbose "Connecting to $Server"
                    $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

                    Write-Verbose "Creating Application Pool with name `"$ApplicationPoolName`""
                    $NewApplicationPool = $ServerManager.ApplicationPools.CreateElement()
                    $NewApplicationPool.Enable32BitAppOnWin64 = $Enable32bit
                    $NewApplicationPool.AutoStart = $AutoStart
                    $NewApplicationPool.Name = $ApplicationPoolName
                    $NewApplicationPool.ManagedPipelineMode = New-Object Microsoft.Web.Administration.ManagedPipelineMode

                    if ($ApplicationPool.ManagedPipelineMode -eq "Classic")
                    {
                        $NewApplicationPool.ManagedPipelineMode.Value__ = 0
                    }

                    $NewApplicationPool.ManagedRuntimeVersion = $ManagedRuntimeVersion
                    $NewApplicationPool.ProcessModel.IdentityType = $IdentityType
                    $ServerManager.ApplicationPools.Add($NewApplicationPool) | Out-Null
                }

                catch
                {
                    Write-Warning -Message "$($Error[0])"
                }

                finally
                {
                    Write-Verbose "Committing changes . . ."
                    $ServerManager.CommitChanges()
                    $ServerManager.Dispose()
                }
            }
        }
    }
    End
    {
    }
 }



<#
.Synopsis
   Sets the user an IIS application pool runs under.
.DESCRIPTION
   Sets the IdentityType of an application pool to SpecificUser and sets the username and password.
.EXAMPLE
   Set-IisApplicationPoolUser -ComputerName 8.8.8.8, 8.8.4.4 -ApplicationPoolName contoso, foo, bar, cheese, sausage -Username "domain\foo" -Password bar
#>
function Set-IisApplicationPoolUser
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]$ComputerName,

        # Param2 help description
        [Parameter(Mandatory=$false,
                   Position=1)]
        [string[]]$ApplicationPoolName,

        [Parameter(Mandatory=$false,
                   Position=2)]
        [string]$UserName,

        [Parameter(Mandatory=$false,
                   Position=3)]
        [string]$Password
    )

    Begin
    {
    }
    Process
    {
        ForEach ($Server in $ComputerName)
        {
            Write-Verbose "Connecting to $Server"
            $ServerManager = [Microsoft.Web.Administration.ServerManager]::OpenRemote("$Server")

            try
            {
                foreach ($ApplicationPool in $ApplicationPoolName)
                {
                    Write-Verbose "Setting $ApplicationPool to run under user $UserName"
                    $ApplicationPool = $ServerManager.ApplicationPools["$ApplicationPool"]
                    $ApplicationPool.ProcessModel.IdentityType = 'SpecificUser'
                    $ApplicationPool.ProcessModel.Password = $Password
                    $ApplicationPool.ProcessModel.UserName = $UserName
                }
            }

            catch
            {
                Write-Warning -Message "$($Error[0])"
            }

            finally
            {
                Write-Verbose "Committing changes . . ."
                $ServerManager.CommitChanges()
                $ServerManager.Dispose()
            }
        }
    }
    End
    {
    }
}
