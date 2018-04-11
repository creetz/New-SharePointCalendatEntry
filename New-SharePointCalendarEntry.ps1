<#
    .SYNOPSIS
    New-SharePointCalendatEntry.ps1
   
   	Christian Reetz
    (Updated by Christian Reetz)
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	11.04.2018
	
    .DESCRIPTION

    This script opens multiple SharePointCalendar and creates events, which can be managed by a csv-file.
    Please use the example csv-file --> New-SharedPointCalendar.csv
    The SharePointCalendars are managed by the settings.xml-File.
   	Please look at the setting.xml-ExampleFile, too.
    
    my recommendation:
    Use multiple Teams/o365-groups create multiple kalendar to share with guests and other-users.

    #>

$settings = Import-Clixml -Path $PSScriptRoot\settings.xml

#Add references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
$cred = Get-Credential

foreach ($SharePointCalendar in $settings.SharePointCalendar)
{
    $SiteURL = $SharePointCalendar.split(';')[0]
    $SharePointKalendarCategory = $SharePointCalendar.split(';')[1]
    $SharePointKalendarTitle = $SharePointCalendar.split(';')[2]
    $SharePointKalendarIsEnabled = $SharePointCalendar.split(';')[3]

    if ($SharePointKalendarIsEnabled -eq 1)
    {
        Write-Host -ForegroundColor White "Connect to: " -NoNewline
        Write-Host -ForegroundColor Yellow $SiteURL

        Write-Host -ForegroundColor White "CalendarTitle: " -NoNewline
        Write-Host -ForegroundColor Yellow "$SharePointKalendarTitle (category: $SharePointKalendarCategory)" 

        $events = Import-Csv -Path $PSScriptRoot\New-SharePointCalendarEntry.csv -Delimiter ';' | ? {$_.Category -eq "$SharePointKalendarCategory"}

        Write-Host -ForegroundColor White "Event-Count: " -NoNewline
        Write-Host -ForegroundColor Yellow $events.count
        
        #Bind to site collection
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.username,$cred.Password)
        $Context.Credentials = $Creds
        
        #Retrieve lists
        $Lists = $Context.Web.Lists
        $Context.Load($Lists)
        $Context.ExecuteQuery()
        
        $List = $Context.Web.Lists.GetByTitle($SharePointKalendarTitle)
        $Context.Load($List)
        $Context.ExecuteQuery()

        Write-Host -ForegroundColor White "Create Events:"
        
        foreach ($event in $events)
        {
            Write-Host -ForegroundColor Green "$($event.EventDate) - $($event.EndDate) | $($event.Title)"
            
            #Adds an item to the list
            $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $Item = $List.AddItem($ListItemInfo)
            $Item["Title"] = "$($event.Title)"
            $Item["EventDate"] = $event.EventDate
            $Item["EndDate"] = $event.EndDate
            $Item["Description"] = "$($event.Description)"
            $Item["Category"] = "$($event.Category)"
            $Item["Location"] = "$($event.Location)"
            $Item.Update()
            $Context.ExecuteQuery()   
        }
    }   
}