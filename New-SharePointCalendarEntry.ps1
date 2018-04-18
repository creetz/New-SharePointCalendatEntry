<#
    .SYNOPSIS
    New-SharePointCalendarEntry.ps1
   
   	Christian Reetz
    (Updated by Christian Reetz)
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	11.04.2018
	
    .DESCRIPTION

    This script connects to multiple SharePointCalendar and creates events.
    
    The SharePointCalendars are managed by a settings.xml-File
    
    The events which will be created are managed by the csv-File.
    New-ShrePointCalendarEntry.csv
    
    My recommendation:
    Use multiple Teams/o365-groups and create sharepoint-calendarlists to share with guests and other-users.

    #>

$settings = Import-Clixml -Path $PSScriptRoot\settings.xml
$now = get-date
$now2 = "$($now.year)-$($now.month)-$($now.day)T$($now.hour):$($now.minute):$($now.Second)Z"

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

        #View Items
        <#
        $camlQuery = new-object Microsoft.SharePoint.Client.CamlQuery;
        $camlQuery.ViewXml = "<View>
        <Query>
         <Where> 
              <Eq>   
               <FieldRef Name='Category' /> 
                    <Value Type='Text'>$SharePointKalendarCategory</Value>
                 </Eq>
               </Where>
         </Query>
        </View>"
        
        $listItems = $List.GetItems($camlQuery);
        $context.Load($listItems)
        $context.ExecuteQuery();
        
        foreach($item in $listItems)
        {
         Write-Host "Title: " $item["Title"]
        }
        #>
        #Delete Items
        
        $camlQuery = new-object Microsoft.SharePoint.Client.CamlQuery;
        $camlQuery.ViewXml = "<View>
        <Query>
         <Where> 
              <And>
              <Eq>   
               <FieldRef Name='Category' /> 
                    <Value Type='Text'>$SharePointKalendarCategory</Value>
               </Eq> 
               <Gt>   
               <FieldRef Name='EventDate' /> 
                    <Value Type='DateTime' IncludeTimeValue='false'>$now2</Value>
               </Gt>
               </And>
          </Where>
         </Query>
        </View>"
        
        #$caml="<View><Query><Where><Eq><FieldRef Name='Created'/><Value Type='DateTime' IncludeTimeValue='FALSE'>2014-09-11</Value></Eq></Where></Query></View>"

        $listItems = $List.GetItems($camlQuery);
        $context.Load($listItems)
        $context.ExecuteQuery();

        write-host "Total Number of List Items found:"$ListItems.Count
 
        if ($ListItems.Count -gt 0)
        {
            #Loop through each item and delete
            For ($i = $ListItems.Count-1; $i -ge 0; $i--)
            {
                $ListItems[$i].DeleteObject()
            }
            $Context.ExecuteQuery()
             
            Write-Host "All List Items deleted Successfully!"
        }


        #Add Items

        Write-Host -ForegroundColor White "Create Events:"
        
        foreach ($event in $events)
        {
            $eventdate = ($event.eventdate -split '/' -split ':' -split ' ')
            $eventdate2 = Get-Date -Month $eventdate[0] -Day $eventdate[1] -Year $eventdate[2] -Hour $eventdate[3] -Minute $eventdate[4]
            if ($eventdate2 -gt $now)
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
}