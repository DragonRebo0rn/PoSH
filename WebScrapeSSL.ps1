<#
        .AUTHOR
            DragonReb0rn
        
        .SYNOPSIS
            Scapes the SSL Labs website based on URLs from an input file, performs a delta between previous and current month, then moves files accordingly.

        .DESCRIPTION
            Scrapes each website in the input file and pulls back errors, warnings, highlights, and overall rating. All results are exported to a CSV file.
            Results are then compared from previous month and the changes are recorded in a seperate CSV.
            Once the delta is performed, previous month result is moved into archive and current month is moved into previous month and results are emailed out.

        .NOTES
            Requires Powershell v3

        .UPDATES
            Updated 2015-07-23
            Updated 2015-08-03 (EmailScrapeData update)

    #>

Function WebScrape {
#Create empty array
$array = @()

#pull in contents of the input file
$URLS = Get-Content C:\_Scripts\_Webscrape\URLs.csv

#prompt for credentials when testing(should be commented out in Production)
$creds = Get-Credential

#begin foreach loop to cycle through the URLs
foreach ($url in $urls) 
    {

        $URL

         Try {
        
            #starts the site evaluation process on SSL Labs
            $IE=new-object -com internetexplorer.application
            $IE.navigate2($URl)
            $IE.visible=$False
            }

        Catch {

            #close any open internet explorer processes
            get-process iexplore -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue

            #sleep this process for 10 seconds if the resource is not available
            Start-Sleep -s 10

            #retry the process if it errors
            $IE=new-object -com internetexplorer.application
            $IE.navigate2($URl)
            $IE.visible=$False

              }

        #Wait for 120 seconds for SSL Labs to finish processing the site evaluation
        Start-Sleep -s 120
       
        
        Try {

            #Scrape the contents of the SSL Labs page and store all results
            $Scrape = Invoke-WebRequest -URI $URL -Credential $creds

            Start-Sleep -s 15

            #Scrape the contents of the SSL Labs page and store all results again
            $Scrape = Invoke-WebRequest -URI $URL -Credential $creds

            Start-Sleep -s 15

            }

        Catch {

            #close any open internet explorer processes
            get-process iexplore -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue

            Start-Sleep -s 5

            #Open an IE session cause the proxy sux0rz
            $IE=new-object -com internetexplorer.application
            $IE.navigate2($URL)
            $IE.visible=$False

            Start-Sleep -s 120

            #Scrape the contents of the SSL Labs page and store all results
            $Scrape = Invoke-WebRequest -URI $URL -Credential $creds -ErrorAction SilentlyContinue

              }

        #Store specific content from the scrape into variables
        $warningBox = $scrape.ParsedHtml.getElementsByTagName("div") | where "classname" -match "warningBox" | select -ExpandProperty InnerText
        $errorBox = $scrape.ParsedHtml.getElementsByTagName("div") | where "classname" -match "errorBox" | select -ExpandProperty InnerText
        $highlightBox = $scrape.ParsedHtml.getElementsByTagName("div") | where "classname" -match "highlightBox" | select -ExpandProperty InnerText
        $Site = $scrape.ParsedHtml.getElementsByTagName("div") | where "classname" -match "reportTitle" | select -ExpandProperty InnerText
        $Rating = $scrape.ParsedHtml.getElementsByTagName("div") | where "classname" -like "rating_*" | select -ExpandProperty InnerText

	    #Create a new PSObject and store the scraped content from the variables in the object					                
	    $Entry = New-Object psobject
            $Entry | Add-Member -memberType noteProperty -name "Site" -Value $Site
            $Entry | Add-Member -memberType noteProperty -name "Rating" -Value $Rating
	    $Entry | Add-Member -memberType noteProperty -name "warningBox" -Value $warningBox
	    $Entry | Add-Member -memberType noteProperty -name "errorBox" -Value $errorBox
	    $Entry | Add-Member -memberType noteProperty -name "highlightBox" -Value $highlightBox
    
        #Join entries with multiple values into one string in preparation for export
        $entry.warningBox = $entry.warningBox -join ';'
        $entry.errorBox = $entry.errorBox -join ';'

        #catch all for sites that can't resolve the URL
        if (!$Site){
        $Site = "Unable to resolve $URL"
                   }

        #add contents of the PSObject into the array
        $array += $Entry

        #end the internet explorer processes opened
        get-process iexplore -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue

        $Site = ""

			   
    }

#export the contents of the array into a CSV
$array | Export-CSV C:\_Scripts\_Webscrape\Current\Results.csv -NoTypeInformation -Force


#remove string not needed in file
[io.file]::readalltext("C:\_Scripts\_Webscrape\Current\Results.csv").replace("Due to a recently discovered bug in Apple's code, your browser is exposed to MITM attacks. Click here for more information. ","") | Out-File C:\_Scripts\_Webscrape\Current\$(Get-Date -f yyyy-MM-dd)-Results.csv -Encoding ascii –Force

#remove old file with old string
#Remove-Item "C:\_Scripts\_Webscrape\Current\Results.csv" -Force

}

Function WebScrapeDelta {

#import previous and current month's SSL Labs results
$OldResults = import-csv -path C:\_Scripts\_Webscrape\previous\*Results.csv
$NewResults = import-csv -path C:\_Scripts\_Webscrape\current\*Results.csv

#create empty array for storing data
$output = @()
    foreach ($Column in $OldResults)
    {
        #match the sites from previous and current, then begin the comparisons
        $Resultz = $NewResults | Where-Object {$Column.Site -eq $_.Site}
        $RatingResults = if ($Column.Rating -ne $Resultz.Rating) {"Previous: " + $Column.Rating + " | Now:" + $Resultz.Rating}
        $ErrorBoxResults = if ($Column.errorBox -ne $Resultz.errorBox)  {"Previous: " + $Column.errorBox + " | Now:" + $Resultz.errorBox}
        $WarningBoxResults = if ($Column.warningBox -ne $Resultz.warningBox)  {"Previous: " + $Column.warningBox + " | Now:" + $Resultz.warningBox}
        $HighlightBoxResults = if ($Column.highlightBox -ne $Resultz.highlightBox)  {"Previous: " + $Column.highlightBox + " | Now:" + $Resultz.highlightBox}

        #Output the changes detected in a PSObject
        $output += New-Object PSObject -Property @{
            Site = $Column.Site
            Rating = $RatingResults
            ErrorBox = $ErrorBoxResults
            WarningBox = $WarningBoxResults
            HighlightBox = $HighlightBoxResults

        }
    }

#export the array into a CSV file with the relevant column names
$output | Select-Object site,rating,errorbox,warningbox,highlightbox | export-csv -path C:\_Scripts\_Webscrape\Delta\$(Get-Date -f yyyy-MM-dd)-Changes.csv -NoTypeInformation
}

Function MoveScrapeData {

#Move the previous months results into the archive folder
Move-Item C:\_Scripts\_Webscrape\Previous\*Results.csv C:\_Scripts\_Webscrape\Archive

#Move the current month's results into the previous folder
Move-Item C:\_Scripts\_Webscrape\Current\*Results.csv C:\_Scripts\_Webscrape\Previous

}

Function EmailScrapeDelta {

#Set mail attributes
$relay = ""
$To = ""
$From = ""
$Subject = "SSL Labs Delta"
$Attachments = @()
$Attachments += "C:\_Scripts\_Webscrape\Delta\*changes.csv"
$Attachments += "C:\_Scripts\_Webscrape\Current\*results.csv"

#Send the message
Send-MailMessage -SmtpServer $relay -to $to -from $from -Subject $Subject -Attachments $Attachments

#Move the delta file into the previous folder in preparation for the next run
Move-Item C:\_Scripts\_Webscrape\Delta\*Changes.csv C:\_Scripts\_Webscrape\Delta\Archive
}

#Begin Web Scrape
WebScrape

#sleep for 5 seconds to give the file a chance to generate
Start-Sleep -s 5

#perform the delta between previous and current month
WebScrapeDelta

#sleep for 5 seconds to give the file a chance to generate
Start-Sleep -s 5

#Email the scrape delta data and move the file into the previous folder
EmailScrapeDelta

#Execute the file moves
MoveScrapeData
