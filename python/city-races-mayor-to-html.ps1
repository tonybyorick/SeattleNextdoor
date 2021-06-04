#===================================================================
#                   city-races-mayor-to-html.ps1
#
# In:    City-Races-Endorsements-Mayor.xlsx
# Out:   city-mayoral-race-2021.html
# Bitly: http://bit.ly/SeattleView-21M
#
# Changes:
# May 31 Cell with candidate name formatted with black borders
#===================================================================


#------------------------------------------
# FUNCTIONS
#------------------------------------------

function Get-TimesResult {
   Param ([int]$a,[int]$b)
   $c = $a * $b
   Write-Output $c
}

#------------------------------------------
# function print_cell
#
#
#------------------------------------------
function print_cell {
   Param ([String]$text, [String]$OutputFullPath, [String]$col)

   $noFormatText = $text
   $noFormatText = $noFormatText.Replace('$', "")
   $noFormatText = $noFormatText.Replace(',', "")
   
   #$outText =  "
   #<td>$text</td>
   #"


#border-top: 5px solid red;

   if ($noFormatText -match "^\d+$") {   # Right Justify Numbers
     $outText = "
     <td style='text-align: right; border-top: 5px solid black;'>$text</td>
     " 
     }
   elseif ($col -eq 1) {
     $outText = "
     <td style='color: white; background-color: #669999; border-top: 5px solid black; border-right: 5px solid black; border-bottom: 5px solid black;border-left: 5px solid black; font-size: 50pt;'>$text</td> 
     "
      }
   else {
      $outText =  "
      <td>$text</td>
      "
      }
   Add-Content -Path $OutputFullPath -Value $outText
   }


#------------------------------------------
# function getCampaignURL
#
# $campaignURL = getCampaignURL($row, $col)
#------------------------------------------
function getCampaignURL {
   Param ([int]$row,[int]$col)

   $cpURL =  $ExcelWorkSheet.Cells.Item($row,$col).text
   $result = "<a href='$cpURL' target='_blank'>Campaign Page</a>"
   return $result
   }


#------------------------------------------
# function getCandEndorseURL
#
# $campaignURL = getCampaignURL($row, $col)
#------------------------------------------
function getCandEndorseURL {
   Param ([int]$row,[int]$col)

   $url =  $ExcelWorkSheet.Cells.Item($row,$col).text
   
   if ($url -eq "" ) {
      return 'No Endorsements Page'
      }
   
   $result = "<a href='$url' target='_blank'>Endorsements Page</a>"
   return $result
   }



#------------------------------------------
# function getCampTags
#
# $formattedTagLines = getCampTags -row $row -col $col
#------------------------------------------
function getCampTags {
   Param ([int]$row,[int]$col)

   $campTags = $ExcelWorkSheet.Cells.Item($row,$col).text
   
   if ($campTags -eq "" ) {
      return '&nbsp;'
      }

   $campTags = $campTags.Replace("<>", "</li><li>")

   $campTags = "<u>Experience and Campaign Goals</u><ul>" + $campTags +
   "</li></ul>"
   
   $result = "$campTags"
   return $result
   }



  #$formattedTagLines = getCampTags -row $thisRow -col $colTagString



# #333399 Dark Blue
# #800000 Brown
# #006666 Dark Cyan
#------------------------------------------
# HTML Section 1
#------------------------------------------
$section_1 = "<!DOCTYPE html>
 <HTML>
 <head>
 <meta http-equiv='content-type' content='text/html; charset=UTF-8'>
 <title>Mayoral Race</title>
 
 <style>
 h1 {
    color: #333399;
    margin-left: 0px;
    font-size: 30pt;
    }
* {
  font-family: Calibri, 'Open Sans', Helvetica, sans-serif;
}
body {
    background-color: white;
}

table, td, th {
  border: 1px solid gray;
}

table {
	border-collapse: collapse;
 }

td {
   font-size: 30pt;
   padding:5px;
}

span.linkDesc {
    color: red;
    font-size: 12pt;
    font-style: italic;
    }
 
span.textDesc {
    color: black;
    font-size: 12pt;
    }



@media only screen and (min-width: 768px) {

.styled-table {
    border-collapse: collapse;
    margin: 25px 0;
    font-size: 0.9em;
    font-family: sans-serif;
    min-width: 400px;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}

.styled-table thead tr {
    background-color: #006666;
    color: #ffffff;
    text-align: left;
}

h1 {
    color: #006666;;   /* Dark Cyan */
    margin-left: 0px;
    font-size: 38pt;
    }

td {
   font-size: 30pt;
   padding:5px;
}

span.linkDesc {
    color: brown;
    font-size: 28pt;
    font-style: italic;
    }
    
span.textDesc {
    color: black;
    font-size: 28pt;
    }
    
    
}
</style>
 
</head>
<body> 
"

#------------------------------------------
# HTML Section 2
#------------------------------------------
$section_2 = "<h1>Seattle 2021 Mayoral Race</h1>

<span class='textDesc'>Candidates are listed by campaign contribution amount as shown on the <a href = 'http://web6.seattle.gov/ethics/elections/campaigns.aspx?cycle=2021&type=contest&IDNum=188&leftmenu=collapsed' target='_blank'>Seattle Ethics and Election Commissons</a> page. Experience and campaign goals listed below are sourced from campaign websites of each candidate.
</span>
"

#------------------------------------------
# HTML Section 3
#------------------------------------------
$section_3 = "
<br>

<span class='linkDesc'> 
<a href = 'http://web6.seattle.gov/ethics/elections/campaigns.aspx?cycle=2021&type=contest&IDNum=188&leftmenu=collapsed' target='_blank'>reference link: Ethics and Elections Commission</a>
</span>
<br>
<span class='linkDesc'> 
<a href = 'http://www.seattle.gov/democracyvoucher/program-data' target='_blank'>reference link: Democracy Voucher Program</a>
</span>




</body>
</html>"

#------------------------------------------
# HTML Table 1 Start
#------------------------------------------
$hdr_style = "style='text-align: center; vertical-align: bottom;'"
$section_table_hdr_1 = "
<thead>
<tr>
<td $hdr_style>Candidate</td>
<td $hdr_style>Contributions to Date</td>
<td $hdr_style>Number of Contributors to Date</td>
<td $hdr_style>Average Contribution Size</td>
<td $hdr_style>Democracy Voucher Donations to Date</td>
</tr>
</thead>
"

# NextDoor Rating <br>2-approve,<br> 1-neither approve or disapprove,<br> 0-disaprove


#------------------------------------------
# Link Definitions
#------------------------------------------
$pos8CampaignLink = "http://web6.seattle.gov/ethics/elections/campaigns.aspx?cycle=2021&type=contest&IDNum=189&leftmenu=collapsed"
$pos9CampaignLink = "http://web6.seattle.gov/ethics/elections/campaigns.aspx?cycle=2021&type=contest&IDNum=190&leftmenu=collapsed"
$CACampaignLink = "http://web6.seattle.gov/ethics/elections/campaigns.aspx?cycle=2021&type=contest&IDNum=193&leftmenu=collapsed"

$GlumazLink = "https://www.glumazforseattlecitycouncil.org/"
$TsimermanLink = "http://alexforamerica.com/"
$IshiiLink = "https://www.booger.rocks/"
$MartinLink = "https://electkatemartin.com/#issues"
$MosquedaLink = "https://www.teamteresa.org/priorities/"

$ThomasLink = "https://www.peopleforbrianna.org/"
$GrantLink = "https://www.claireforcouncil2021.com/"
$EichnerLink = "https://coreyeichner.com/"
$OliverLink = "https://nikkitafornine.com/policies"
$WilliamsonLink = "https://www.google.com/search?q=rebecca+williamson+seattle+city+council&source=hp&ei=nM6ZYIHMIvXL0PEP1IaKmAg&iflsig=AINFCbYAAAAAYJncrMXsgjBI9u8OJ6xvxybv8G5dNdtO&oq=Rebecca+Williamson&gs_lcp=Cgdnd3Mtd2l6EAEYATICCAAyAggAMgIILjICCAAyBQgAEMkDMgIIADICCAAyAggAMgIIADICCAA6CAgAEOoCEI8BUNQRWNQRYNM7aAFwAHgAgAE1iAE1kgEBMZgBAKABAqABAaoBB2d3cy13aXqwAQY&sclient=gws-wiz"
$NelsonLink = "https://www.saraforcitycouncil.com/issues"

$FortneyLink = "https://www.electstevefortney.com"
$HolmesLink = "https://www.holmesforseattle.com/"
$KernerLink = "https://www.google.com/search?q=elect+isabelle+kerner+for+city+attorney&client=opera&hs=ws&sxsrf=ALeKk00l0b6xX52bd0grrGvjMjQ83wNxBg%3A1621532055401&ei=l52mYKXiF-fT5NoP5IayqAk&oq=elect+isabelle+kerner+for+city+attorney&gs_lcp=Cgdnd3Mtd2l6EAMyBwgjELADECdQAFgAYLW2AWgCcAB4AIABdIgBdJIBAzAuMZgBAKoBB2d3cy13aXrIAQHAAQE&sclient=gws-wiz&ved=0ahUKEwjliZ3R5djwAhXnKVkFHWSDDJUQ4dUDCA0&uact=5"

$mlkLaborAbout = "https://www.mlklabor.org/about/"
$mlkLaborMayorEndorse = "https://www.mlklabor.org/news/mlk-labor-endorses-lorena-gonzalez-for-seattle-mayor/"

#------------------------------------------
# Initalize HTML output and write section 1
#$section_1 | Out-File -FilePath $OutputFullPath -encoding ASCII

write-host "`r`n"

$inputFolder = "C:\Users\gb105\OneDrive\go\home\All Seattle\NextDoor\NSCIA"
$inputFullPath = $inputFolder + "\" + "City-Races-Endorsements-Mayor-v2.xlsx"
write-host "Input: " $inputFullPath

# Set Output Folder and Filename
$OutputFolder = $inputFolder
$OutputFullPath = $OutputFolder + "\" + "city-mayoral-race-2021.html"

write-host "Output: " + $OutputFullPath

# Initalize HTML output and write section 1
$section_1 | Out-File -FilePath $OutputFullPath -encoding ASCII

# Write Section 2
Add-Content -Path $OutputFullPath -Value $section_2

#------------------------------------------
#         Open Excel Spreadsheet
#------------------------------------------

# Run Excel Application
$ExcelObj = New-Object -comobject Excel.Application
$ExcelObj.visible=$true

# Open Excel Workbook
$ExcelWorkBook = $ExcelObj.Workbooks.Open($inputFullPath)

# Print names of worksheets
$ExcelWorkBook.Sheets| fl Name, index

# Open worksheet 1
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Sheet1")

# Get Max Rows
$rowMax = ($ExcelWorkSheet.UsedRange.Rows).count
$colMax = ($ExcelWorkSheet.UsedRange.Columns).count
write-host ("rowMax: [" + $rowMax + "]" + " colMax: [" + $colMax + "]") 

# Define Other Columns
$colCampaignLink = 6    # Col F hardcoded here to have Campaign URL
$colCandEndorseLink = 8 # Col H hardcoded for link to candidates endorse page
$colTagString = 7       # Col G hardcoded for link to candidate tags

#----------------------------------
# Write the table header
#----------------------------------

#<table class="styled-table">
#<!-- <table style='width:100%; border: 1px solid red'> -->

$tableStart = "<table class='styled-table'>"

#$tableStart = "<table style='width:100%; border: 1px solid red'>"
Add-Content -Path $OutputFullPath -Value $tableStart
Add-Content -Path $OutputFullPath -Value $section_table_hdr_1

$campaignURL = ""
$candEndorseURL = ""
#$campaignTags = ""
$candidateName = ""
#$tagLines = "" # Campaign taglines in 1 string from Col G
$formattedTagLines = ""



#------------------------------------------
#                Main Loop
#------------------------------------------
#  Loop for all rows
#  i is ROW  j is COLUMN
#------------------------------------------
#for ($i=1; $i -le $rowMax+1; $i++) {
#for ($i=1; $i -le $rowMax+0; $i++) {
for ($i=1; $i -lt $rowMax+0; $i++) {

   Add-Content -Path $OutputFullPath -Value "<tr>"   
   $j = 1

   #------------------------------------------
   #  Loop for all columns
   #------------------------------------------
   while ($j -le $colCampaignLink - 1) {
  
      # Get Cell Text
      $cellText = $ExcelWorkSheet.Cells.Item($i+1,$j).text
      Write-Host ("row: [" + ($i+1) + "]" + " i: [" + $i + "]" + " j: [" + $j + "]"  + " cellText: [" + $cellText + "]") 
      
      if ($j -eq 1 -And  $cellText -eq "" ) {
         Write-Host ("BREAK")
         break
         }
      
      # Print the cell
      print_cell $cellText $OutputFullPath $j
      $j++
      }

   # End primary table row      
   Add-Content -Path $OutputFullPath -Value "</tr>"  

   $thisRow = $i + 1   
   $campaignURL = getCampaignURL -row $thisRow -col $colCampaignLink
   write-host ("campaignURL [" + $campaignURL + "]" + "`r`n") 
   
   $candEndorseURL = getCandEndorseURL -row $thisRow -col $colCandEndorseLink

   $formattedTagLines = getCampTags -row $thisRow -col $colTagString

   # Write secondary table row
   Add-Content -Path $OutputFullPath -Value "<tr>"
   Add-Content -Path $OutputFullPath -Value "<td>$campaignURL</td>"
   Add-Content -Path $OutputFullPath -Value "<td>$candEndorseURL</td>"
   Add-Content -Path $OutputFullPath -Value "<td colspan=3>$formattedTagLines</td>"
   #Add-Content -Path $OutputFullPath -Value "<td>&nbsp;</td>"
   Add-Content -Path $OutputFullPath -Value "</tr>"
   }

#------------------------------------------
#                Clean Up
#------------------------------------------
$tableEnd = "</table>"
Add-Content -Path $OutputFullPath -Value $tableEnd
   
# Write Section 3
Add-Content -Path $OutputFullPath -Value $section_3
   
# Quit Excel
$ExcelObj.quit()   