#---------------------------------------------
# GetDailyBingPicture.ps1 
# Automate and download the daily Bing image 
# to the user´s Teams background folder
# Martina Grom, @magrom, atwork.at
#  MicRiley - I modified this to add multi-day retrieval - because I kept forgetting to run this (and my scheduled task proved unreliable)
#---------------------------------------------

#----------------------
#  CONFIG VARIABLES

$numImages = 7 #the number of images to pull (7 is supposedly the max for the service, and will cause this script to look back 7 days and fill-in any gaps in retrieval.  Quicking testing in Sept2020 suggests that 8 is the max, but leaving at 7.)

#----------------------

# Use the Bing.com API. 
# The idx parameter determines the day: 0 is the current day, 1 is the previous day, etc. This goes back for max. 7 days. 
# The n parameter defines how many pictures you want to load. Usually, n=1 to get the latest picture (of today) only. 
# The mkt parameter defines the culture, like en-US, de-DE, etc.
$uri = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n="+$numImages+"&mkt=en-US"

# Get the picture metadata
$response = Invoke-WebRequest -Method Get -Uri $uri

# Extract the image content
$body = ConvertFrom-Json -InputObject $response.Content


#Iterate over each image
ForEach ($anImage in $body.images)
{
    $fileurl = "https://www.bing.com/"+$anImage.url
    $filename = $anImage.startdate+"-"+$anImage.title.Replace(" ","-").Replace("?","")+".jpg"

    # Download the picture to %APPDATA%\Microsoft\Teams\Backgrounds\Uploads
    $fileDir = $env:APPDATA+"\Microsoft\Teams\Backgrounds\Uploads\"
    $filepath = $fileDir + $filename
    Invoke-WebRequest -Method Get -Uri $fileurl -OutFile $filepath

}

# Job done. 
# You can use that script manually, as daily task, or in your Startup folder. Enjoy!



