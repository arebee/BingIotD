# Get the bing image of the day.

$root_url = 'http://www.bing.com'
$dst_dir = Join-Path $HOME '/Documents/Pictures/deskFeed/'
if (!(Test-Path $dst_dir -PathType Container)) {
    new-item -Path $dst_dir -ItemType Container | Out-Null
}
$validFileTypes = ("jpg","jpeg","png")
$targetRes = "1920x1080" # Retina MacBook Pro 2015
[Int32]$numberOfImagesToFetch = 20
[Int32]$urlPageSize = 8
[Int32]$page = 0
[Int32]$remainder = 0
[Int32]$imgCount = 0
$envProgressShow = $ProgressPreference

# Hide progress bars
$ProgressPreference = 'SilentlyContinue'

# This script needs to be online to work correctly. If we can't reach the internet
# quit after writing an error.

[boolean]$netState = Test-NetConnection -Port 80 -InformationLevel Quiet
Write-Verbose "Net connectionstate = $netState"
if ($netState -eq $False) {
    Write-Error 'Unable to connect to the network'
    return
}

# Main script logic

if ((Test-Path $dst_dir) -eq $False) {
    New-Item -Path $dst_dir -ItemType Directory 
}

if ($numberOfImagesToFetch -gt 120) {
    $numberOfImagesToFetch = 20
} # arbitary
if ($numberOfImagesToFetch -lt 1) {
    $numberOfImagesToFetch = 1
}

$imageNodeCount = 0

[xml]$rss_feed = ''
# :loopBegin while ($imgCount -lt $numberOfImagesToFetch) {
# $imageNodeCount++
# $imgNodeCount is the number of //image items that have been fetched
:loopBegin while ($imgNodeCount -lt $numberOfImagesToFetch) {
    Write-Verbose $page
    $xmlUri = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=$($page * $urlPageSize)&n=8&mkt=en-US"
    Write-Verbose $xmlUri
    [xml]$rss_feed_request = Invoke-Webrequest -Uri $xmlUri
    if ($rss_feed.HasChildNodes -eq $False) {
        $rss_feed.Load($xmlUri) | Out-Null
        $imgNodeCount = (Select-Xml '//image' $rss_feed).Count
        Write-Verbose $imgNodeCount
    }
    else {
        $images = Select-Xml '//image' $rss_feed_request
        if ($images.Count -eq 0) {
            # $imgCount = $numberOfImagesToFetch
            # $imageNodeCount++
            # $imgNodeCount = $numberOfImagesToFetch
            $page++ | out-null
            break loopBegin
        }
        foreach ($imageNode in $images) {
            if ($imgNodeCount -gt $numberOfImagesToFetch) {
                break
            }
            $importNode = $rss_feed.ImportNode($imageNode.Node, $True)
            $rss_feed.DocumentElement.AppendChild($importNode) | Out-Null
            $imgNodeCount++
        }
    }
    $page++ | out-Null
}

write-Host $(Select-Xml '//image' $rss_feed).Count


#Clear-Host # bug in Mac PowerShell - shows progress even when assigned.
# Write-Verbose "$((Select-Xml '//image' $rss_feed).Count) items"

foreach ($image in (Select-Xml '//image' $rss_feed)) {
    [string]$urlBase = $image.Node.Item('urlBase').InnerText
    [string]$url = $image.Node.Item('url').InnerText
    [string]$type = $url.Substring($url.LastIndexOf('.')+1)
    [string]$itemDateFileName = $image.Node.Item('startdate').InnerText -replace "(\d\d\d\d)(\d\d)(\d\d)",'$1-$2-$3'
    if (($urlBase -ne $null) -and ($itemDateFileName -ne $null) -and ($type -in $validFileTypes)) {
        Write-Verbose "Valid"
        $itemDateFileName = $itemDateFileName + '.' + $type.tolower()
        # Write-Verbose "$urlBase`n$url`n$itemDateFileName`n$type"
        $outFileName = join-path $dst_dir $itemDateFileName
        # Write-Verbose $outFileName
        if ((Test-Path -Path $outFileName -PathType Leaf) -eq $False) {
            $imageUrl = $root_url + $urlBase + "_$targetRes." + $type
            Write-Verbose "Fetching $imageUrl"
            Invoke-Webrequest -Uri $imageUrl -OutFile $outFileName | Out-Null
        }
        else {
            Write-Verbose "$outFileName exists already"
        }
    }
}

# Reset the $ProgressPreference var
$ProgressPreference = $envProgressShow