# http://www.matrix.nwmls.com/Matrix/Public/Portal.aspx?L=1&k=2238933XSKD0&p=DE-26917322--F

. $PSScriptRoot\Search-Firefox.ps1

function ensure( $program )
{
    if( -not (gps $program -ea Ignore) )
    {
        start $program
        while( -not (gps $program -ea Ignore) )
        {
            sleep -Milliseconds 200
        }
    }
}

ensure firefox
ensure googleearth

filter get( [string] $key )
{
    $psitem | parse "$key\s*:\s+(\S+(\s\S+)*)"
}

function school( $name )
{
    if( -not $name ) {return}

    $found = Invoke-BingQuery "schooldigger wa $name" | select -f 1
    start $found.Url
}

function redfin( $address )
{
    $found = Invoke-BingQuery "redfin wa $address" | select -f 1
    start ($found.Url.Trim("#!") + "#schools")
}

function earth( $address )
{
    $window = select-window google* | Set-WindowActive
    sleep -Milliseconds 200
    $control = $window | Select-Control -Title "search_field_" -Recurse
    sleep -Milliseconds 200
    $control | Send-Keys "^a$addres){ENTER}"
}

function Search-Map( $text )
{
    # Encode text of the query
    $query = "http://dev.virtualearth.net/REST/V1/Imagery/Map/Road/"
    $query += [Uri]::EscapeDataString( $text )
    $query += "?mapLayer=TrafficFlow&key=Ahg6X-c77_g5gp0YXI9hfR2ly296HlEuGe2pqSmmepHcBKtQdMXXkPT0Q6ES89SZ"

    # Download the picture
    start $query
}

function Get-DriveDuration( $from = "redmond,wa", $to = "Microsoft Building 42, WA" )
{
    $query = "http://dev.virtualearth.net/REST/V1/Routes/Driving?o=xml"
    $query += "&wp.0=" + [Uri]::EscapeDataString($from)
    $query += "&wp.1=" + [Uri]::EscapeDataString($to)
    $query += "&key=Ahg6X-c77_g5gp0YXI9hfR2ly296HlEuGe2pqSmmepHcBKtQdMXXkPT0Q6ES89SZ"

    $xml = [xml] (Download-Url $query)
    $seconds = $xml.Response.ResourceSets.ResourceSet.Resources.Route.TravelDuration
    [timespan]::FromSeconds($seconds).ToString()
}

function Get-DriveMap( $from = "redmond,wa", $to = "Microsoft Building 42, WA" )
{
    $query = "http://dev.virtualearth.net/REST/v1/Imagery/Map/Road/Routes"
    $query += "?wp.0=" + [Uri]::EscapeDataString($from)
    $query += "&wp.1=" + [Uri]::EscapeDataString($to)
    $query += "&key=Ahg6X-c77_g5gp0YXI9hfR2ly296HlEuGe2pqSmmepHcBKtQdMXXkPT0Q6ES89SZ"

    start $query
}

function Invoke-BingQuery
{
    param
    (
        [Parameter(Mandatory)]
        [string] $Query,
        [switch] $Raw
    )

    $env:BingAPIKey = "Qjkn4cg7ZPjqoY25ZmHa2L61zz1mScshcv9fOOxyCbI"
    $bingUrl = "https://api.datamarket.azure.com/Bing/Search/Web?Query='$($Query)'&`$format=JSON"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes("$($Env:BingAPIKey):$($Env:BingAPIKey)")
    $authorisation = "Basic {0}" -f [System.Convert]::ToBase64String($bytes)
    $result = (Invoke-RestMethod -Uri $BingUrl -Headers @{Authorization=$authorisation})

    if( $raw )
    {
        $result
    }
    else
    {
        $result.d.results
    }
}


Add-Type -Assembly PresentationCore
Add-Type -AssemblyName System.Web
$text = [Windows.Clipboard]::GetText()

$address = ($text -split "`r?`n")[20]
$mls = $text | get 'MLS#'
$status = $text | get 'Status'
$county = $text | get 'County'
$community = $text | get 'Commty'

$price = $text | get 'LP'
$footage = $text | get 'SF'
$year = $text | get 'Yr Built'
$pricePerFootage = $text | get 'Price/SF'
$lot = $text | get 'Lot Size'
$elementary = $text | get 'Elementary'
$middle = $text | get 'Jr High'
$senior = $text | get 'Snr High'
$hoa = $text | get 'Mnth Dues'

$taxes = $text | get 'Ann Taxes'
$style = $text | get 'Style Code'
$roof = $text | get 'Roof'
$exterior = $text | get 'Exterior'
$sewer = $text | get 'Sewer'
$level = $text | get 'Lot Top/Veg'
$details = $text | get 'Lot Dtls'
$features = $text | get 'Site Feat'

$cool = $text | get 'Heat/Cool'
$energy = $text | get 'Energy'
$heat = $text | get 'Wtr Heatr Ty/Loc'
$floor = $text | get 'Floor Cvr'
$appliances = $text | get 'Appliances'
$interior = $text | get 'Interior Ft'

$result = construct address mls status county community price footage year pricePerFootage lot elementary middle senior hoa taxes style roof exterior sewer level details features cool energy heat floor appliances interior
$result
Get-DriveDuration $result.address
Get-DriveMap $result.address

earth $result.address
school $result.senior
school $result.middle
school $result.elementary
redfin $result.address

# search on imap
