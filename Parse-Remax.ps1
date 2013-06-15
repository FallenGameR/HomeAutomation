# http://www.matrix.nwmls.com/Matrix/Public/Portal.aspx?L=1&k=2238933XSKD0&p=DE-26917322--F

. $PSScriptRoot\Search-Firefox.ps1

function ensure( $program )
{
    if( -not (gps $program) )
    {
        start $program
        while( -not (gps $program) )
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

filter schools
{
    function go( $school )
    {
        if( -not $school ) {return}

        $found = Invoke-BingQuery "schooldigger wa $school" | select -f 1
        start $found.Url
    }

    go $psitem.senior
    go $psitem.middle
    go $psitem.elementary
}

filter redfin
{
    $found = Invoke-BingQuery "redfin wa $($psitem.address)" | select -f 1
    start ($found.Url.Trim("#!") + "#schools")
}

filter earth
{
    Select-Window googleearth | Select-Control -Title "search_field_" -Recurse | Send-Keys "^a$($psitem.address){ENTER}"
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
$result | earth
$result | schools
$result | redfin
$result | earth

# search on imap
# search path from address to redmond building 42

<#
http://dev.virtualearth.net/REST/V1/Routes/Driving?o=xml&wp.0=london&wp.1=leeds&avoid=minimizeTolls&key=Ahg6X-c77_g5gp0YXI9hfR2ly296HlEuGe2pqSmmepHcBKtQdMXXkPT0Q6ES89SZ
http://dev.virtualearth.net/REST/V1/Imagery/Map/Road/space%20needle%20seattle?mapLayer=TrafficFlow&key=Ahg6X-c77_g5gp0YXI9hfR2ly296HlEuGe2pqSmmepHcBKtQdMXXkPT0Q6ES89SZ
http://dev.virtualearth.net/REST/v1/Imagery/Map/Road/Routes?wp.0=Seattle,WA&wp.1=Redmond,WA&key=Ahg6X-c77_g5gp0YXI9hfR2ly296HlEuGe2pqSmmepHcBKtQdMXXkPT0Q6ES89SZ
#>
