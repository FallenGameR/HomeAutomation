# http://www.matrix.nwmls.com/Matrix/Public/Portal.aspx?L=1&k=2238933XSKD0&p=DE-26917322--F

filter get( [string] $key )
{
    $psitem | parse "$key\s*:\s+(\S+(\s\S+)*)"
}

filter schools
{
    function go( $school )
    {
        if( -not $school ) {return}

        $query = [Web.HttpUtility]::UrlEncode( "schooldigger wa $school" )
        $url = "http://www.google.com/search?btnI=1&q=$query"
        start $url
    }

    go $psitem.senior
    go $psitem.middle
    go $psitem.elementary
}

Add-Type -Assembly PresentationCore
Add-Type -AssemblyName System.Web
$text = [Windows.Clipboard]::GetText()

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

$result = construct price footage year pricePerFootage lot elementary middle senior hoa taxes style roof exterior sewer level details features cool energy heat floor appliances interior
$result | schools
$result

# better search on schools
# search on redfin, to schools
# search on google earth
# search on imap
