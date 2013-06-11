# http://www.matrix.nwmls.com/Matrix/Public/Portal.aspx?L=1&k=2238933XSKD0&p=DE-26917322--F

filter get( [string] $key )
{
    $psitem | parse "$key\s*:\s+(\S+(\s\S+)*)"
}

Add-Type -Assembly PresentationCore
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


