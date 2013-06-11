<#
Prop Type: 	Single Family 	Sub Prop: 	Residential
Beds: 	4 	Baths: 	2.5 	SF: 	2,480 	SF Source: 	KCR
Yr Built: 	1984 	Price/SF: 	$201.61 	Lot Size: 	.341 ac/14,873 sf
Elementary: 	Blackwell Elem 	Jr High: 	Inglewood Jnr High
Snr High: 	Eastlake High 		School D: 	Lake Washington
Map Book: 	Thomas Brothers 	Gd/Map: 	H1, 567
Mnth Dues: 	$9 	H/O Incl:
Mntly Rent: 		Cat/Dog:
Directions:
228th/Sahalee Way to East on NE 37th Way, right on 208th Ave NE, right on NE 42nd Street, right on 211th Ct NE.

Ann Taxes: 	$5,017 	Tax Year: 	2012 	Snr Expt: 	No 	Form 17: 	Provided
Ttl Cvr Prk: 	2 	Prk Spc: 		Prk Typ:
Garage-Attached
Style Code: 	12 - 2 Story 	Bld Nm:
Bld Cond: 		Bld Info:
Built On Lot
Bsmnt: 	None 	Roof: 	Cedar Shake
Foundation: 		Exterior:
Wood
Sewer: 	Sewer Connected 	First Refusal: 	No 	Seller's Concess:
Lot Top/Veg:
Level, Partial Slope
Lot Dtls:
Cul-de-sac, Curbs, Paved Street
Site Feat:
Cable TV, Deck, Fenced-Partially


Main Beds: 	0 	M 1/2 Baths: 	1 	M 3/4 Baths: 	0 	M F Baths: 	0
Upper Beds: 	4 	U 1/2 Baths: 	0 	U 3/4 Baths: 	0 	U F Baths: 	2
Lw Beds: 	0 	L 1/2 Baths: 	0 	L 3/4 Baths: 	0 	L F Baths: 	0
Ttl Beds: 	4 	Ttl 1/2 Baths: 	1 	Ttl 3/4 Baths: 	0 	Ttl F Baths: 	2
Ttl Baths: 	2.5 	M Fplc: 		Ttl Fplc: 	1
Heat/Cool: 	Forced Air 	Energy: 	Natural Gas
Wtr Heatr Ty/Loc: 	Gas / Garage
Floor Cvr:
Hardwood, Vinyl, Wall to Wall Carpet
Appliances:
Dishwasher, Dryer, Garbage Disposal, Microwave, Range/Oven, Refrigerator, Washer
Interior Ft:
Bath Off Master, Dbl Pane/Storm Windw, Dining Room, Skylights, Walk-in Closet

#>

filter get( [string] $key )
{
    $psitem | parse "$key\s*:\s+(\S+(\s\S+)*)"
}

Add-Type -Assembly PresentationCore
$text = [Windows.Clipboard]::GetText()
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

