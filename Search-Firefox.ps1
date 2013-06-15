ipmo $PSScriptRoot\wasp.dll

# Make sure firefox is running
if( -not (gps firefox) )
{
    start firefox
    while( -not (gps firefox) )
    {
        sleep -Milliseconds 200
    }
}

# Functions to use
function fire( $command )
{
    Select-Window firefox | Send-Keys $command
    sleep -m 200
    $command
}

function goto( $text )
{
    # NOTE: doesn't work when google asks 'did you mean ...'
    fire "{ESC}^l"
    fire "g $text{ENTER}"
    sleep 3
    fire "^fSearch tools{ENTER}{ESC}"
    fire "{TAB}{TAB}{TAB}{TAB}{TAB}{ENTER}"
}

<#

goto "schooldigger wa Bennett Elem"
goto "redfin 2035 W Lake Sammamish Pkwy SE, Bellevue 98008"

while( $true )
{
    sleep 1
    $x = [Windows.Forms.Cursor]::Position.X
    $y = [Windows.Forms.Cursor]::Position.Y
    "X = $x      Y = $y"
}
#>
