function passGenerator{

    $special = $(33..47|%{[char]$_}) + $(58..64|%{[char]$_}) +$(91..96|%{[char]$_}) + $(123..126|%{[char]$_})|Get-Random -c 2

    $number = 48..57|%{[char]$_}|Get-Random -c 3

    $upCase = 65..90|%{[char]$_}|Get-Random -c 4

    $lowCase = 97..122|%{[char]$_}|Get-Random -c 5

    $pass = $special + $number + $upCase +$lowCase
    
    return $pass -join ""
}
