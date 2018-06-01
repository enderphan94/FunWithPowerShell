function passGenerator{
    $Length = 35
 
    $PasswordCharCodes = {33..126}.invoke()
 
    #Exclude ",',/,`,O,0
    34,39,47,96,48,79 | foreach {[void]$PasswordCharCodes.Remove($_)}
 
    $PasswordChars = [char[]]$PasswordCharCodes
 
    do { 
        $NewPassWord =  $(foreach ($i in 1..$length) 
        { Get-Random -InputObject $PassWordChars }) -join '' 
    }
 
    until (
        ( $NewPassword -cmatch '[A-Z]' ) -and
        ( $NewPassWord -cmatch '[a-z]' ) -and
        ( $NewPassWord -imatch '[0-9]' ) -and 
        ( $NewPassWord -imatch '[^A-Z0-9]' )
    ) 
          
    return $NewPassword
}
passGenerator
