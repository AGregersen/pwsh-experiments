# https://stackoverflow.com/questions/62564356/set-get-custom-file-properties-with-powershell
function Set-FileProperty {
    param(
        [string]$filename,
        [string]$propertyname,
        [string]$propertyvalue

    )
    $shell  = New-Object -COMObject Shell.Application
    $myfile = Get-Item -Path $filename
    $file   = $myfile.Name
    $path   = $myfile.DirectoryName
    
    $shellfolder = $shell.Namespace($path)
    $shellfile   = $shellfolder.ParseName($file)

    0..32767 | %{
        $value = $shellfolder.GetDetailsOf($shellfile, $_)
        if ($value) {
            write-output "$_ $value"
        }
        #if ( $value -eq 54849 ) { $_, $shellfolder.GetDetailsOf($null, $_), $shellfolder.GetDetailsOf($shellfile, $_) }
     }
}

$filename = 'Z:\Movies\Battleship.mp4'