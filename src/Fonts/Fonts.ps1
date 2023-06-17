Write-Host "Installing fonts.."
$Source = "*"
$Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
$TempFolder = "C:\Windows\Fonts"

# New-Item $TempFolder -Type Directory -Force | Out-Null

# Install Fonts
Get-ChildItem -Path $Source -Include '*.ttf', '*.ttc', '*.otf' -Recurse | ForEach-Object {
    If (-not(Test-Path "C:\Windows\Fonts\$( $_.Name )"))
    {
        Write-Host Installing font  $($_.BaseName) For All User
        $Font = "$TempFolder\$( $_.Name )"
        Copy-Item $( $_.FullName ) -Destination $TempFolder
        $Destination.CopyHere($Font, 0x10)
    }
    else {
        Write-Host $($_.Name) already installed
    }
}