Write-Host "Installing fonts.."
$Source = "*"
$Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)

# New-Item $TempFolder -Type Directory -Force | Out-Null

# Install Fonts
Get-ChildItem -Path $Source -Include '*.ttf', '*.ttc', '*.otf' -Recurse | ForEach-Object {
    If (-not(Test-Path "C:\Windows\Fonts\$( $_.Name )"))
    {
        Write-Host Installing font  $($_.BaseName) For All User

        Get-ChildItem ($_.Name) | %{ $Destination.CopyHere($_.FullName)}

        # Copy the font to the dir so we can test for it later and not install twice
        Copy-Item $( $_.FullName ) "C:\Windows\Fonts"

    }
    else {
        Write-Host $($_.Name) already installed
    }
}