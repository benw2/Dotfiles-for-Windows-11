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

        Copy-Item $( $_.FullName ) "C:\Windows\Fonts"
        New-ItemProperty -Name $_.BaseName -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" -PropertyType string -Value $_.name 
    }
    else {
        Write-Host $($_.Name) already installed
    }
}