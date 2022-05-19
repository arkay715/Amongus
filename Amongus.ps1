function speak {

    [CmdletBinding()]
    param (	
    [Parameter (Mandatory = $True, ValueFromPipeline = $True)]
    [string]$Sentence
    ) 
    
    $s=New-Object -ComObject SAPI.SpVoice
    $s.Rate = -2
    $s.Speak($Sentence)
    
    }
function Set-Volume {
        Param(
            [Parameter(Position=0)]
            [string[]]
            $vol
        )
        $volume = switch ( $vol )
        {
            "max"  { 175 }
            "off" { 174 }
        }
        $k=[Math]::Ceiling(100/2);$o=New-Object -ComObject WScript.Shell;for($i = 0;$i -lt $k;$i++){$o.SendKeys([char] $volume)}
        }
        
        Function Set-WallPaper {}
 
            <#
             
                .SYNOPSIS
                Applies a specified wallpaper to the current user's desktop
                
                .PARAMETER Image
                Provide the exact path to the image
             
                .PARAMETER Style
                Provide wallpaper style (Example: Fill, Fit, Stretch, Tile, Center, or Span)
              
                .EXAMPLE
                Set-WallPaper -Image "C:\Wallpaper\Default.jpg"
                Set-WallPaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
              
            #>
            
             
            param (
                [parameter(Mandatory=$True)]
                # Provide path to image
                [string]$Image,
                # Provide wallpaper style that you would like applied
                [parameter(Mandatory=$False)]
                [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
                [string]$Style
            )
             
            $WallpaperStyle = Switch ($Style) {
              
                "Fill" {"10"}
                "Fit" {"6"}
                "Stretch" {"2"}
                "Tile" {"0"}
                "Center" {"0"}
                "Span" {"22"}
              
            }
             
            If($Style -eq "Tile") {
             
                New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
                New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force
             
            }
            Else {
             
                New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
                New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force
             
            }
             
            Add-Type -TypeDefinition @" 
            using System; 
            using System.Runtime.InteropServices;
              
            public class Params
            { 
                [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
                public static extern int SystemParametersInfo (Int32 uAction, 
                                                               Int32 uParam, 
                                                               String lpvParam, 
                                                               Int32 fuWinIni);
            }
"@

    Invoke-WebRequest https://wallpaperaccess.com/full/4707938.jpg ?dl=1 -O $env:TMP\image.jpg


    Set-Volume max
    "day mung us" | speak

