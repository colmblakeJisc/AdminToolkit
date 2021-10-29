Connect-SPOService -Url https://jiscdemo-admin.sharepoint.com/
$themepalette =@{
    "themePrimary" = "#ae0001";
    "themeLighterAlt" = "#fcf2f2";
    "themeLighter" = "#f2cbcb";
    "themeLight" = "#e7a1a1";
    "themeTertiary" = "#ce5252";
    "themeSecondary" = "#b71616";
    "themeDarkAlt" = "#9c0000";
    "themeDark" = "#840000";
    "themeDarker" = "#610000";
    "neutralLighterAlt" = "#faf9f8";
    "neutralLighter" = "#f3f2f1";
    "neutralLight" = "#edebe9";
    "neutralQuaternaryAlt" = "#e1dfdd";
    "neutralQuaternary" = "#d0d0d0";
    "neutralTertiaryAlt" = "#c8c6c4";
    "neutralTertiary" = "#a19f9d";
    "neutralSecondary" = "#605e5c";
    "neutralPrimaryAlt" = "#3b3a39";
    "neutralPrimary" = "#323130";
    "neutralDark" = "#201f1e";
    "black" = "#000000";
    "white" = "#ffffff";
    }
Add-SPOTheme -Identity "Hogwarts" -Palette $themepalette -IsInverted $false -Overwrite
Set-SPOHideDefaultThemes $true