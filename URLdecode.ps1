$url = ""
if (-not ($blnTest)) { $url = Read-Host "Enter URL to Decode"}
else {$url = $StrTestUrl}

<# Script Start #>
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
$outputs = ""

<# Step 1 - MS decode #>
$safelinkRegexMS = 'https://*.safelinks.protection.outlook.com/*'

if ($($url.trim()) -like $safelinkRegexMS) {
  $urlParts = $url.split("?")[1]
  $params = $urlParts.split("&")
  $decodedurl = ""
  for ($i = 0; $i -lt $params.length; $i++) {
    $namval = $params[$i].split('=')
    if ($namval[0] -eq "url") {
      $decodedurl = $namval[1]
      $outputs = [System.Web.HttpUtility]::UrlDecode($decodedurl)
    }
  }
} else {
  # no MS match, assume just PP
  $outputs = $url
}

<# Step 1 Validation #>
if ($blnTest) {
  if ($strDecodedTestUrl1 -eq $outputs ) {  write-host "First Decode SUCCESS" -foreground Green}
  else {write-host "First Decode FAILURE `r`n`t >> $outputs" -foreground Red}
} else { }

<# Step 2 - PP decode #>
$url = $outputs
$safelinkRegexPP = 'https://urldefense.proofpoint.com*'
if ($($url.trim()) -like $safelinkRegexPP) {
  $content = ($url.Split([string[]]"?u=", [StringSplitOptions]"None")[1]).Split([string[]]"&d=", [StringSplitOptions]"None")[0]
  # $start = URL.value.indexOf("?u=") + 4
  # $end = URL.value.indexOf("&d=")

  $final = $content
  $final = $final.replace('-3A', ':').replace('_', '/').replace('-7E', '~').replace('-2560', '`').replace('-21', '!').replace('-40', '@').replace('-23', '#')
  $final = $final.replace('-24', '$').replace('-25', '%').replace('-255E', '^').replace('-26', '&').replace('-2A', '*').replace('-28', '"').replace('-29', ')')
  $final = $final.replace('-5F', '_').replace('-2B', '+').replace('-2D', '-').replace('-3D', '=').replace('-257B', '{').replace('257D', '}').replace('-257C', '|')
  # possible error on the line below for the single quotes
  $final = $final.replace('-5B', '[').replace('-5D', ']').replace('-255C', '\').replace('-26quote-3B', '''').replace('-3B', ';').replace('-26-2339-3B', '''').replace('-26lt-3B', '<')
  $final = $final.replace('-26gt-3B', '>').replace('-3F', '?').replace('-2C', ',')
  $outputs = $final
}

<# Step 2 - Validation #>
if ($blnTest) {
  if ($strDecodedTestUrl2 -eq $outputs ) {  write-host "Second Decode SUCCESS" -foreground Green}
  else {write-host "Second Decode FAILURE `r`n`t >> $outputs" -foreground Red}
} else { }

<# End Script - Output #>
write-host "`r`n`r`n`nMS-Decoded:`r`n`t$outputs"
