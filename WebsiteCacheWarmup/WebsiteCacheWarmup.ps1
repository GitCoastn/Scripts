param (
    [string]$sitemapUrl = "https://your-domain.com/sitemap.xml"
)

Write-Host "Website Cache Warmup"
Write-Host "********************"
Write-Host ""

Write-Host "Requesting Sitemap: " -NoNewline
$userAgent = [Microsoft.PowerShell.Commands.PSUserAgent]::Chrome
$sitemapReturn = Invoke-WebRequest -Uri $sitemapUrl -UserAgent $userAgent

if ($sitemapReturn.StatusCode -eq 200){
    Write-Host "done." -ForegroundColor Green
    Write-Host ""
    $sitemapContent = $sitemapReturn.Content
    $allUrls = $sitemapContent | Select-String -Pattern 'https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)' -AllMatches `
                               | % { $_.Matches } `
                               | % { $_.Value } `
                               | Sort-Object `
                               | Get-Unique

    Write-Host "Going through each site in sitemap:"

    foreach ($url in $allUrls){
        Write-Host "  - $($url): " -NoNewline
        $urlStatusCode = (Invoke-WebRequest -Uri $url -UserAgent $userAgent).StatusCode
        Write-Host "$($urlStatusCode)"
    }
}
