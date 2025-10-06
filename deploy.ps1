# Deploy script for GitHub Pages
# This copies dist folder contents to gh-pages branch

Write-Host "Building project..." -ForegroundColor Green
npm run build

Write-Host "`nDeploying to gh-pages branch..." -ForegroundColor Green

# Get GitHub Desktop's git path
$gitPath = "$env:LOCALAPPDATA\GitHubDesktop\app-*\resources\app\git\cmd\git.exe"
$git = Get-Item $gitPath | Select-Object -First 1 -ExpandProperty FullName

if (-not $git) {
    Write-Host "Git not found. Please ensure GitHub Desktop is installed." -ForegroundColor Red
    exit 1
}

& $git checkout gh-pages
if ($LASTEXITCODE -ne 0) {
    Write-Host "Creating gh-pages branch..." -ForegroundColor Yellow
    & $git checkout --orphan gh-pages
    & $git rm -rf .
}

# Copy dist files
Copy-Item -Path "dist\*" -Destination "." -Recurse -Force

& $git add .
& $git commit -m "Deploy to GitHub Pages"
& $git push origin gh-pages --force

& $git checkout main

Write-Host "`nDeployment complete! Your add-in will be available at:" -ForegroundColor Green
Write-Host "https://drian-a11y.github.io/numeriq-addin/" -ForegroundColor Cyan
