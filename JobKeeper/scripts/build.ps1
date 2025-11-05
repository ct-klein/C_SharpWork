# Build script for JobKeeper
param(
    [string]$Configuration = "Release"
)

Write-Host "Building JobKeeper..." -ForegroundColor Green

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
dotnet clean ../JobKeeper.sln

# Restore packages
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
dotnet restore ../JobKeeper.sln

# Build solution
Write-Host "Building solution in $Configuration mode..." -ForegroundColor Yellow
dotnet build ../JobKeeper.sln --configuration $Configuration --no-restore

# Run tests
Write-Host "Running tests..." -ForegroundColor Yellow
dotnet test ../JobKeeper.sln --configuration $Configuration --no-build --verbosity normal

Write-Host "Build completed!" -ForegroundColor Green
