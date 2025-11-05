# Run script for JobKeeper
param(
    [string]$Configuration = "Debug"
)

Write-Host "Running JobKeeper..." -ForegroundColor Green

$projectPath = "../src/JobKeeper.WinForms/JobKeeper.WinForms/JobKeeper.WinForms.csproj"

dotnet run --project $projectPath --configuration $Configuration
