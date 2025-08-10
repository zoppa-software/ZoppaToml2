$result = dotnet test --collect:"XPlat Code Coverage"
$lastLine = $result | Select-Object -Last 1
ReportGenerator -reports:$lastLine.Trim() -targetdir:"coveragereport" -reporttypes:Html