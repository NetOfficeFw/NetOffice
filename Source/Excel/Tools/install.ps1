param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("ExcelApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
