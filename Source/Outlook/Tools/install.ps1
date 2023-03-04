param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("OutlookApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
