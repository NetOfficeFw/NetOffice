param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("ADODBApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
