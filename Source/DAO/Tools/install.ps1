param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("DAOApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
