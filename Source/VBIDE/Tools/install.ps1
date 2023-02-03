param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("VBIDEApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
