param($installPath, $toolsPath, $package, $project)

$ref = $project.Object.References.Item("MSDATASRCApi")
if ($ref -and $ref.EmbedInteropTypes)
{
    $ref.EmbedInteropTypes = $false
}
