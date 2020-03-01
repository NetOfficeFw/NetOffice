param($installPath, $toolsPath, $package, $project)

$libraries = @(
    "AccessApi",
    "ADODBApi",
    "DAOApi",
    "ExcelApi",
    "MSComctlLibApi",
    "MSDATASRCApi",
    "MSFormsApi",
    "MSHTMLApi",
    "MSProjectApi",
    "NetOffice",
    "OfficeApi",
    "OutlookApi",
    "OWC10Api",
    "PowerPointApi",
    "PublisherApi",
    "VBIDEApi",
    "VisioApi",
    "WordApi"
)

foreach ($library in $libraries) {
    $ref = $project.Object.References.Item($library)
    if ($ref -and $ref.EmbedInteropTypes)
    {
        $ref.EmbedInteropTypes = $false
    }
}
