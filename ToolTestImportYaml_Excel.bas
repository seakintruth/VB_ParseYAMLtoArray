Attribute VB_Name = "basToolTestImportYaml"
Public Sub ToolTestImportYaml(strYamlFilePath As String, rngDestination As Range)
Dim arryYaml As Variant
Dim rngDestinationReturn As Range
    arryYaml = ParseYAMLtoArray(strYamlFilePath)
    rngDestination.Worksheet.Activate
    rngDestination.Activate
    Set rngDestinationReturn = rngDestination.Worksheet.Range( _
        rngDestination.Address, _
        rngDestination.Offset( _
            UBound(arryYaml, 1) - LBound(arryYaml, 1), _
            UBound(arryYaml, 2) - LBound(arryYaml, 2) _
        ).Address _
    )
    'Assign the values to the destination range
    rngDestinationReturn.Value = arryYaml
End Sub
