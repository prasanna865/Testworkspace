[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "Sheet1_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared Sheet1 = let
  Source = Excel.Workbook(Web.Contents("https://ecolab-my.sharepoint.com/personal/prasannareddy_vustala_ecolab_com/Documents/Desktop/EDO%20workspaces%20(Total%20list).xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data],
  #"Changed column type" = Table.TransformColumnTypes(#"Navigation 1", {{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}}),
  #"Promoted headers" = Table.PromoteHeaders(#"Changed column type", [PromoteAllScalars = true]),
  #"Removed bottom rows" = Table.RemoveLastN(#"Promoted headers", 5),
  #"Removed bottom rows 1" = Table.RemoveLastN(#"Removed bottom rows", 1)
in
  #"Removed bottom rows 1";
shared Sheet1_DataDestination = let
  Pattern = Lakehouse.Contents([CreateNavigationProperties = false, EnableFolding = false, HierarchicalNavigation = null, PreserveOrder = null]),
  Navigation_1 = Pattern{[workspaceId = "d3829ec6-95d6-498c-ac08-74d5f0854256"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "e2d5ed02-f0e3-4dab-8a65-94a9a51d00fd"]}[Data],
  TableNavigation = Navigation_2{[Id = "Sheet1", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
