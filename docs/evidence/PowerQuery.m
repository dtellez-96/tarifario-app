section Section1;

shared TARIFARIOS_AGENTES = let
    Origen = Folder.Files("C:\Users\laupac\OneDrive - ASSA ABLOY Group\TARIFARIOS_AGENTES"),
    #"Archivos ocultos filtrados1" = Table.SelectRows(Origen, each [Attributes]?[Hidden]? <> true),
    #"Invocar función personalizada1" = Table.AddColumn(#"Archivos ocultos filtrados1", "Transformar archivo", each #"Transformar archivo"([Content])),
    #"Columnas con nombre cambiado1" = Table.RenameColumns(#"Invocar función personalizada1", {"Name", "Source.Name"}),
    #"Otras columnas quitadas1" = Table.SelectColumns(#"Columnas con nombre cambiado1", {"Source.Name", "Transformar archivo"}),
    #"Columna de tabla expandida1" = Table.ExpandTableColumn(#"Otras columnas quitadas1", "Transformar archivo", Table.ColumnNames(#"Transformar archivo"(#"Archivo de ejemplo"))),
    #"Tipo cambiado" = Table.TransformColumnTypes(#"Columna de tabla expandida1",{{"Source.Name", type text}, {"Agente", type text}, {"Fecha_Inicio", type date}, {"Fecha_Vencimiento", type date}, {"POL", type text}, {"POD", type text}, {"VIA", type text}, {"NAVIERA", type text}, {"FLETE_20ST", Int64.Type}, {"FLETE_40ST", Int64.Type}, {"FLETE_40HC", Int64.Type}, {"DOCUMENTACION_BL", type any}, {"TT", type any}, {"DIAS_LIBRES", type any}, {"Moneda", type text}, {"Observaciones", type text}}),
    #"Columnas quitadas" = Table.RemoveColumns(#"Tipo cambiado",{"Agente"}),
    #"Personalizada agregada" = Table.AddColumn(#"Columnas quitadas", "AGENTE", each Text.Upper(Text.BeforeDelimiter([Source.Name], "."))),
    #"Columnas reordenadas" = Table.ReorderColumns(#"Personalizada agregada",{"Source.Name", "AGENTE", "Fecha_Inicio", "Fecha_Vencimiento", "POL", "POD", "VIA", "NAVIERA", "FLETE_20ST", "FLETE_40ST", "FLETE_40HC", "DOCUMENTACION_BL", "TT", "DIAS_LIBRES", "Moneda", "Observaciones"}),
    #"Personalizada agregada1" = Table.AddColumn(#"Columnas reordenadas", "Estado Vigencia", each if [Fecha_Vencimiento] = null then "VALIDAR"
else if [Fecha_Vencimiento] < Date.From(DateTime.LocalNow()) then "VENCIDA"
else if [Fecha_Vencimiento] <= Date.AddDays(Date.From(DateTime.LocalNow()), 3) then "POR VENCER"
else "VIGENTE"),
    #"Tipo cambiado1" = Table.TransformColumnTypes(#"Personalizada agregada1",{{"TT", Int64.Type}, {"AGENTE", type text}, {"FLETE_20ST", Currency.Type}, {"FLETE_40ST", Currency.Type}, {"FLETE_40HC", Currency.Type}, {"DOCUMENTACION_BL", Currency.Type}, {"DIAS_LIBRES", Int64.Type}}),
    #"Columnas quitadas1" = Table.RemoveColumns(#"Tipo cambiado1",{"Observaciones"}),
    #"Errores reemplazados" = Table.ReplaceErrorValues(#"Columnas quitadas1", {{"TT", null}}),
    #"Errores reemplazados1" = Table.ReplaceErrorValues(#"Errores reemplazados", {{"DIAS_LIBRES", null}}),
    #"Errores reemplazados2" = Table.ReplaceErrorValues(#"Errores reemplazados1", {{"DOCUMENTACION_BL", null}})
in
    #"Errores reemplazados2";

shared Parámetro1 = #"Archivo de ejemplo" meta [IsParameterQuery=true, BinaryIdentifier=#"Archivo de ejemplo", Type="Binary", IsParameterQueryRequired=true];

shared #"Transformar archivo de ejemplo" = let
    Origen = Excel.Workbook(Parámetro1, null, true),
    tblTarifario_Table = Origen{[Item="tblTarifario",Kind="Table"]}[Data]
in
    tblTarifario_Table;

shared #"Archivo de ejemplo" = let
    Origen = Folder.Files("C:\Users\laupac\OneDrive - ASSA ABLOY Group\TARIFARIOS_AGENTES"),
    Navegación1 = Origen{0}[Content]
in
    Navegación1;

[ FunctionQueryBinding = "{""exemplarFormulaName"":""Transformar archivo de ejemplo""}" ]
shared #"Transformar archivo" = let
    Origen = (Parámetro1) => let
        Origen = Excel.Workbook(Parámetro1, null, true),
        tblTarifario_Table = Origen{[Item="tblTarifario",Kind="Table"]}[Data]
    in
        tblTarifario_Table
in
    Origen;