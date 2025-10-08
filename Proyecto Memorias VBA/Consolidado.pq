let
    // Lista de nombres de tablas a anexar (sin duplicados)
    TablaNombres = Excel.CurrentWorkbook(){[Name="Tabla18"]}[Content],
    ListaNombres = List.Distinct(TablaNombres[Tb]),

    
    // Filtrar solo las tablas cuyo nombre está en la lista
    TablasFiltradas = List.Select(
        TablasDisponibles,
        each List.Contains(ListaNombres, _[Name]) and (_[Content] is table)
    ),

    // Extraer solo el contenido de las tablas válidas
    TablasValidas = List.Transform(TablasFiltradas, each _[Content]),

    // Si existe la tabla TBL_Plantilla, incluirla al principio
    TBL_Plantilla = try TablasDisponibles{[Name="TBL_Plantilla"]}[Content] otherwise null,
    TodasLasTablas = if TBL_Plantilla is table then List.Combine({ {TBL_Plantilla}, TablasValidas }) else TablasValidas,

    // Combinar todas
    Resultado = if List.Count(TodasLasTablas) > 0 then Table.Combine(TodasLasTablas) else #table({}, {})
in
    Resultado
