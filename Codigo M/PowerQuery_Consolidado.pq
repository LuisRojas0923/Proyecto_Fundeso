let
    Origen = EXPORTE_PRESUPUESTO,
    #"Solo columnas numericas" = Table.SelectColumns(Origen,{"CANTIDAD", "VR. UNITARIO", "VR. PARCIAL", "1 CORTE", "1 CORTE %", "1 CORTE VALOR", "2 CORTE", "2 CORTE %", "2 CORTE VALOR", "3 CORTE", "3 CORTE %", "3 CORTE VALOR", "4 CORTE", "4 CORTE %", "4 CORTE VALOR", "TOTAL CANTIDAD EJECUTADA", "TOTAL % EJECUTADO", "TOTAL VALOR EJECUTADO", "CANTIDAD DISPONIBLE"}),
    
    // Agrupar y sumar todas las columnas en una sola operación (más eficiente)
    TOTAL = Table.Group(#"Solo columnas numericas", {}, {
        {"Concepto", each "COSTO DIRECTO", type text},
        {"TOTAL", each "TOTAL", type text},
        {"CANTIDAD", each null},
        {"VR. UNITARIO", each List.Sum([#"VR. UNITARIO"]), type number},
        
        {"VR. PARCIAL", each List.Sum([#"VR. PARCIAL"]), type number},
      //  {"1 CORTE", each List.Sum([#"1 CORTE"]), type number},
       // {"1 CORTE %", each List.Sum([#"1 CORTE %"]), type number},
        {"1 CORTE VALOR", each List.Sum([#"1 CORTE VALOR"]), type number},
      //  {"2 CORTE", each List.Sum([#"2 CORTE"]), type number},
       // {"2 CORTE %", each List.Sum([#"2 CORTE %"]), type number},
        {"2 CORTE VALOR", each List.Sum([#"2 CORTE VALOR"]), type number},
     //   {"3 CORTE", each List.Sum([#"3 CORTE"]), type number},
     //   {"3 CORTE %", each List.Sum([#"3 CORTE %"]), type number},
        {"3 CORTE VALOR", each List.Sum([#"3 CORTE VALOR"]), type number},
      //  {"4 CORTE", each List.Sum([#"4 CORTE"]), type number},
      //  {"4 CORTE %", each List.Sum([#"4 CORTE %"]), type number},
        {"4 CORTE VALOR", each List.Sum([#"4 CORTE VALOR"]), type number},
      // {"TOTAL CANTIDAD EJECUTADA", each List.Sum([#"TOTAL CANTIDAD EJECUTADA"]), type number},
      //  {"TOTAL % EJECUTADO", each List.Sum([#"TOTAL % EJECUTADO"]), type number},
        {"TOTAL VALOR EJECUTADO", each List.Sum([#"TOTAL VALOR EJECUTADO"]), type number}
      //  {"CANTIDAD DISPONIBLE", each List.Sum([#"CANTIDAD DISPONIBLE"]), type number}
    }),
    #"Consultas combinadas" = Table.NestedJoin(TOTAL, {"TOTAL"}, Porcentajes, {"Personalizado"}, "Porcentajes", JoinKind.LeftOuter),
    #"Se expandió Porcentajes" = Table.ExpandTableColumn(#"Consultas combinadas", "Porcentajes", {"ORDEN", "CONCEPTO", "% DISTRIBUCION"}, {"ORDEN", "CONCEPTO.1", "% DISTRIBUCION"}),
    #"Columnas reordenadas" = Table.ReorderColumns(#"Se expandió Porcentajes",{"CONCEPTO.1", "% DISTRIBUCION", "Concepto", "TOTAL", "VR. UNITARIO", "VR. PARCIAL", "1 CORTE VALOR", "2 CORTE VALOR", "3 CORTE VALOR", "4 CORTE VALOR", "TOTAL VALOR EJECUTADO"}),
    #"Filas ordenadas" = Table.Sort(#"Columnas reordenadas",{{"CONCEPTO.1", Order.Ascending}}),
    #"Columnas con nombre cambiado" = Table.RenameColumns(#"Filas ordenadas",{{"VR. PARCIAL", "VR. PARCIAL1"}}),
    #"Columnas quitadas" = Table.RemoveColumns(#"Columnas con nombre cambiado",{"Concepto", "TOTAL"}),
    #"Multiplicación insertada" = Table.AddColumn(#"Columnas quitadas", "VR. PARCIAL", each [VR. PARCIAL1] * [#"% DISTRIBUCION"], type number),
    #"Columnas quitadas2" = Table.RemoveColumns(#"Multiplicación insertada",{"VR. UNITARIO"}),
    #"Personalizada agregada" = Table.AddColumn(#"Columnas quitadas2", "VR. UNITARIO", each null),
    #"Multiplicación insertada4" = Table.AddColumn(#"Personalizada agregada", "VR. PARCIAL ACTA 4", each [#"% DISTRIBUCION"] * [4 CORTE VALOR], type number),
    #"Multiplicación insertada3" = Table.AddColumn(#"Multiplicación insertada4", "VR. PARCIAL ACTA 3", each [#"% DISTRIBUCION"] * [3 CORTE VALOR], type number),
    #"Multiplicación insertada2" = Table.AddColumn(#"Multiplicación insertada3", "VR. PARCIAL ACTA 2", each [#"% DISTRIBUCION"] * [2 CORTE VALOR], type number),
    #"Multiplicación insertada1" = Table.AddColumn(#"Multiplicación insertada2", "VR. PARCIAL ACTA 1", each [1 CORTE VALOR] * [#"% DISTRIBUCION"], type number),
    #"Columnas quitadas1" = Table.RemoveColumns(#"Multiplicación insertada1",{"TOTAL VALOR EJECUTADO"}),
    #"Suma insertada" = Table.AddColumn(#"Columnas quitadas1", "TOTAL VALOR EJECUTADO", each List.Sum({[VR. PARCIAL ACTA 4], [VR. PARCIAL ACTA 3], [VR. PARCIAL ACTA 2], [VR. PARCIAL ACTA 1]}), type number),
    
    #"% AVANCE ACTA 1" = Table.AddColumn(
        #"Suma insertada", 
        "% AVANCE ACTA 1", 
        each 
            if ([CONCEPTO.1] = "COSTOS DIRECTO" or 
                [CONCEPTO.1] = "ADMINISTRACION" or 
                [CONCEPTO.1] = "IMPREVISTOS" or 
                [CONCEPTO.1] = "UTILIDAD" or 
                [CONCEPTO.1] = "TOTAL AIU") and 
               [VR. PARCIAL] <> null and 
               [VR. PARCIAL] <> 0 and 
               [VR. PARCIAL ACTA 1] <> null 
            then [VR. PARCIAL ACTA 1] / [VR. PARCIAL]
            else null,
        type number
    ),
    
    #"% AVANCE ACTA 2" = Table.AddColumn(
        #"% AVANCE ACTA 1", 
        "% AVANCE ACTA 2", 
        each 
            if ([CONCEPTO.1] = "COSTOS DIRECTO" or 
                [CONCEPTO.1] = "ADMINISTRACION" or 
                [CONCEPTO.1] = "IMPREVISTOS" or 
                [CONCEPTO.1] = "UTILIDAD" or 
                [CONCEPTO.1] = "TOTAL AIU") and 
               [VR. PARCIAL] <> null and 
               [VR. PARCIAL] <> 0 and 
               [VR. PARCIAL ACTA 2] <> null 
            then [VR. PARCIAL ACTA 2] / [VR. PARCIAL]
            else null,
        type number
    ),
    
    #"% AVANCE ACTA 3" = Table.AddColumn(
        #"% AVANCE ACTA 2", 
        "% AVANCE ACTA 3", 
        each 
            if ([CONCEPTO.1] = "COSTOS DIRECTO" or 
                [CONCEPTO.1] = "ADMINISTRACION" or 
                [CONCEPTO.1] = "IMPREVISTOS" or 
                [CONCEPTO.1] = "UTILIDAD" or 
                [CONCEPTO.1] = "TOTAL AIU") and 
               [VR. PARCIAL] <> null and 
               [VR. PARCIAL] <> 0 and 
               [VR. PARCIAL ACTA 3] <> null 
            then [VR. PARCIAL ACTA 3] / [VR. PARCIAL]
            else null,
        type number
    ),
    
    #"% AVANCE ACTA 4" = Table.AddColumn(
        #"% AVANCE ACTA 3", 
        "% AVANCE ACTA 4", 
        each 
            if ([CONCEPTO.1] = "COSTOS DIRECTO" or 
                [CONCEPTO.1] = "ADMINISTRACION" or 
                [CONCEPTO.1] = "IMPREVISTOS" or 
                [CONCEPTO.1] = "UTILIDAD" or 
                [CONCEPTO.1] = "TOTAL AIU") and 
               [VR. PARCIAL] <> null and 
               [VR. PARCIAL] <> 0 and 
               [VR. PARCIAL ACTA 4] <> null 
            then [VR. PARCIAL ACTA 4] / [VR. PARCIAL]
            else null,
        type number
    ),
    
    #"% AVANCE ACTA ACUMULADO" = Table.AddColumn(
        #"% AVANCE ACTA 4", 
        "% AVANCE ACTA ACUMULADO", 
        each 
            if ([CONCEPTO.1] = "COSTOS DIRECTO" or 
                [CONCEPTO.1] = "ADMINISTRACION" or 
                [CONCEPTO.1] = "IMPREVISTOS" or 
                [CONCEPTO.1] = "UTILIDAD" or 
                [CONCEPTO.1] = "TOTAL AIU") and 
               [VR. PARCIAL] <> null and 
               [VR. PARCIAL] <> 0 
            then [TOTAL VALOR EJECUTADO] / [VR. PARCIAL]
            else null,
        type number
    ),
    
    // Calcular TOTAL AIU como subtotal de ADMINISTRACION, IMPREVISTOS y UTILIDAD
    #"TOTAL AIU Calculado" = Table.AddColumn(
        #"% AVANCE ACTA ACUMULADO", 
        "TOTAL AIU VR. PARCIAL", 
        each 
            if [CONCEPTO.1] = "TOTAL AIU" then
                let
                    adminRow = Table.SelectRows(#"% AVANCE ACTA ACUMULADO", each [CONCEPTO.1] = "ADMINISTRACION"),
                    imprevRow = Table.SelectRows(#"% AVANCE ACTA ACUMULADO", each [CONCEPTO.1] = "IMPREVISTOS"),
                    utilRow = Table.SelectRows(#"% AVANCE ACTA ACUMULADO", each [CONCEPTO.1] = "UTILIDAD"),
                    adminVal = if Table.RowCount(adminRow) > 0 then adminRow{0}[VR. PARCIAL] else 0,
                    imprevVal = if Table.RowCount(imprevRow) > 0 then imprevRow{0}[VR. PARCIAL] else 0,
                    utilVal = if Table.RowCount(utilRow) > 0 then utilRow{0}[VR. PARCIAL] else 0
                in
                    adminVal + imprevVal + utilVal
            else [VR. PARCIAL],
        type number
    ),
    
    #"TOTAL AIU ACTA 1" = Table.AddColumn(
        #"TOTAL AIU Calculado", 
        "TOTAL AIU ACTA 1", 
        each 
            if [CONCEPTO.1] = "TOTAL AIU" then
                let
                    adminRow = Table.SelectRows(#"TOTAL AIU Calculado", each [CONCEPTO.1] = "ADMINISTRACION"),
                    imprevRow = Table.SelectRows(#"TOTAL AIU Calculado", each [CONCEPTO.1] = "IMPREVISTOS"),
                    utilRow = Table.SelectRows(#"TOTAL AIU Calculado", each [CONCEPTO.1] = "UTILIDAD"),
                    adminVal = if Table.RowCount(adminRow) > 0 then adminRow{0}[VR. PARCIAL ACTA 1] else 0,
                    imprevVal = if Table.RowCount(imprevRow) > 0 then imprevRow{0}[VR. PARCIAL ACTA 1] else 0,
                    utilVal = if Table.RowCount(utilRow) > 0 then utilRow{0}[VR. PARCIAL ACTA 1] else 0
                in
                    adminVal + imprevVal + utilVal
            else [VR. PARCIAL ACTA 1],
        type number
    ),
    
    #"TOTAL AIU ACTA 2" = Table.AddColumn(
        #"TOTAL AIU ACTA 1", 
        "TOTAL AIU ACTA 2", 
        each 
            if [CONCEPTO.1] = "TOTAL AIU" then
                let
                    adminRow = Table.SelectRows(#"TOTAL AIU ACTA 1", each [CONCEPTO.1] = "ADMINISTRACION"),
                    imprevRow = Table.SelectRows(#"TOTAL AIU ACTA 1", each [CONCEPTO.1] = "IMPREVISTOS"),
                    utilRow = Table.SelectRows(#"TOTAL AIU ACTA 1", each [CONCEPTO.1] = "UTILIDAD"),
                    adminVal = if Table.RowCount(adminRow) > 0 then adminRow{0}[VR. PARCIAL ACTA 2] else 0,
                    imprevVal = if Table.RowCount(imprevRow) > 0 then imprevRow{0}[VR. PARCIAL ACTA 2] else 0,
                    utilVal = if Table.RowCount(utilRow) > 0 then utilRow{0}[VR. PARCIAL ACTA 2] else 0
                in
                    adminVal + imprevVal + utilVal
            else [VR. PARCIAL ACTA 2],
        type number
    ),
    
    #"TOTAL AIU ACTA 3" = Table.AddColumn(
        #"TOTAL AIU ACTA 2", 
        "TOTAL AIU ACTA 3", 
        each 
            if [CONCEPTO.1] = "TOTAL AIU" then
                let
                    adminRow = Table.SelectRows(#"TOTAL AIU ACTA 2", each [CONCEPTO.1] = "ADMINISTRACION"),
                    imprevRow = Table.SelectRows(#"TOTAL AIU ACTA 2", each [CONCEPTO.1] = "IMPREVISTOS"),
                    utilRow = Table.SelectRows(#"TOTAL AIU ACTA 2", each [CONCEPTO.1] = "UTILIDAD"),
                    adminVal = if Table.RowCount(adminRow) > 0 then adminRow{0}[VR. PARCIAL ACTA 3] else 0,
                    imprevVal = if Table.RowCount(imprevRow) > 0 then imprevRow{0}[VR. PARCIAL ACTA 3] else 0,
                    utilVal = if Table.RowCount(utilRow) > 0 then utilRow{0}[VR. PARCIAL ACTA 3] else 0
                in
                    adminVal + imprevVal + utilVal
            else [VR. PARCIAL ACTA 3],
        type number
    ),
    
    #"TOTAL AIU ACTA 4" = Table.AddColumn(
        #"TOTAL AIU ACTA 3", 
        "TOTAL AIU ACTA 4", 
        each 
            if [CONCEPTO.1] = "TOTAL AIU" then
                let
                    adminRow = Table.SelectRows(#"TOTAL AIU ACTA 3", each [CONCEPTO.1] = "ADMINISTRACION"),
                    imprevRow = Table.SelectRows(#"TOTAL AIU ACTA 3", each [CONCEPTO.1] = "IMPREVISTOS"),
                    utilRow = Table.SelectRows(#"TOTAL AIU ACTA 3", each [CONCEPTO.1] = "UTILIDAD"),
                    adminVal = if Table.RowCount(adminRow) > 0 then adminRow{0}[VR. PARCIAL ACTA 4] else 0,
                    imprevVal = if Table.RowCount(imprevRow) > 0 then imprevRow{0}[VR. PARCIAL ACTA 4] else 0,
                    utilVal = if Table.RowCount(utilRow) > 0 then utilRow{0}[VR. PARCIAL ACTA 4] else 0
                in
                    adminVal + imprevVal + utilVal
            else [VR. PARCIAL ACTA 4],
        type number
    ),
    
    #"TOTAL AIU EJECUTADO" = Table.AddColumn(
        #"TOTAL AIU ACTA 4", 
        "TOTAL AIU EJECUTADO", 
        each 
            if [CONCEPTO.1] = "TOTAL AIU" then
                let
                    adminRow = Table.SelectRows(#"TOTAL AIU ACTA 4", each [CONCEPTO.1] = "ADMINISTRACION"),
                    imprevRow = Table.SelectRows(#"TOTAL AIU ACTA 4", each [CONCEPTO.1] = "IMPREVISTOS"),
                    utilRow = Table.SelectRows(#"TOTAL AIU ACTA 4", each [CONCEPTO.1] = "UTILIDAD"),
                    adminVal = if Table.RowCount(adminRow) > 0 then adminRow{0}[TOTAL VALOR EJECUTADO] else 0,
                    imprevVal = if Table.RowCount(imprevRow) > 0 then imprevRow{0}[TOTAL VALOR EJECUTADO] else 0,
                    utilVal = if Table.RowCount(utilRow) > 0 then utilRow{0}[TOTAL VALOR EJECUTADO] else 0
                in
                    adminVal + imprevVal + utilVal
            else [TOTAL VALOR EJECUTADO],
        type number
    ),
    
    #"Filas ordenadas1" = Table.Sort(#"TOTAL AIU EJECUTADO",{{"ORDEN", Order.Ascending}}),
    
    #"Seleccion final columnas" = Table.SelectColumns(#"Filas ordenadas1", {
        
        "CONCEPTO.1", 
        "CANTIDAD",
        "% DISTRIBUCION", 
        "VR. UNITARIO", 
        "VR. PARCIAL", 
        "% AVANCE ACTA 1",
        "VR. PARCIAL ACTA 1", 
        "% AVANCE ACTA 2",
        "VR. PARCIAL ACTA 2", 
        "% AVANCE ACTA 3",
        "VR. PARCIAL ACTA 3", 
        "% AVANCE ACTA 4",
        "VR. PARCIAL ACTA 4", 
        "% AVANCE ACTA ACUMULADO",
        "TOTAL VALOR EJECUTADO"
    })
in
    #"Seleccion final columnas"