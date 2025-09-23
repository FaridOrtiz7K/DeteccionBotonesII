#Persistent
SetTitleMatchMode, 2  ; Partial window title matching

; Create the GUI
Gui, Add, Text,, Inicia en:          ; Label for counter start
Gui, Add, Edit, vStartCount w50, 1   ; Default: Start at 1
Gui, Add, Text,, Número de lotes:    ; Label for loop count
Gui, Add, Edit, vLoopCount w50, 589  ; Default: 589 loops
Gui, Add, Button, gStartScript, Iniciar  ; Start button
Gui, Show,, Configuración del Script  ; Window title
return

; Start button action
StartScript:
    Gui, Submit, NoHide  ; Save user input
    MsgBox, 4, Confirmar, ¿Iniciar script desde %StartCount% con %LoopCount% lotes? (Presione Yes para comenzar)
    IfMsgBox, No
        return
    
    Sleep, 5000  ; Wait 5 seconds
    
    ; Read CSV (adjust path as needed)
    FileRead, csvContent, C:\Users\cmf05\Documents\AutoHotkey\Lotes SIGP with Nums & NSE.csv
    if (csvContent = "") {
        MsgBox, Error: El archivo CSV no se pudo leer.
        return
    }
    
    lines := StrSplit(csvContent, "`n")
    if (lines.MaxIndex() < 1) {
        MsgBox, Error: CSV vacío o con formato incorrecto.
        return
    }

    Cont := StartCount  ; Initialize counter with user input
    Loop, %LoopCount%   ; Loop for user-specified count
    {
        if GetKeyState("Esc", "P") {
            MsgBox, Script detenido por el usuario.
            break
        }

; Parse the current line (row) from the CSV
        currentLine := lines[Cont + 1]  ; +1 to skip the header row
        columns := StrSplit(currentLine, ",")  ; Split by comma (adjust delimiter if needed)

        ; Ensure there are enough columns
        if (columns.MaxIndex() < 5)  ; At least 5 columns (B, D, E, F, etc.)
        {
            MsgBox, Not enough columns in the CSV file!
            return
        }

        ; Left click at (89, 263) - Select in the list
        Click, 89, 263
        Sleep, 1500

        ; Left click at (1483, 519) - Case numero
        Click, 1483, 519
        Sleep, 1500

        ; Press delete key
        Send, {Delete}
        Sleep, 1000

        ; Write the value from column B (index 2)
        Send, % columns[2]
        Sleep, 1500

        ; Check the value in column D (index 4)
        if (columns[4] > 0)
        {
            ; Select the point called USO at (1507, 636)
            Click, 1507, 650
            Sleep, 2500

            ; Send the number of clicks down based on column D
            Loop, % columns[4]
            {
                Send, {Down}
                Sleep, 2000
            }
        }

        ; Press Actualizar at (1290, 349)
        Click, 1290, 349
        Sleep, 1500

        ; Check the value in column E (index 5)
        if (columns[5] = "U")
        {
            ; Left click at (169, 189) Seleccionar en mapa
            Click, 169, 189
            Sleep, 2000

            ; Left click at (1463, 382) Asignar un solo NSE
            Click, 1463, 382
            Sleep, 2000

            ; Left click at (1266, 590) Casilla un solo NSE
            Click, 1266, 590
            Sleep, 2000

            ; Handle "U" logic for columns F to P (indices 6 to 16)
            coordsSelectU := {6: [1268, 637], 7: [1268, 661], 8: [1268, 685], 9: [1268, 709], 10: [1268, 733], 11: [1268, 757], 12: [1268, 781], 13: [1268, 825], 14: [1268, 856], 15: [1268, 881], 16: [1268, 908]}
            Loop, 11
            {
                colIndex := A_Index + 5
                if (columns[colIndex] > 0)
                {
                    ; Extract X and Y coordinates from the array
                    xPos := coordsSelectU[colIndex][1]
                    yPos := coordsSelectU[colIndex][2]

                    ; Click at the appropriate coordinate
                    Click, %xPos%, %yPos%
                    Sleep, 3000
                }
            }

            ; Left click at (1306, 639) Asignar NSE (Confirm)
            Click, 1306, 639
            Sleep, 2000
        }
        else if (columns[5] = "V")
        {
            ; Left click at (169, 189) Seleccionar en mapa
            Click, 169, 189
            Sleep, 3000

            ; Left click at (1491, 386) Asignar varios NSE
            Click, 1491, 386
            Sleep, 3000

            ; Handle "V" logic for columns F to P (indices 6 to 16)
            coordsSelect := {6: [1235, 563], 7: [1235, 602], 8: [1235, 630], 9: [1235, 668], 10: [1235, 702], 11: [1600, 563], 12: [1600, 602], 13: [1600, 630], 14: [1235, 772], 15: [1235, 804], 16: [1235, 838]}
            coordsType := {6: [1365, 563], 7: [1365, 602], 8: [1365, 630], 9: [1365, 668], 10: [1365, 702], 11: [1730, 563], 12: [1730, 602], 13: [1730, 630], 14: [1365, 772], 15: [1365, 804], 16: [1365, 838]}
            Loop, 11
            {
                colIndex := A_Index + 5
                if (columns[colIndex] > 0)
                {
                    ; Extract X and Y coordinates from the array coordsSelect
                    xPosCS := coordsSelect[colIndex][1]
                    yPosCS := coordsSelect[colIndex][2]
                    ; Extract X and Y coordinates from the array coordsType
                    xPosCT := coordsType[colIndex][1]
                    yPosCT := coordsType[colIndex][2]					
					
                    ; Click and type the value
                    Click, %xPosCS%, %yPosCS%
                    Sleep, 2000
                    Click, %xPosCT%, %yPosCT%
                    Sleep, 2000
					Send, % columns[colIndex]
                    Sleep, 2000
                }
            }

            ; Left click at (1648, 752) - Asignar NSE (Confirm)
            Click, 1648, 752
            Sleep, 2000
            ; Left click at (1598, 823) - Cerrar Asignar NSE 
            Click, 1598, 823
            Sleep, 2000			
        }

	if (columns[17] > 0)
	{
		; Left click at (1563, 385) Boton Administrar Servicios
         Click, 1563, 385
		 Sleep, 2000
		; Left click at (100, 114) Selecciona detalle con NSE
         Click, 100, 114
		 Sleep, 2000 
			if (columns[18] > 0) ; VOZ COBRE TELMEX LINEAS DE COBRE
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 127, 383 ; Cantidad
				Sleep, 2000
				Send, % columns[18] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar
				Sleep, 2000
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
	
			}
			if (columns[19] > 0) ; Datos s/dom
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 2 ; Datos
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}
				Sleep, 2000
				Click, 127, 383 ; Cantidad
				Sleep, 2000
				Send, % columns[19] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar
				Sleep, 2000	
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			}
			if (columns[20] > 0) ; Datos-cobre-telmex-inf
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 2 ; Datos
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}
				Sleep, 2000
				Click, 159, 355 ; Producto
				Sleep, 2000
				Send, {Down}
				Sleep, 2000
				Send, {Enter}
				Sleep, 2000	
				Click, 127, 383 ; Cantidad
				Sleep, 2000
				Send, % columns[20] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar
				Sleep, 2000	
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			
			}
			if (columns[21] > 0) ; Datos-fibra-telmex-inf
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 2 ; Datos
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}
				Sleep, 2000
				Click, 152, 294 ; Tipo
				Sleep, 2000
				Send, {Down}
				Sleep, 2000
				Send, {Enter}
				Sleep, 2000
				Click, 150, 323 ; Empresa
				Sleep, 2000
				Send, {Down}
				Sleep, 2000
				Send, {Enter}
				Sleep, 2000	
				Click, 127, 383 ; Cantidad
				Sleep, 2000
				Send, % columns[21] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar	
				Sleep, 2000	
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
		
			}
			if (columns[22] > 0) ; TV cable otros
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 3 ; TV
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000	
				Click, 150, 323 ; Empresa
				Sleep, 2000		
				Loop, 4 ; otros
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000	
				Click, 127, 383 ; Cantidad
				Sleep, 2000
				Send, % columns[22] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar	
				Sleep, 2000
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			}
			if (columns[23] > 0) ; Dish
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 3 ; TV
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}		
				Sleep, 2000
				Click, 152, 294 ; Tipo
				Sleep, 2000
				Loop, 2 ; Satelital
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000
				Click, 150, 323 ; Empresa
				Sleep, 2000		
				Send, {Down} ; Dish
				Sleep, 2000
				Send, {Enter}
				Sleep, 2000
				Click, 127, 383 ; Cantidad
				Sleep, 2000				
				Send, % columns[23] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar	
				Sleep, 2000			
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			}
			if (columns[24] > 0) ; TVS
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 3 ; TV
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}		
				Sleep, 2000
				Click, 152, 294 ; Tipo
				Sleep, 2000
				Loop, 2 ; Satelital
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Click, 150, 323 ; Empresa
				Sleep, 2000		
				Loop, 2 ; otro
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000	
				Click, 127, 383 ; Cantidad
				Sleep, 2000				
				Send, % columns[24] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar
				Sleep, 2000		
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			}
			if (columns[25] > 0) ; SKY
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 3 ; TV
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000	
				Click, 152, 294 ; Tipo
				Sleep, 2000
				Loop, 2 ; Satelital
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000
				Click, 150, 323 ; Empresa
				Sleep, 2000		
				Loop, 3 ; SKY
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000	
				Click, 127, 383 ; Cantidad
				Sleep, 2000				
				Send, % columns[25] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar	
				Sleep, 2000		
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			}
			if (columns[26] > 0) ; VETV
			{
				; Left click at (100, 114) Selecciona detalle con NSE
				Click, 100, 114
				Sleep, 2000 
				Click, 138, 269 ; Servicio
				Sleep, 2000
				Loop, 3 ; TV
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000	
				Click, 152, 294 ; Tipo
				Sleep, 2000
				Loop, 2 ; Satelital
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000
				Click, 150, 323 ; Empresa
				Sleep, 2000		
				Loop, 5 ; VETV
				{
					Send, {Down}
					Sleep, 2000
				}
				Send, {Enter}	
				Sleep, 2000
				Click, 127, 383 ; Cantidad
				Sleep, 2000				
				Send, % columns[26] ;Cantidad de servicios
				Sleep, 2000
				Click, 82, 423 ; Guardar	
				Sleep, 2000		
				Loop, 5 ; Error
				{
					Click, 704, 384
					Sleep, 2000
				}				
			}
		Click, 882, 49 ; Cerrar Boton Administrar Servicios
		Sleep, 5000
	} 
	    ; Left click again at (89, 263)
        Click, 89, 263
        Sleep, 3000

        ; Press the directional key down
        Send, {Down}
        Sleep, 3000

        Cont++
    }

    ; After loop: Press F5 every 3 minutes until Esc
    ;MsgBox, Script completado. Presionará F5 cada 3 minutos (presione Esc para salir).
	Click, 39, 55 ;Logo GE
    while !GetKeyState("Esc", "P") {
        Send, {F5}
        Sleep, 3000  ; 3 minutes = 180000 180,000 ms
    }
    MsgBox, Script finalizado.
return

; Close GUI if window is exited
GuiClose:
    ExitApp