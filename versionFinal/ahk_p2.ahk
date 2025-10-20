#Persistent
#SingleInstance force

; Variables globales
global StartCount := 1
global LoopCount := 589
global CSVFile := "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
global CurrentLine := 0
global IsRunning := false

; Coordenadas para tipo U
global CoordsSelectU := {6: [1268, 637], 7: [1268, 661], 8: [1268, 685], 9: [1268, 709]
                    , 10: [1268, 733], 11: [1268, 757], 12: [1268, 781], 13: [1268, 825]
                    , 14: [1268, 856], 15: [1268, 881], 16: [1268, 908]}

; Coordenadas para tipo V
global CoordsSelect := {6: [1235, 563], 7: [1235, 602], 8: [1235, 630], 9: [1235, 668]
                   , 10: [1235, 702], 11: [1600, 563], 12: [1600, 602], 13: [1600, 630]
                   , 14: [1235, 772], 15: [1235, 804], 16: [1235, 838]}

global CoordsType := {6: [1365, 563], 7: [1365, 602], 8: [1365, 630], 9: [1365, 668]
                 , 10: [1365, 702], 11: [1730, 563], 12: [1730, 602], 13: [1730, 630]
                 , 14: [1365, 772], 15: [1365, 804], 16: [1365, 838]}

; Función principal de loop
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_command.txt
        
        ; Parsear comando: comando|param1|param2|...
        Array := StrSplit(comando, "|")
        comando_principal := Array[1]
        
        if (comando_principal = "SET_CONFIG") {
            StartCount := Array[2]
            LoopCount := Array[3]
            if (Array[4] != "") {
                CSVFile := Array[4]
            }
            FileAppend, CONFIG_OK, ahk_response.txt
        }
        else if (comando_principal = "READ_CSV") {
            FileRead, csvContent, %CSVFile%
            if (csvContent = "") {
                FileAppend, ERROR|No se pudo leer el CSV, ahk_response.txt
            } else {
                FileAppend, CSV_OK, ahk_response.txt
            }
        }
        else if (comando_principal = "EXECUTE_NSE") {
            IsRunning := true
            ExecuteNSEScript()
        }
        else if (comando_principal = "STOP_SCRIPT") {
            IsRunning := false
            FileAppend, STOPPED, ahk_response.txt
        }
        else if (comando_principal = "GET_STATUS") {
            if (IsRunning) {
                FileAppend, RUNNING, ahk_response.txt
            } else {
                FileAppend, IDLE, ahk_response.txt
            }
        }
        
        ; Confirmación para Python
        FileAppend, DONE, ahk_done.txt
    }
    Sleep, 100  ; Revisar más frecuentemente
}

; Función principal de ejecución NSE
ExecuteNSEScript() {
    ; Leer CSV
    FileRead, csvContent, %CSVFile%
    if (csvContent = "") {
        FileAppend, ERROR|CSV vacío, ahk_response.txt
        IsRunning := false
        return
    }
    
    lines := StrSplit(csvContent, "`n")
    if (lines.MaxIndex() < 1) {
        FileAppend, ERROR|CSV con formato incorrecto, ahk_response.txt
        IsRunning := false
        return
    }
    
    FileAppend, STARTED, ahk_response.txt
    
    Sleep, 3000  ; Espera inicial de 3 segundos
    
    Cont := StartCount
    
    Loop, %LoopCount% {
        if (!IsRunning) {
            FileAppend, STOPPED, ahk_response.txt
            break
        }
        
        CurrentLine := Cont
        
        ; Parsear línea actual
        if (Cont + 1 > lines.MaxIndex()) {
            break
        }
        
        currentLine := lines[Cont + 1]
        columns := StrSplit(currentLine, ",")
        
        if (columns.MaxIndex() < 5) {
            Cont++
            continue
        }
        
        ; Proceso principal
        Click, 89, 263
        Sleep, 1500
        Click, 1483, 519
        Sleep, 1500
        Send, {Delete}
        Sleep, 1000
        Send, % columns[2]
        Sleep, 1500
        
        ; Manejar columna D
        if (columns[4] > 0) {
            Click, 1507, 650
            Sleep, 2500
            Loop, % columns[4] {
                Send, {Down}
                Sleep, 2000
            }
        }
        
        Click, 1290, 349
        Sleep, 1500
        
        ; Manejar tipo U o V
        if (columns[5] = "U") {
            Click, 169, 189
            Sleep, 2000
            Click, 1463, 382
            Sleep, 2000
            Click, 1266, 590
            Sleep, 2000
            
            ; Lógica U para columnas 6-16
            Loop, 11 {
                colIndex := A_Index + 5
                if (columns[colIndex] > 0) {
                    x := CoordsSelectU[colIndex][1]
                    y := CoordsSelectU[colIndex][2]
                    Click, %x% %y%
                    Sleep, 3000
                }
            }
            
            Click, 1306, 639
            Sleep, 2000
        }
        else if (columns[5] = "V") {
            Click, 169, 189
            Sleep, 3000
            Click, 1491, 386
            Sleep, 3000
            
            ; Lógica V para columnas 6-16
            Loop, 11 {
                colIndex := A_Index + 5
                if (columns[colIndex] > 0) {
                    ; Click en coordenada select
                    x_cs := CoordsSelect[colIndex][1]
                    y_cs := CoordsSelect[colIndex][2]
                    Click, %x_cs% %y_cs%
                    Sleep, 2000
                    
                    ; Click en coordenada type
                    x_ct := CoordsType[colIndex][1]
                    y_ct := CoordsType[colIndex][2]
                    Click, %x_ct% %y_ct%
                    Sleep, 2000
                    
                    ; Escribir valor
                    Send, % columns[colIndex]
                    Sleep, 2000
                }
            }
            
            Click, 1648, 752
            Sleep, 2000
            Click, 1598, 823
            Sleep, 2000
        }
        
        ; Manejar servicios (columnas 18-26)
        if (columns[18] > 0) {
            Click, 1563, 385
            Sleep, 2000
            Click, 100, 114
            Sleep, 2000
            
            ; VOZ COBRE TELMEX
            if (columns[19] > 0) {
                HandleVozCobre(columns[19])
            }
            
            ; Datos s/dom
            if (columns[20] > 0) {
                HandleDatosSDom(columns[20])
            }
            
            ; Datos-cobre-telmex-inf
            if (columns[21] > 0) {
                HandleDatosCobreTelmex(columns[21])
            }
            
            ; Datos-fibra-telmex-inf
            if (columns[22] > 0) {
                HandleDatosFibraTelmex(columns[22])
            }
            
            ; TV cable otros
            if (columns[23] > 0) {
                HandleTVCableOtros(columns[23])
            }
            
            ; Dish
            if (columns[24] > 0) {
                HandleDish(columns[24])
            }
            
            ; TVS
            if (columns[25] > 0) {
                HandleTVS(columns[25])
            }
            
            ; SKY
            if (columns[26] > 0) {
                HandleSKY(columns[26])
            }
            
            ; VETV
            if (columns[27] > 0) {
                HandleVETV(columns[27])
            }
            
            Click, 882, 49
            Sleep, 5000
        }
        
        ; Finalizar iteración
        Click, 89, 263
        Sleep, 3000
        Send, {Down}
        Sleep, 3000
        
        Cont++
        
        ; Actualizar progreso
        FileAppend, PROGRESS|%Cont%, ahk_progress.txt
    }
    
    ; Proceso posterior al bucle
    Click, 39, 55
    
    FileAppend, COMPLETED, ahk_response.txt
    IsRunning := false
}

; Funciones individuales de servicios
HandleVozCobre(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleDatosSDom(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(2)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleDatosCobreTelmex(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(2)
    Click, 159, 355
    Sleep, 2000
    NavigateDown(1)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleDatosFibraTelmex(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(2)
    Click, 152, 294
    Sleep, 2000
    NavigateDown(1)
    Click, 150, 323
    Sleep, 2000
    NavigateDown(1)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleTVCableOtros(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(3)
    Click, 150, 323
    Sleep, 2000
    NavigateDown(4)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleDish(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(3)
    Click, 152, 294
    Sleep, 2000
    NavigateDown(2)
    Click, 150, 323
    Sleep, 2000
    NavigateDown(1)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleTVS(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(3)
    Click, 152, 294
    Sleep, 2000
    NavigateDown(2)
    Click, 150, 323
    Sleep, 2000
    NavigateDown(2)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleSKY(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(3)
    Click, 152, 294
    Sleep, 2000
    NavigateDown(2)
    Click, 150, 323
    Sleep, 2000
    NavigateDown(3)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

HandleVETV(cantidad) {
    Click, 100, 114
    Sleep, 2000
    Click, 138, 269
    Sleep, 2000
    NavigateDown(3)
    Click, 152, 294
    Sleep, 2000
    NavigateDown(2)
    Click, 150, 323
    Sleep, 2000
    NavigateDown(5)
    Click, 127, 383
    Sleep, 2000
    Send, %cantidad%
    Sleep, 2000
    Click, 82, 423
    Sleep, 2000
    HandleErrorClick()
}

; Funciones auxiliares
NavigateDown(count) {
    Loop, %count% {
        Send, {Down}
        Sleep, 2000
    }
}

HandleErrorClick() {
    Loop, 5 {
        Click, 704, 384
        Sleep, 2000
    }
}

; Hotkey para detener manualmente
Esc::
    IsRunning := false
    FileAppend, MANUAL_STOP, ahk_response.txt
return