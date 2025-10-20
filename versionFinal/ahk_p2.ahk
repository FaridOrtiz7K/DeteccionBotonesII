#Persistent
#SingleInstance force

; Variables globales
global StartCount := 1
global LoopCount := 589
global CSVFile := "C:\Users\cmf05\Documents\AutoHotkey\NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
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
        
        if (comando_principal = "CLICK") {
            x := Array[2]
            y := Array[3]
            sleep_time := Array[4]
            Click, %x% %y%
            Sleep, %sleep_time%
        }
        else if (comando_principal = "SEND") {
            texto := Array[2]
            sleep_time := Array[3]
            Send, %texto%
            Sleep, %sleep_time%
        }
        else if (comando_principal = "PRESS") {
            tecla := Array[2]
            sleep_time := Array[3]
            Send, {%tecla%}
            Sleep, %sleep_time%
        }
        else if (comando_principal = "SET_CONFIG") {
            StartCount := Array[2]
            LoopCount := Array[3]
            if (Array[4] != "") {
                CSVFile := Array[4]
            }
        }
        else if (comando_principal = "READ_CSV") {
            FileRead, csvContent, %CSVFile%
            if (csvContent = "") {
                FileAppend, ERROR|No se pudo leer el CSV, ahk_response.txt
            } else {
                FileAppend, SUCCESS, ahk_response.txt
            }
        }
        else if (comando_principal = "EXECUTE_NSE") {
            IsRunning := true
            ExecuteNSEScript()
            IsRunning := false
        }
        else if (comando_principal = "STOP_SCRIPT") {
            IsRunning := false
            FileAppend, STOPPED, ahk_response.txt
            ;ExitApp  ; Opcional: descomentar si quieres que se cierre completamente
        }
        else if (comando_principal = "GET_STATUS") {
            if (IsRunning) {
                FileAppend, RUNNING|%CurrentLine%, ahk_response.txt
            } else {
                FileAppend, IDLE, ahk_response.txt
            }
        }
        else if (comando_principal = "U_LOGIC") {
            ; Parámetros: col6|col7|...|col16
            HandleULogic(Array)
        }
        else if (comando_principal = "V_LOGIC") {
            ; Parámetros: col6|col7|...|col16
            HandleVLogic(Array)
        }
        else if (comando_principal = "SERVICES") {
            ; Parámetros: col18|col19|...|col26
            HandleServices(Array)
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
        return
    }
    
    lines := StrSplit(csvContent, "`n")
    if (lines.MaxIndex() < 1) {
        FileAppend, ERROR|CSV con formato incorrecto, ahk_response.txt
        return
    }
    
    Sleep, 5000  ; Espera inicial
    
    Cont := StartCount
    
    Loop, %LoopCount% {
        if (!IsRunning) {
            break
        }
        
        CurrentLine := Cont
        
        ; Parsear línea actual
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
            
            ; Preparar parámetros para U_LOGIC
            u_params := []
            Loop, 11 {
                colIndex := A_Index + 5
                if (columns[colIndex] > 0) {
                    u_params.Push(columns[colIndex])
                } else {
                    u_params.Push(0)
                }
            }
            HandleULogic(u_params)
        }
        else if (columns[5] = "V") {
            Click, 169, 189
            Sleep, 3000
            Click, 1491, 386
            Sleep, 3000
            
            ; Preparar parámetros para V_LOGIC
            v_params := []
            Loop, 11 {
                colIndex := A_Index + 5
                if (columns[colIndex] > 0) {
                    v_params.Push(columns[colIndex])
                } else {
                    v_params.Push(0)
                }
            }
            HandleVLogic(v_params)
        }
        
        ; Manejar servicios
        if (columns[18] > 0) {
            ; Preparar parámetros para SERVICES
            service_params := []
            Loop, 9 {
                colIndex := A_Index + 18
                if (columns[colIndex] > 0) {
                    service_params.Push(columns[colIndex])
                } else {
                    service_params.Push(0)
                }
            }
            HandleServices(service_params)
        }
        
        ; Finalizar iteración
        Click, 89, 263
        Sleep, 3000
        Send, {Down}
        Sleep, 3000
        
        Cont++
        
        ; Actualizar estado
        FileAppend, PROGRESS|%Cont%, ahk_progress.txt
    }
    
    ; Proceso posterior al bucle
    Click, 39, 55
    
    FileAppend, COMPLETED, ahk_response.txt
    
    ; Presionar F5 cada 3 minutos hasta que se detenga
    while (IsRunning) {
        Send, {F5}
        Sleep, 180000  ; 3 minutos
    }
}

; Función para manejar lógica U
HandleULogic(params) {
    Loop, 11 {
        colIndex := A_Index + 5
        if (params[colIndex] > 0) {
            x := CoordsSelectU[colIndex][1]
            y := CoordsSelectU[colIndex][2]
            Click, %x% %y%
            Sleep, 3000
        }
    }
    Click, 1306, 639
    Sleep, 2000
}

; Función para manejar lógica V
HandleVLogic(params) {
    Loop, 11 {
        colIndex := A_Index + 5
        if (params[colIndex] > 0) {
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
            Send, % params[colIndex]
            Sleep, 2000
        }
    }
    Click, 1648, 752
    Sleep, 2000
    Click, 1598, 823
    Sleep, 2000
}

; Función para manejar servicios
HandleServices(params) {
    Click, 1563, 385
    Sleep, 2000
    Click, 100, 114
    Sleep, 2000
    
    ; VOZ COBRE TELMEX
    if (params[1] > 0) {
        HandleVozCobre(params[1])
    }
    
    ; Datos s/dom
    if (params[2] > 0) {
        HandleDatosSDom(params[2])
    }
    
    ; Datos-cobre-telmex-inf
    if (params[3] > 0) {
        HandleDatosCobreTelmex(params[3])
    }
    
    ; ... (implementar los demás servicios de manera similar)
    
    Click, 882, 49
    Sleep, 5000
}

; Funciones individuales de servicios (implementar según necesidad)
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
    ; Implementar según la lógica específica
    Click, 100, 114
    Sleep, 2000
    ; ... más pasos
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