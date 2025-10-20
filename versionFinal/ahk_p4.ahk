#Persistent
#SingleInstance force

; Variables globales
global NumRa := 1
global NumTxtType := 0
global IsRunning := false
global CSVFile := "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"

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
            if (Array[2] != "") {
                CSVFile := Array[2]
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
        else if (comando_principal = "EXECUTE_GE_ROUTE") {
            IsRunning := true
            ExecuteGERoute()
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

; Función para ejecutar la automatización GE Route
ExecuteGERoute() {
    ; Check if the CSV file exists
    if (!FileExist(CSVFile)) {
        FileAppend, ERROR|CSV no existe, ahk_response.txt
        IsRunning := false
        return
    }

    ; Read the CSV file
    FileRead, csvContent, %CSVFile%
    lines := StrSplit(csvContent, "`n")
    
    ; Check if we have enough lines
    if (lines.MaxIndex() < 1) {
        FileAppend, ERROR|CSV vacío, ahk_response.txt
        IsRunning := false
        return
    }

    FileAppend, STARTED, ahk_response.txt
    
    Sleep, 3000  ; Espera inicial de 3 segundos

    NumRa := 1
    NumTxtType := 0

    Loop, 9 {
        if (!IsRunning) {
            FileAppend, STOPPED, ahk_response.txt
            break
        }

        ; Select Agregar ruta de GE
        Click, 327, 381
        Sleep, 2000

        ; Archivo
        Click, 1396, 608
        Sleep, 3000

        ; Abrir
        Click, 1406, 634
        Sleep, 3000

        ; Documents
        Click, 1120, 666
        Sleep, 3000

        ; Press Alt + n
        Send, !n
        Sleep, 3000

        ; Check if we are at a valid iteration
        if (NumRa <= lines.MaxIndex()) {
            ; Split the current line into columns
            values1 := StrSplit(lines[NumRa], ",")

            ; Ensure the second column exists
            if (values1.MaxIndex() >= 2) {
                NumTxtType := values1[1]
                Send, RA %NumTxtType%.kml
                Sleep, 1000
            }
        }

        Sleep, 3000
        Send, {Enter}
        Sleep, 3000

        ; Select Agregar ruta de GE
        Click, 327, 381
        Sleep, 2000

        ; Cargar ruta
        Click, 1406, 675
        Sleep, 2000

        ; Select Lote
        Click, 70, 266
        Sleep, 2000

        ; Seleccionar en el mapa
        Click, 168, 188
        Sleep, 2000

        ; Anotar
        Click, 1366, 384
        Sleep, 2000

        ; Agregar texto adicional
        Click, 1449, 452
        Sleep, 2000

        ; Escriba el texto adicional
        Click, 1421, 526
        Sleep, 2000

        ; Press Backspace
        Send, {Backspace}
        Sleep, 1000

        ; Check if we are at a valid iteration
        if (NumRa <= lines.MaxIndex()) {
            ; Split the current line into columns
            values := StrSplit(lines[NumRa], ",")

            ; Ensure the second column exists
            if (values.MaxIndex() >= 2) {
                ; Write the value from the second column
                Send, % values[2]
                Sleep, 1000

                ; Agregar texto
                Click, 1263, 572
                Sleep, 3000

                ; Cerrar
                Click, 1338, 570
                Sleep, 2000
            }
        }

        ; Every 10 steps, press Ctrl + S
        if (Mod(A_Index, 10) = 0) {
            Send, ^s
            Sleep, 6000
        }

        ; Limpiar trazo
        Click, 360, 980
        Sleep, 1000

        ; Select Lote
        Click, 70, 266
        Sleep, 2000

        ; Press Down
        Send, {Down}
        Sleep, 2000

        NumRa++

        ; Actualizar progreso
        FileAppend, PROGRESS|%NumRa%, ahk_progress.txt
    }

    ; Press Ctrl + S at the end
    Send, ^s
    FileAppend, COMPLETED, ahk_response.txt
    IsRunning := false
}

; Hotkey para detener manualmente
Esc::
    IsRunning := false
    FileAppend, MANUAL_STOP, ahk_response.txt
return