
#Persistent
#SingleInstance force

; Script de AutoHotkey para escribir texto en coordenadas específicas
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_writer_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_writer_command.txt
        
        ; Parsear comando: x,y,texto
        Array := StrSplit(comando, "|")
        x_campo := Array[1]
        y_campo := Array[2]
        texto := Array[3]
        
        ; Hacer click en las coordenadas especificadas
        Click, %x_campo% %y_campo%
        Sleep, 300
        
        ; Seleccionar y limpiar el campo (opcional)
        Send, ^a
        Sleep, 100
        Send, {Delete}
        Sleep, 100
        
        ; Escribir el texto
        SendInput, %texto%
        Sleep, 300
        
        ; Confirmación para Python
        FileAppend, done, ahk_writer_done.txt
    }
    Sleep, 500  ; Revisar cada medio segundo
}
