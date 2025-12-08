
#Persistent
#SingleInstance force

; Script de AutoHotkey para copiar y eliminar valores
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_command.txt
        
        ; Parsear comando: x,y
        Array := StrSplit(comando, ",")
        x_campo := Array[1]
        y_campo := Array[2]
        
        ; Ejecutar acciones: copiar y eliminar
        Click, %x_campo% %y_campo%
        Sleep, 200
        
        ; Seleccionar todo el texto
        Send, ^a
        Sleep, 100
        
        ; Copiar el valor seleccionado al portapapeles
        Send, ^c
        Sleep, 200
        
        ; Guardar el valor del portapapeles en un archivo
        ClipboardTemp := Clipboard
        FileAppend, %ClipboardTemp%, ahk_copied_value.txt
        Sleep, 100
        
        ; Confirmación para Python
        FileAppend, done, ahk_done.txt
    }
    Sleep, 500  ; Revisar cada medio segundo
}
