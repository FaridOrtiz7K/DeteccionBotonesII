#Persistent
#SingleInstance force
#WinActivateForce

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
        
        ; Activar la ventana primero
        WinActivate, A
        
        ; Hacer click en las coordenadas especificadas
        MouseMove, %x_campo%, %y_campo%, 0
        Sleep, 100
        Click, %x_campo% %y_campo%
        Sleep, 800
        
        ; Asegurar que el campo está enfocado
        Sleep, 300
        Click, %x_campo% %y_campo%
        Sleep, 500
        
        ; Seleccionar todo el texto existente
        Send, ^a
        Sleep, 300
        Send, {Delete}
        Sleep, 300
        
        ; Escribir el texto
        SendInput, %texto%
        Sleep, 500
        
        ; Verificar que se escribió el texto
        Send, ^a
        Sleep, 200
        Send, ^c
        Sleep, 200
        
        ; Guardar texto copiado para verificación
        clipboard_text := Clipboard
        FileAppend, Texto escrito: %clipboard_text%`n, ahk_writer_debug.txt
        
        ; Restaurar selección
        Send, {Right}
        Sleep, 100
        
        ; Confirmación para Python
        FileAppend, done, ahk_writer_done.txt
        
        ; Limpiar portapapeles
        Clipboard := ""
    }
    Sleep, 300
}