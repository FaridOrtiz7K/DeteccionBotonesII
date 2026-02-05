
#Persistent
#SingleInstance force

; Script de AutoHotkey para clics y flechas down
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_click_down_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_click_down_command.txt
        
        ; Parsear comando: x,y,veces_down
        Array := StrSplit(comando, "|")
        x_campo := Array[1]
        y_campo := Array[2]
        veces_down := Array[3]
        
        ; Hacer click en las coordenadas especificadas
        Click, %x_campo% %y_campo%
        Sleep, 300
        
        ; Presionar flecha down las veces especificadas
        Loop, %veces_down% {
            Send, {Down}
            Sleep, 150
        }
        
        ; Confirmación para Python
        FileAppend, done, ahk_click_down_done.txt
    }
    Sleep, 500  ; Revisar cada medio segundo
}
