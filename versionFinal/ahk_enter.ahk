#Persistent
#SingleInstance force

; Script de AutoHotkey para presionar Enter múltiples veces
SetBatchLines, -1  ; Máxima velocidad de ejecución

Loop {
    ; Esperar comandos de Python
    IfExist, ahk_command.txt
    {
        FileRead, comando, ahk_command.txt
        if (ErrorLevel = 0) {
            FileDelete, ahk_command.txt
            
            ; Validar que el comando sea un número
            veces := Trim(comando)
            if veces is integer
            {
                if (veces > 0 && veces <= 100)  ; Límite razonable
                {
                    ; Presionar Enter el número de veces especificado
                    Loop, %veces% {
                        Send, {Enter}
                        Sleep, 600  ; Pausa entre cada Enter
                    }
                    
                    ; Confirmación para Python (crear y mantener el archivo)
                    FileDelete, ahk_done.txt  ; Limpiar posible archivo anterior
                    FileAppend, done, ahk_done.txt
                }
                else
                {
                    FileDelete, ahk_done.txt
                    FileAppend, error: número inválido, ahk_done.txt
                }
            }
            else
            {
                FileDelete, ahk_done.txt
                FileAppend, error: no es número, ahk_done.txt
            }
        }
    }
    Sleep, 500  ; Revisar cada medio segundo
}