#Persistent
#SingleInstance force

Loop {
    IfExist, ahk_ctrl_s_command.txt
    {
        FileDelete, ahk_ctrl_s_command.txt
        Send, ^s
        Sleep, 300
        FileAppend, 1, ahk_ctrl_s_done.txt
    }
    Sleep, 500
}