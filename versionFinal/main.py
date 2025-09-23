from controller import ScriptController


def main():
    # Asegurar que pyautogui falle rápido si hay error
    import pyautogui
    pyautogui.FAILSAFE = True
    
    app = ScriptController()
    app.run()

if __name__ == "__main__":
    main()