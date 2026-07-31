# Python - SolidWorks Connector

Minimal pattern for bridging SolidWorks with Python: a Python script launches SolidWorks and runs a macro; the macro (VBA/VB/C#) does the SolidWorks-side work (export DXF/CSV/images, etc.) and then shells out to a Python executable to continue the pipeline.

All snippets below exist as separate files in this repo (`macrorun.py`, `OPEN_FILE.swp`, `OPEN_PYTHON_EXECUTABLE.swp`). Note GitHub doesn't render `.swp` files.

A full example project using this method: [python-solidworks-integration-example](https://github.com/KRoses96/python-solidworks-integration-example). It was built for a specific use case that outgrew its original scope, so plan for scalability up front if you can. It has useful macros for common manufacturing automation tasks.

---

### 1. `macrorun.py`: launch SolidWorks and run a macro

Closes any running SolidWorks instance before launching (with a confirmation prompt to avoid losing unsaved work) and then runs the target macro.

```python
import subprocess
import os
import psutil
import tkinter as tk
import sys
from tkinter import ttk

script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
parent_dir = os.path.dirname(script_dir)


solidworksPath = "/YourSolidworksPath"


def open_solidworks_macro():

    macro_path = os.path.join(parent_dir, "OPEN_FILE.swp")
    solidworks_executable = (
        solidworksPath  
    )
    print(solidworks_executable)
    try:
        if not os.path.exists(solidworks_executable):
            raise Exception("SolidWorks executable not found.")

        cmd = [solidworks_executable, "/m", macro_path]
        subprocess.Popen(cmd, shell=True)
        print(f'SolidWorks macro "{macro_path}" has been opened.')
        SystemExit

    except Exception as e:
        error_message = f"An error occurred: {str(e)}"
        print(error_message)
        SystemExit


def check_sldworks_running():
    for proc in psutil.process_iter():
        try:
            if proc.name() == "SLDWORKS.exe":
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


if __name__ == "__main__":
    sld_bool = check_sldworks_running()
    if sld_bool:
        root = tk.Tk()
        root.title("Be careful!")

        window_width = 230
        window_height = 100
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)

        label = tk.Label(
            root,
            text="O Solidworks will close,\n Save files before closing!",
        )
        label.grid(row=0, column=0, columnspan=2, pady=10)

        def exit_program():
            root.destroy()
            subprocess.call(["taskkill", "/f", "/im", "SLDWORKS.exe"])
            open_solidworks_macro()
            wait_root = tk.Tk()
            wait_root.overrideredirect(True) 
            wait_root.geometry(
                "+{}+{}".format(
                    (wait_root.winfo_screenwidth() - 250) // 2,
                    (wait_root.winfo_screenheight() - 100) // 2,
                )
            )
            wait_root.configure(bg="#34495E")

            border_frame = ttk.Frame(wait_root, style="White.TFrame")
            border_frame.pack(padx=10, pady=10, fill="both", expand=True)

            wait_label = tk.Label(
                border_frame,
                text="Wait,\nSolidworks initializing",
                fg="white",
                bg="#34495E",
            )
            wait_label.pack(pady=10)

            def close_wait_window():
                wait_root.destroy()

            wait_root.after(5000, close_wait_window)
            wait_root.mainloop()
            raise SystemExit

        yes_button = tk.Button(root, text="Ok", command=exit_program, width=10)
        yes_button.grid(row=1, column=0, padx=5, pady=10)

        def close_window():
            root.destroy()
            raise SystemExit

        root.protocol("WM_DELETE_WINDOW", close_window)

        no_button = tk.Button(root, text="Cancel", command=close_window, width=10)
        no_button.grid(row=1, column=1, padx=5, pady=10)

        root.mainloop()

    open_solidworks_macro()

    SystemExit
```

### 2. SolidWorks macro: reads `op.txt` and dispatches

Opens a file, reads an operation code from a text file, and runs the matching macro/executable. Useful for driving different automations from buttons in a Python GUI that just rewrites `op.txt`.

```vba
Option Explicit
Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2
Dim doc As SldWorks.ModelDoc2
Dim fileerror As Long
Dim filewarning As Long
Dim txt_character
Dim CPath As String
Dim directoryPath_2 As String
Dim directoryPath_3 As String
Dim MacroPath As String
Dim exePath As String
Dim boolstatus As Boolean
Dim lErrors As Long
Dim lWarnings As Long

Sub main()
    Dim Filter As String
    Dim fileName As String
    Dim fileConfig As String
    Dim fileDispName As String
    Dim fileOptions As Long

    Set swApp = Application.SldWorks

    Filter = "SOLIDWORKS Files (*.sldasm)|*.sldasm|All Files (*.*)|*.*|"
    fileName = swApp.GetOpenFileName("File to Attach", "", Filter, fileOptions, fileConfig, fileDispName)
    Set doc = swApp.OpenDoc6(fileName, swDocPART, swOpenDocOptions_Silent, "", fileerror, filewarning)
    Set doc = swApp.OpenDoc6(fileName, swDocASSEMBLY, swOpenDocOptions_Silent, "", fileerror, filewarning)
    
    Dim FilePath As String
    Dim FileContent As String
    Dim Op As String
    

    CPath = swApp.GetCurrentMacroPathName()
    Dim directoryPath As String
    directoryPath = Left(CPath, InStrRev(CPath, "\"))
    
  
    FilePath = Left(directoryPath, Len(directoryPath) - 7) & "op.txt"
    
    Open FilePath For Input As #1
    FileContent = Input$(LOF(1), 1)
    Close #1
    

    Op = Mid(FileContent, 9, 1)
    

    Select Case Op
        Case "0"
            directoryPath_2 = 
            RunMacro directoryPath_2, "MacroName", "main"
            

            Dim swModel As SldWorks.ModelDoc2
            Set swModel = swApp.GetFirstDocument
            boolstatus = swModel.Save3(swSaveAsOptions_Silent, lErrors, lWarnings)
            
 
            MacroPath = Left(directoryPath, Len(directoryPath) - 7) & "exe"
            
 
            exePath = MacroPath & "\Imprimir.exe"

            Dim objShell As Object
            Dim command As String
            Set objShell = CreateObject("WScript.Shell")

            command = "cmd.exe /c """ & exePath & """"
            Shell exePath, vbHide

    End Select
End Sub

Sub RunMacro(path As String, moduleName As String, procName As String)
    swApp.RunMacro2 path, moduleName, procName, swRunMacroOption_e.swRunMacroUnloadAfterRun, 0
End Sub
```

### 3. Example macro: hand off to an executable

Doesn't do any SolidWorks-specific work itself. This is the placeholder for your own SolidWorks logic before launching the next executable in the chain. Make sure any files it produces are accessible to the Python side.

```vba
Dim swApp As SldWorks.SldWorks
Dim CPath As String
Dim directoryPath_2 As String
Dim directoryPath_3 As String
Dim directoryPath_4 As String

Sub main()
    
    Set swApp = Application.SldWorks
    CPath = swApp.GetCurrentMacroPathName()
    Dim directoryPath As String
    directoryPath = Left(CPath, InStrRev(CPath, "\"))
    
    exePath = 'executable path

    Dim objShell As Object
    Dim command As String
    Set objShell = CreateObject("WScript.Shell")
    
    command = "cmd.exe /c """ & exePath & """"
    
    Shell exePath, vbHide
    
End Sub 
```
