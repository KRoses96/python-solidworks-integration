# Python - Solidworks Connector

---

### Problem

Integrating SolidWorks with other applications can often be a challenging task. Whenever possible, I prefer to use Python for automation and scripting, as it offers a more streamlined approach. That's why I'm excited to present a solution that effectively integrates the SolidWorks API with more advanced tools, enabling greater flexibility and sophistication in complex projects.

---

### Solution

Begin by creating a Python script that launches SolidWorks and triggers the initial macro. Whether you’re programming with VBA, VB, or C#, use your preferred language to extract the necessary data—such as DXF files, CSV tables, or rendered images. Once the data extraction is complete, your SolidWorks macro should then initiate a Python executable or script to continue the operation.

This example demonstrates how to establish the initial connection. From here, you can build out a more extensive, multi-stack solution tailored to your specific needs.

All of the following code snippets are in this repo as separate files if you prefer to simply clone it, note that .swp is not recognized by github.

---

If you'd like to see an example of a full project using this method, check out the following repository:

Please note that this solution was originally developed for a very specific case and wasn’t initially intended to scale to the magnitude it eventually reached. My best advice is to consider scalability from the outset when developing your solution. Nonetheless, the repository contains a wealth of useful SolidWorks macros and offers valuable insights into tackling common engineering challenges in industrial manufacturing.

---

### Examples

This is the primary method for starting a SolidWorks macro using Python. It’s a straightforward process, but it’s important to ensure that you don’t accidentally start a new instance of SolidWorks. To prevent conflicts, this script will close any running instances of SolidWorks before executing the macro. I've also added a warning message to alert the user if SolidWorks is already running and there’s unsaved work. You can remove this warning if you're fully automating the process and want to eliminate any user input from your application.

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

        no_button.grid(row=1, column=1, padx=5, pady=10)

        root.mainloop()

    open_solidworks_macro()

    SystemExit
```

##### Solidworks Macro to start a python executable:

This macro is designed to open a SolidWorks file and initiate a specific macro based on the contents of a text file. The macro reads cases from the text file and maps them to corresponding macros. I've frequently used this approach to integrate with my Python GUI, allowing the GUI to dynamically update these values in the handle functions. This setup enables the execution of different automations with each button, making the workflow more flexible and efficient.

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

##### Example Macro

Note that this macro's primary function is to launch an executable—it doesn't perform any SolidWorks operations on its own. This is where you would implement the SolidWorks-specific logic before handing control back to Python. It's crucial to ensure that any files generated or modified by this macro are accessible to your Python script, allowing Python to seamlessly continue the workflow.

Again if you want to check on some macros for this methodology you can check the following repo:


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


