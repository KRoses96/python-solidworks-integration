"""

Author: Manuel Rosa
Description: Starts solidworks and runs a macro

This is a way of creating a connection between solidworks and python by starting up solidworks and running a macro.
The macro OPEN_FILE.swp will then run all solidworks operations and run all python executables.


"""

import subprocess
import os
import psutil
import tkinter as tk
import sys
from tkinter import ttk

script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
parent_dir = os.path.dirname(script_dir)
options = os.path.join("op.txt")


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
            wait_root.configure(bg="#34495E")  # Set background color

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
