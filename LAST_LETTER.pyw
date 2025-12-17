import threading
import time
import random
import subprocess
import ctypes
import webbrowser

import keyboard
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
from packaging import version
import requests

from english_words import get_english_words_set

import sys, os

CURRENT_VERSION="1.0"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


class LastLetterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Made by Blassian3")
        self.root.resizable(False, False)
        self.root.attributes('-topmost', True)


        self.root.overrideredirect(True)
        self.root.configure(bg="black")


        title_bar = tk.Frame(self.root, bg="black", height=32)
        title_bar.pack(fill="x")

        title_label = tk.Label(
            title_bar,
            text="Made by Blassian3",
            bg="black",
            fg="#8C92AC",
            font=("Segoe UI", 10)
        )
        title_label.pack(side="left", padx=10)

        close_btn = tk.Button(
            title_bar,
            text="X",
            bg="black",
            fg="white",
            borderwidth=0,
            command=self.root.destroy
        )
        close_btn.pack(side="right", padx=10)



        def start_move(event):
            self.x = event.x
            self.y = event.y

        def stop_move(event):
            self.x = None
            self.y = None

        def do_move(event):
            self.root.geometry(f"+{event.x_root - self.x}+{event.y_root - self.y}")

        title_bar.bind("<Button-1>", start_move)
        title_bar.bind("<ButtonRelease-1>", stop_move)
        title_bar.bind("<B1-Motion>", do_move)


        self.canvas = tk.Canvas(self.root, width=420, height=380, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.stars = []
        for _ in range(120):
            x = random.randint(0, 420)
            y = random.randint(0, 380)
            size = random.randint(1, 2)
            speed = random.uniform(0.2, 0.7)
            star_id = self.canvas.create_oval(x, y, x+size, y+size, fill="#E0FFFF")
            self.stars.append((star_id, speed))
        

        self.animate_stars()  


        self.ui = tk.Frame(self.canvas, bg="black")
        self.canvas.create_window(0, 0, anchor="nw", window=self.ui)

        try:
            self.root.iconbitmap(resource_path("LastLetter.ico"))
        except:
            pass

        self.wordlist = []
        self.wordlist_loaded = False
        self.wordlist_error = None
        self.used_words = set()

        self.prefix_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="Random Words")

        label = tk.Label(self.ui, text="Starting letters:", fg="white", bg="black", font=("Segoe UI", 11))
        label.grid(row=0, column=0, sticky="w")

        self.entry = tk.Entry(self.ui, textvariable=self.prefix_var, width=25, bg="#111", fg="white", insertbackground="white", relief="flat")
        self.entry.grid(row=1, column=0, columnspan=2, sticky="we", pady=(2, 8))
        self.entry.focus_set()
        self.entry.bind("<Control-Return>", self.on_ctrl_enter)

        self.status_var = tk.StringVar(value="Loading word list...")
        status_label = tk.Label(self.ui, textvariable=self.status_var, fg="gray", bg="black")
        status_label.grid(row=2, column=0, sticky="w")

        self.roblox_status_var = tk.StringVar(value="Roblox not Found")
        self.roblox_status_label = tk.Label(self.ui, textvariable=self.roblox_status_var, fg="red", bg="black")
        self.roblox_status_label.grid(row=2, column=1, sticky="e")

        btn_style = dict(bg="#222", fg="white", activebackground="#333", relief="flat")

        tk.Button(self.ui, text="Play round", command=self.on_play_round, **btn_style).grid(row=3, column=0, sticky="we", pady=(4, 4))
        tk.Button(self.ui, text="View Used Words", command=self.show_used_words, **btn_style).grid(row=3, column=1, sticky="we", pady=(4, 4), padx=(4, 0))

        speed_label = tk.Label(self.ui, text="Typing speed (ms per char):", fg="white", bg="black")
        speed_label.grid(row=4, column=0, columnspan=2, sticky="w")

        self.speed_var = tk.DoubleVar(value=30.0)
        self.speed_scale = tk.Scale(
            self.ui,
            from_=10,
            to=1000,
            orient="horizontal",
            variable=self.speed_var,
            showvalue=True,
            length=250,
            resolution=5,
            bg="black",
            fg="white",
            highlightbackground="black",
            troughcolor="#222"
        )
        self.speed_scale.grid(row=5, column=0, columnspan=2, sticky="we", pady=(0, 4))

        ttk.Style().configure("Dark.TCombobox",
            fieldbackground="#111",
            background="#111",
            foreground="white"
        )

        mode_frame = tk.Frame(self.ui, bg="black")
        mode_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=(4, 0))

        tk.Label(mode_frame, text="Word selection mode:", fg="white", bg="black").pack(side="left")
        self.mode_combobox = ttk.Combobox(
            mode_frame,
            textvariable=self.mode_var,
            values=["Random Words", "Short Words", "Long Words"],
            state="readonly",
            width=15,
            style="Dark.TCombobox"
        )
        self.mode_combobox.pack(side="left", padx=10)

    
        self.typing_mode_var = tk.StringVar(value="Normal")
        typing_mode_frame = tk.Frame(self.ui, bg="black")
        typing_mode_frame.grid(row=7, column=0, columnspan=2, sticky="we", pady=(4, 0))

        tk.Label(typing_mode_frame, text="Typing mode:", fg="white", bg="black").pack(side="left")
        self.typing_mode_combobox = ttk.Combobox(
            typing_mode_frame,
            textvariable=self.typing_mode_var,
            values=["Normal", "Fixed Speed", "Jitter", "Cycle", "Burst", "Random"],
            state="readonly",
            width=15,
            style="Dark.TCombobox"
        )
        self.typing_mode_combobox.pack(side="left", padx=10)

        tk.Button(self.ui, text="Clear Used Words", command=self.on_clear_cache, **btn_style).grid(
            row=8, column=0, sticky="we", pady=4
        )
        tk.Button(self.ui, text="Quit", command=self.root.destroy, **btn_style).grid(
            row=8, column=1, sticky="we", pady=4, padx=(4, 0)
        )

        credit_label = tk.Label(self.ui, text="Made by Blassian3 on discord", fg="#8C92AC", bg="black")
        credit_label.grid(row=9, column=0, columnspan=2, sticky="e", pady=(2,0))

        self.root.bind("<FocusIn>", lambda event: self.entry.focus_set())

        threading.Thread(target=self.load_wordlist, daemon=True).start()
        self.refresh_roblox_status()

        self.live_buffer = ""
        keyboard.hook(self._on_key_event)

    def load_wordlist(self) -> None:
        try:
            words_set = get_english_words_set(["web2"], lower=False)
            self.wordlist = sorted(words_set)
            self.wordlist_loaded = True
            self.wordlist_error = None
            self.root.after(0, lambda: self.status_var.set("Word list loaded."))
        except Exception as exc:
            self.wordlist_loaded = False
            self.wordlist_error = exc
            self.root.after(0, lambda: self.status_var.set("Failed to load word list."))

    def animate_stars(self):
        for star_id, speed in self.stars:
            self.canvas.move(star_id, 0, speed)
            x1, y1, x2, y2 = self.canvas.coords(star_id)
            if y1 > 380:
                new_x = random.randint(0, 420)
                new_y = 0
                self.canvas.coords(star_id, new_x, new_y, new_x+2, new_y+2)

        self.root.after(30, self.animate_stars)

    def find_completion(self, prefix: str) -> str | None:
        if not self.wordlist_loaded or not self.wordlist:
            return None

        lower_prefix = prefix.lower()
        candidates = []

        for word in self.wordlist:
            if word in self.used_words:
                continue
            if not word.lower().startswith(lower_prefix):
                continue
            if len(word) <= len(prefix):
                continue
            candidates.append(word)

        if not candidates:
            return None

        mode = self.mode_var.get()
        if mode == "Short Words":
            chosen = min(candidates, key=len)
        elif mode == "Long Words":
            chosen = max(candidates, key=len)
        else:
            chosen = random.choice(candidates)

        self.used_words.add(chosen)
        return chosen[len(prefix):]

    def _is_roblox_running(self) -> bool:
        try:
            startupinfo = None
            creationflags = 0
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW

            output = subprocess.check_output(
                ["tasklist"],
                text=True,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
        except Exception:
            return False

        return "RobloxPlayerBeta.exe" in output

    def _focus_roblox_window(self) -> None:
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            hwnd = user32.FindWindowW(None, "Roblox")
            if hwnd:
                if user32.IsIconic(hwnd):
                    user32.ShowWindow(hwnd, 9)
                user32.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def refresh_roblox_status(self) -> None:
        running = self._is_roblox_running()
        if running:
            self.roblox_status_var.set("Roblox Found")
            self.roblox_status_label.config(fg="green")
        else:
            self.roblox_status_var.set("Roblox not Found")
            self.roblox_status_label.config(fg="red")

    def on_clear_cache(self) -> None:
        self.used_words.clear()
        if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
            self.update_used_words_list()

    def on_play_round(self) -> None:
        prefix = self.prefix_var.get().strip()
        if not prefix:
            messagebox.showwarning("Last Letter Helper", "Please enter starting letters.")
            return

        if not self.wordlist_loaded:
            messagebox.showwarning("Last Letter Helper", "Word list is still loading or failed.")
            return

        completion = self.find_completion(prefix)
        if completion is None:
            messagebox.showinfo("Last Letter Helper", "No word found.")
            return

        self.root.withdraw()
        self.prefix_var.set("")

        if self._is_roblox_running():
            self._focus_roblox_window()

        threading.Thread(target=self._type_after_delay, args=(completion,), daemon=True).start()

    def on_ctrl_enter(self, event):
        self.on_play_round()
        return "break"

    def show_used_words(self) -> None:
        if not hasattr(self, 'used_words_window') or not self.used_words_window.winfo_exists():
            self.used_words_window = tk.Toplevel(self.root)
            self.used_words_window.title("Used Words")
            self.used_words_window.resizable(True, True)
            self.used_words_window.attributes('-topmost', True)

            frame = ttk.Frame(self.used_words_window)
            frame.pack(fill='both', expand=True, padx=5, pady=5)

            scrollbar = ttk.Scrollbar(frame)
            scrollbar.pack(side='right', fill='y')

            self.used_words_listbox = tk.Listbox(
                frame,
                yscrollcommand=scrollbar.set,
                font=('Consolas', 10),
                width=30,
                height=15
            )
            self.used_words_listbox.pack(side='left', fill='both', expand=True)

            scrollbar.config(command=self.used_words_listbox.yview)

            self.used_words_count = tk.Label(
                self.used_words_window,
                text=f"Words used: {len(self.used_words)}",
                anchor='w'
            )
            self.used_words_count.pack(fill='x', padx=5, pady=(0, 5))

            close_button = ttk.Button(
                self.used_words_window,
                text="Close",
                command=self.used_words_window.destroy
            )
            close_button.pack(pady=(0, 5))

            self.used_words_window.update_idletasks()

        self.update_used_words_list()
        self.used_words_window.lift()
        self.used_words_window.focus_force()

    def update_used_words_list(self) -> None:
        if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
            self.used_words_listbox.delete(0, tk.END)

            for word in sorted(self.used_words, key=str.lower):
                self.used_words_listbox.insert(tk.END, word)

            self.used_words_count.config(text=f"Words used: {len(self.used_words)}")
            self.used_words_window.after(500, self.update_used_words_list)

    def _on_key_event(self, event):
        if event.event_type != "down":
            return

        key = event.name

        if len(key) == 1 and key.isalpha():
            self.live_buffer += key.lower()
            return

        if key in ("space", "enter"):
            prefix = self.live_buffer.strip()

            if len(prefix) > 0:
                completion = self.find_completion(prefix)
                if completion:
                    threading.Thread(
                        target=self._type_after_delay,
                        args=(completion,),
                        daemon=True
                    ).start()

            self.live_buffer = ""

    def _type_after_delay(self, completion: str) -> None:
        time.sleep(1.0)
        try:
            base = float(self.speed_var.get()) / 1000.0
            mode = self.typing_mode_var.get()

            last_cycle_time = base
            burst_counter = 0

            for ch in completion:
                if mode == "Normal":
                    delay = base
                elif mode == "Fixed Speed":
                    delay = base
                elif mode == "Jitter":
                    delay = max(0.005, base + random.uniform(-0.010, 0.010))
                elif mode == "Cycle":
                    last_cycle_time += 0.002
                    if last_cycle_time >= base + 0.020:
                        last_cycle_time = base - 0.020
                    delay = max(0.005, last_cycle_time)
                elif mode == "Burst":
                    burst_counter += 1
                    if burst_counter % 5 == 0:
                        delay = base * 2
                    else:
                        delay = max(0.005, base * 0.5)
                elif mode == "Random":
                    delay = random.uniform(0.005, base * 2)
                else:
                    delay = base

                keyboard.press_and_release(ch)
                time.sleep(delay)

            keyboard.send("enter")
            time.sleep(1.0)

        finally:
            self.root.after(0, self.root.deiconify)

            if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
                self.update_used_words_list()

def check_for_updates():
    try:
        response = requests.get(
            "https://raw.githubusercontent.com/elDziad00/LastLetter/main/version.txt",
            timeout=5
        )
        response.raise_for_status()
        latest_version = response.text.strip()
        
        if version.parse(latest_version) > version.parse(CURRENT_VERSION):
            if messagebox.askyesno(
                "Update Available",
                f"Version {latest_version} is available!\n"
                f"You are currently using version {CURRENT_VERSION}.\n\n"
                "Would you like to download the latest version?"
            ):
                webbrowser.open("https://github.com/elDziad00/LastLetter/releases/latest")
    except Exception as e:
        print(f"Error checking for updates: {e}")

def main():
    root = tk.Tk()
    app = LastLetterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
