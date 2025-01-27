import json
import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
import sys
import win32gui
import win32con
import win32process
import time
import textwrap
import pyautogui
from openai import OpenAI
import msvcrt
import requests
import psutil
import re
import threading
from fuzzywuzzy import fuzz
import numpy as np
import cv2
import pytesseract

pyautogui.FAILSAFE = True
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def click_text_ocr(text_to_find, debug=False, min_ratio=70):
    try:
        screenshot = pyautogui.screenshot()
        img_np = np.array(screenshot)
        gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
        
        custom_config = r'--oem 3 --psm 11'
        data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
        
        matches = []
        
        if debug:
            print(f"\nSearching for: '{text_to_find}'")


        for i, text in enumerate(data['text']):
            if text.strip():
                
                simple_ratio = fuzz.ratio(text_to_find.lower(), text.lower())
                partial_ratio = fuzz.partial_ratio(text_to_find.lower(), text.lower())
                token_ratio = fuzz.token_set_ratio(text_to_find.lower(), text.lower())
                
                best_ratio = max(simple_ratio, partial_ratio, token_ratio)
                confidence = int(data['conf'][i])
                
                
                length_penalty = 0.3 if len(text.strip()) < 2 else 1.0
                length_bonus = min(len(text.strip()) / max(len(text_to_find), 1), 1.0)
                exact_bonus = 2.0 if text.lower() == text_to_find.lower() else 1.0
                
                match_score = best_ratio * length_penalty * length_bonus * exact_bonus
                
                if debug:
                    print(f"\nFound text: '{text}'")
                    print(f"Simple ratio: {simple_ratio}%")
                    print(f"Partial ratio: {partial_ratio}%")
                    print(f"Token ratio: {token_ratio}%")
                    print(f"Best ratio: {best_ratio}%")
                    print(f"Length penalty: {length_penalty}")
                    print(f"Length bonus: {length_bonus}")
                    print(f"Exact bonus: {exact_bonus}")
                    print(f"Final score: {match_score}")
                    print(f"OCR confidence: {confidence}%")
                    print(f"Position: x={data['left'][i]}, y={data['top'][i]}, w={data['width'][i]}, h={data['height'][i]}")
                
                if match_score >= min_ratio:
                    matches.append({
                        'text': text,
                        'score': match_score,
                        'ratio': best_ratio,
                        'confidence': confidence,
                        'x': data['left'][i],
                        'y': data['top'][i],
                        'w': data['width'][i],
                        'h': data['height'][i]
                    })

        matches.sort(key=lambda x: (x['score'], x['confidence']), reverse=True)
        
        if matches:
            best_match = matches[0]
            if debug:
                print(f"\nBest match: '{best_match['text']}'")
                print(f"Match score: {best_match['score']}")
                print(f"Original ratio: {best_match['ratio']}%")
                print(f"OCR confidence: {best_match['confidence']}%")
            
            click_x = best_match['x'] + best_match['w']//2
            click_y = best_match['y'] + best_match['h']//2
            
            if debug:
                print(f"Clicking at ({click_x}, {click_y})")
            
            pyautogui.click(click_x, click_y)
            return True
                
        return False
        
    except Exception as e:
        print(f"Error: {e}")
        return False

class Terminal:
    def __init__(self):
        self.shortcuts_file = "shortcuts.json"
        self.ngrok_process = None
        self.ngrok_url = None
        self.shortcuts = self.load_shortcuts()
        self.history = []
        self.history_index = -1
        self.current_line = ""
        self.prompt = ">> "
        self.max_attempts = 32
        
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
                self.openai_api_key = config.get('openai_api_key')
        except FileNotFoundError:
            self.openai_api_key = None
        
        self.client = None
        if self.openai_api_key:
            self.client = OpenAI(api_key=self.openai_api_key)

    def load_shortcuts(self):
        try:
            with open(self.shortcuts_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def save_shortcuts(self):
        with open(self.shortcuts_file, 'w') as f:
            json.dump(self.shortcuts, f, indent=4)

    def print_prompt(self, line=""):
        sys.stdout.write('\r' + ' ' * (len(self.prompt) + 80))  
        sys.stdout.write('\r' + self.prompt + line)
        sys.stdout.flush()

    def close_other_windows(self):
        """Close all windows except the current process"""
        our_pid = os.getpid()
        
        def enum_windows_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            
            
            if not win32gui.GetWindowText(hwnd):
                return
                
            try:
                
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                
                
                if pid == our_pid:
                    return
                    
                
                process = psutil.Process(pid)
                if process.name().lower() in ['explorer.exe', 'taskmgr.exe']:
                    return
                    
                
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        
        win32gui.EnumWindows(enum_windows_callback, None)

    def open_app(self, command):
            pyautogui.hotkey('win', 's')
            pyautogui.write(command)
            pyautogui.press('enter') 
            
    def start_ngrok(self):
        """Start ngrok and get URL via its API"""
        try:
            self.ngrok_process = subprocess.Popen(
                ["C:\\Users\\Aayush\\Downloads\\ngrok-v3-stable-windows-amd64\\ngrok.exe", "http", "9284"],
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            print("Started ngrok, getting URL...")
            time.sleep(2)
            
            try:
                response = requests.get("http://localhost:4040/api/tunnels")
                tunnels = response.json()["tunnels"]
                if tunnels:
                    self.ngrok_url = tunnels[0]["public_url"]
                    print(f"\nNgrok URL: {self.ngrok_url}")
                else:
                    print("No active tunnels found")
            except requests.exceptions.ConnectionError:
                print("Couldn't connect to ngrok API. Is ngrok running?")
            except Exception as e:
                print(f"Error getting ngrok URL: {str(e)}")
                
        except Exception as e:
            print(f"Error starting ngrok: {str(e)}")

    def stop_ngrok(self):
        """Stop all ngrok processes"""
        try:
            
            if self.ngrok_process:
                self.ngrok_process.terminate()
                self.ngrok_process = None
                
            
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'ngrok' in proc.info['name'].lower():
                        process = psutil.Process(proc.info['pid'])
                        process.kill()  
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                
            self.ngrok_url = None
            print("Stopped all ngrok processes")
            
        except Exception as e:
            print(f"Error stopping ngrok: {str(e)}")

    def handle_command(self, command):
        if not command:
            return False

        
        parts = command.strip().split(maxsplit=1)
        cmd = parts[0]
        args = parts[1] if len(parts) > 1 else ""
        if cmd == "A" or cmd == "a":
            print("2")
            return False

        if "open" in command:
            self.open_app(command.split(" ")[1])
        if cmd == "exit":
            return True

        if cmd == "rivals":
            subprocess.Popen("C:\\Program Files (x86)\\Epic Games\\Launcher\\Portal\\Binaries\\Win64\\EpicGamesLauncher.exe")
            status = False
            counter = 0
            while(status == False):
                counter+=1
                status = click_text_ocr("Marvel")
                if(counter >= self.max_attempts):
                    print(f"Max attempts reached.")
                    return False

        if cmd == "fortnite":
            subprocess.Popen("C:\\Program Files (x86)\\Epic Games\\Launcher\\Portal\\Binaries\\Win64\\EpicGamesLauncher.exe")
            status = False
            while(status == False):
                status = click_text_ocr("Fortnite")

        if cmd == "ngrok":
            if self.ngrok_process:
                print("Ngrok is already running!")
                if self.ngrok_url:
                    print(f"Current URL: {self.ngrok_url}")
            else:
                self.start_ngrok()
            return False

        if cmd == "ngrok-stop":
            self.stop_ngrok()
            print("Done.")
            return False

        
        if cmd == "reset":
            self.stop_ngrok()
            self.close_other_windows()
            print("Reset.")

        if cmd == "chrome":
            if not args:
                print("Usage: chrome [profile_number] [optional_url]")
                return False
            
            chrome_parts = args.split(maxsplit=1)
            
            profile_num = chrome_parts[0]

            
            url = chrome_parts[1] if len(chrome_parts) > 1 else ""
            
            try:
                chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                if profile_num == "0":
                    profile_arg = "--profile-directory=Default"
                else:
                    profile_arg = f"--profile-directory=Profile {profile_num}"
                
                if url:
                    subprocess.Popen([chrome_path, profile_arg, url])
                    print(f"Opening Chrome Profile {profile_num} with URL: {url}")
                else:
                    subprocess.Popen([chrome_path, profile_arg])
                    print(f"Opening Chrome Profile {profile_num}")
                
                time.sleep(1)
                print("Done.")
                self.maximize_last_window()
                return False
            except Exception as e:
                print(f"Error opening Chrome: {str(e)}")
                return False

        if cmd == "code" and args:
            
            folder_path = os.path.join("C:\\Users\\Aayush", args)
            try:
                subprocess.Popen([
                    "C:\\Users\\Aayush\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
                    folder_path
                ])
                print("Done.")
                time.sleep(1)
                self.maximize_last_window()
                return False
            except Exception as e:
                print(f"Error opening VSCode: {str(e)}")
                return False

        
        
        if cmd.startswith("setapi "):
            api_key = command[7:].strip()
            if api_key:
                config = {'openai_api_key': api_key}
                with open('config.json', 'w') as f:
                    json.dump(config, f)
                self.openai_api_key = api_key
                self.client = OpenAI(api_key=api_key)
                print("API key saved successfully")
            return False

        if cmd == "school":
            
            try:
                chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                
                profile_arg = f"--profile-directory=Profile 1"
                
                subprocess.Popen([chrome_path, profile_arg])
                
                
                time.sleep(1)
                self.maximize_last_window()
                time.sleep(3)
                click_text_ocr("Aayush")
                print("Done.")
                return False
            except Exception as e:
                print(f"Error opening Chrome: {str(e)}")
                return False

        if cmd == "help":
            print("\nCommands:")
            print("chrome [profile] [url] - Open Chrome with specific profile and optional URL")
            print("code [folder]          - Open VSCode in specified folder under C:/Users/<Username>}")
            print("ngrok                  - Open an ngrok tunnel")
            print("ngrok-stop             - Kill all active ngrok tunnels")
            print("gpt: [message]         - Send a message to GPT-4")
            print("setapi [key]           - Set your OpenAI API key")
            print("settings               - Open shortcut settings")
            print("list                  - Show all shortcuts")
            print("exit                  - Exit the program")
            print("help                  - Show this help message")
            return False

        if cmd == "list":
            if not self.shortcuts:
                print("No shortcuts configured. Use 'settings' to add some.")
                return False
            print("Available Shortcuts:")
            for name, details in self.shortcuts.items():
                args = details.get('arguments', '')
                print(f"{name}: {details['program']} {args}")
            return False

        if cmd == "settings":
            app = ShortcutGUI(self.shortcuts, self.shortcuts_file)
            app.run()
            self.shortcuts = self.load_shortcuts()
            return False

        if cmd.startswith("gpt:"):
            message = command[4:].strip()
            if not message:
                print("Please provide a message for GPT")
                return False
            
            if not self.client:
                print("Please set your OpenAI API key first using 'setapi YOUR_API_KEY'")
                return False
            
            try:
                response = self.client.chat.completions.create(
                    model="gpt-4",
                    messages=[
                        {"role": "system", "content": "Answer the following question to the best of your ability. Put no text before and after the answer, just provide the answer."},
                        {"role": "user", "content": message}
                    ]
                )
                print("\nGPT:", end="\n\n")
                wrapped_text = textwrap.fill(response.choices[0].message.content, width=50)
                print(wrapped_text + "\n")
            except Exception as e:
                print(f"Error communicating with GPT: {str(e)}")
            return False

        if cmd in self.shortcuts:
            shortcut = self.shortcuts[cmd]
            try:
                if shortcut.get('arguments'):
                    subprocess.Popen([shortcut['program'], shortcut['arguments']], shell=True)
                else:
                    subprocess.Popen([shortcut['program']], shell=True)
                print(f"Executed: {cmd}")
                time.sleep(1)
                self.maximize_last_window()
            except Exception as e:
                print(f"Error executing {cmd}: {str(e)}")
        else:
            print(f"Unknown command: {cmd}")
            print('Type "help" to see available commands')
        return False

    def maximize_last_window(self):
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                windows.append(hwnd)
            return True

        windows = []
        win32gui.EnumWindows(callback, windows)
        
        if windows:
            last_window = windows[-1]
            win32gui.ShowWindow(last_window, win32con.SW_MAXIMIZE)
    def update_display(self):
        """
        Updates the terminal display, handling the prompt and current line.
        Clears the previous line completely before writing new content.
        """
        
        current_display = self.prompt + self.current_line
        
        
        if self.last_line_length > 0:
            sys.stdout.write('\r' + ' ' * self.last_line_length + '\r')
        
        
        sys.stdout.write(current_display)
        sys.stdout.flush()
        
        
        self.last_line_length = len(current_display)

    def run(self):
        print('Two v1.0 - type "help" for commands')
        sys.stdout.write(self.prompt)
        sys.stdout.flush()
        
        while True:
            char = msvcrt.getch()
            
            if char in [b'\x00', b'\xe0']:  
                char = msvcrt.getch()
                if char == b'H':  
                    if self.history and self.history_index < len(self.history) - 1:
                        self.history_index += 1
                        
                        sys.stdout.write('\r')
                        sys.stdout.write(' ' * (len(self.prompt) + len(self.current_line)))
                        sys.stdout.write('\r')
                        
                        self.current_line = self.history[-(self.history_index + 1)]
                        sys.stdout.write(self.prompt + self.current_line)
                        sys.stdout.flush()
                elif char == b'P':  
                    
                    sys.stdout.write('\r')
                    sys.stdout.write(' ' * (len(self.prompt) + len(self.current_line)))
                    sys.stdout.write('\r')
                    
                    if self.history_index > 0:
                        self.history_index -= 1
                        self.current_line = self.history[-(self.history_index + 1)]
                        sys.stdout.write(self.prompt + self.current_line)
                    elif self.history_index == 0:
                        self.history_index = -1
                        self.current_line = ""
                        sys.stdout.write(self.prompt)
                    sys.stdout.flush()
                continue
                
            char = char.decode('utf-8', errors='ignore')
            
            if char == '\r':  
                print()  
                if self.current_line:
                    self.history.append(self.current_line)
                if self.handle_command(self.current_line):
                    break
                self.current_line = ""
                self.history_index = -1
                sys.stdout.write(self.prompt)
                sys.stdout.flush()
            elif char == '\b':  
                if self.current_line:
                    self.current_line = self.current_line[:-1]
                    
                    sys.stdout.write('\r' + ' ' * (len(self.prompt) + len(self.current_line) + 1))
                    
                    sys.stdout.write('\r' + self.prompt + self.current_line)
                    sys.stdout.flush()
            else:  
                self.current_line += char
                sys.stdout.write(char)
                sys.stdout.flush()
class ShortcutGUI:
    def __init__(self, shortcuts, shortcuts_file):
        self.root = tk.Tk()
        self.root.title("Quick Launch Settings")
        self.root.geometry("600x400")
        
        self.shortcuts_file = shortcuts_file
        self.shortcuts = shortcuts
        
        
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        
        self.create_input_frame()
        
        
        self.create_shortcuts_list()

    def create_input_frame(self):
        input_frame = ttk.LabelFrame(self.main_frame, text="Add New Shortcut", padding="10")
        input_frame.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
        
        
        ttk.Label(input_frame, text="Shortcut Name:").grid(row=0, column=0, padx=5, pady=5)
        self.shortcut_name = ttk.Entry(input_frame)
        self.shortcut_name.grid(row=0, column=1, padx=5, pady=5)
        
        
        ttk.Label(input_frame, text="Program Path:").grid(row=1, column=0, padx=5, pady=5)
        self.program_path = ttk.Entry(input_frame)
        self.program_path.grid(row=1, column=1, padx=5, pady=5)
        
        
        ttk.Label(input_frame, text="Arguments (optional):").grid(row=2, column=0, padx=5, pady=5)
        self.arguments = ttk.Entry(input_frame)
        self.arguments.grid(row=2, column=1, padx=5, pady=5)
        
        
        ttk.Button(input_frame, text="Add Shortcut", command=self.add_shortcut).grid(row=3, column=0, columnspan=2, pady=10)

    def create_shortcuts_list(self):
        
        columns = ('Shortcut', 'Program', 'Arguments')
        self.tree = ttk.Treeview(self.main_frame, columns=columns, show='headings')
        
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        
        self.tree.grid(row=1, column=0, padx=5, pady=5, sticky='nsew')
        
        
        scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        
        self.update_shortcuts_list()
        
        
        ttk.Button(self.main_frame, text="Delete Selected", command=self.delete_shortcut).grid(row=2, column=0, pady=5)

    def update_shortcuts_list(self):
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        
        for shortcut, details in self.shortcuts.items():
            self.tree.insert('', 'end', values=(shortcut, details['program'], details.get('arguments', '')))

    def add_shortcut(self):
        name = self.shortcut_name.get().strip()
        program = self.program_path.get().strip()
        args = self.arguments.get().strip()
        
        if not name or not program:
            messagebox.showerror("Error", "Shortcut name and program path are required!")
            return
        
        self.shortcuts[name] = {
            'program': program,
            'arguments': args
        }
        
        self.save_shortcuts()
        self.update_shortcuts_list()
        
        
        self.shortcut_name.delete(0, tk.END)
        self.program_path.delete(0, tk.END)
        self.arguments.delete(0, tk.END)

    def delete_shortcut(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a shortcut to delete!")
            return
        
        shortcut = self.tree.item(selected_item[0])['values'][0]
        if shortcut in self.shortcuts:
            del self.shortcuts[shortcut]
            self.save_shortcuts()
            self.update_shortcuts_list()

    def save_shortcuts(self):
        with open(self.shortcuts_file, 'w') as f:
            json.dump(self.shortcuts, f, indent=4)

    def run(self):
        self.root.mainloop()

def main():
    try:
        terminal = Terminal()
        terminal.run()
    except Exception as e:
        print(f"An error occurred: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()