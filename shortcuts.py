import json
import os
import subprocess
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
        self.paths_file = "paths.json"
        self.ngrok_process = None
        self.ngrok_url = None
        self.paths = self.load_paths()
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
    
    def load_paths(self):
        try:
            with open(self.paths_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def save_paths(self):
        with open(self.paths_file, 'w') as f:
            json.dump(self.paths, f, indent=4)

    def add_shortcut(self, name, program, arguments=""):
        if not name or not program:
            print("Error: Shortcut name and program path are required!")
            return
        
        self.shortcuts[name] = {
            'program': program,
            'arguments': arguments
        }
        self.save_shortcuts()
        print(f"Added shortcut: {name}")

    def remove_shortcut(self, name):
        if name in self.shortcuts:
            del self.shortcuts[name]
            self.save_shortcuts()
            print(f"Removed shortcut: {name}")
        else:
            print(f"Error: Shortcut '{name}' not found")

    def list_shortcuts(self):
        if not self.shortcuts:
            print("No shortcuts configured.")
            return
        
        print("\nAvailable Shortcuts:")
        for name, details in self.shortcuts.items():
            args = details.get('arguments', '')
            print(f"{name}: {details['program']} {args}")

    def handle_settings_command(self, args):
        if not args:
            print("\nSettings Commands:")
            print("settings add <name> <program> [arguments] - Add a new shortcut")
            print("settings remove <name>                    - Remove a shortcut")
            print("settings list                            - List all shortcuts")
            return False

        parts = args.split(maxsplit=2)
        subcommand = parts[0]

        if subcommand == "add" and len(parts) >= 3:
            name = parts[1]
            program_args = parts[2].split(maxsplit=1)
            program = program_args[0]
            arguments = program_args[1] if len(program_args) > 1 else ""
            self.add_shortcut(name, program, arguments)
        elif subcommand == "remove" and len(parts) == 2:
            self.remove_shortcut(parts[1])
        elif subcommand == "list":
            self.list_shortcuts()
        else:
            print("Invalid settings command. Type 'settings' for usage.")

    def start_ngrok(self):
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
            print("settings               - Manage shortcuts")
            print("list                   - Show all shortcuts")
            print("exit                   - Exit the program")
            print("help                   - Show this help message")
            return False

        if cmd == "list":
            self.list_shortcuts()
            return False

        if cmd == "settings":
            self.handle_settings_command(args)
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
        current_display = self.prompt + self.current_line
        
        if self.last_line_length > 0:
            sys.stdout.write('\r' + ' '* self.last_line_length + '\r')
        
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
                if char == b'H':  # Up arrow
                    if self.history and self.history_index < len(self.history) - 1:
                        self.history_index += 1
                        
                        sys.stdout.write('\r')
                        sys.stdout.write(' ' * (len(self.prompt) + len(self.current_line)))
                        sys.stdout.write('\r')
                        
                        self.current_line = self.history[-(self.history_index + 1)]
                        sys.stdout.write(self.prompt + self.current_line)
                        sys.stdout.flush()
                elif char == b'P':  # Down arrow
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
            
            if char == '\r':  # Enter key
                print()
                if self.current_line:
                    self.history.append(self.current_line)
                if self.handle_command(self.current_line):
                    break
                self.current_line = ""
                self.history_index = -1
                sys.stdout.write(self.prompt)
                sys.stdout.flush()
            elif char == '\b':  # Backspace
                if self.current_line:
                    self.current_line = self.current_line[:-1]
                    
                    sys.stdout.write('\r' + ' ' * (len(self.prompt) + len(self.current_line) + 1))
                    
                    sys.stdout.write('\r' + self.prompt + self.current_line)
                    sys.stdout.flush()
            else:  # Regular character input
                self.current_line += char
                sys.stdout.write(char)
                sys.stdout.flush()


def main():
    try:
        terminal = Terminal()
        terminal.run()
    except Exception as e:
        print(f"An error occurred: {e}")
        input("Press Enter to exit...")


if __name__ == "__main__":
    main()