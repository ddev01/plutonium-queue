"""
Plutonium Server Monitor and Auto-Connector

This script monitors Plutonium servers for IW5 (Call of Duty: Modern Warfare 3) and automatically
connects to a chosen server when a slot becomes available.

Key Features:
- Fetches server information from multiple API endpoints
- Displays a colored and formatted list of available servers
- Monitors server status and player counts in real-time
- Automatically connects to the chosen server when a slot is available
- Provides visual and audio feedback on connection status

Dependencies:
- requests, pyautogui, pygetwindow, pywinauto, winsound

Usage:
Run the script and follow the on-screen prompts to select a server API and choose a server to monitor.
Press 'r' at any time to return to the previous step.
"""

import re
import sys
import time
import msvcrt
import datetime
from typing import List, Dict, Optional, Any

import requests
import pyautogui
import pygetwindow as gw
from pywinauto import Application
import winsound

# Constants
SERVER_API_URLS = [
    "https://hgmserve.rs/api/server",
    "https://cod.gilletteclan.com/api/server",
]
CHECK_INTERVAL = 1

# ANSI color codes
CYAN = "\033[36m"
YELLOW = "\033[33m"
GREEN = "\033[32m"
RED = "\033[31m"
MAGENTA = "\033[35m"
BLUE = "\033[34m"
RESET_COLOR = "\033[0m"

COLOR_MAP = {
    "^0": "\033[30m",
    "^1": "\033[31m",
    "^2": "\033[32m",
    "^3": "\033[33m",
    "^4": "\033[34m",
    "^5": "\033[36m",
    "^6": "\033[35m",
    "^7": "\033[37m",
}

# Global variables
current_map: Optional[str] = None


def tprint(
    text: str, color: str = RESET_COLOR, end: str = "\n", add_timestamp: bool = True
) -> None:
    """Print text with time prefix, color support, and emojis."""
    if add_timestamp:
        now = datetime.datetime.now()
        current_time = now.strftime("[%H:%M:%S.") + f"{now.microsecond // 1000:03d}] "
        print(f"{CYAN}{current_time}{RESET_COLOR}", end="")
    print_colored_text(f"{color}{text}{RESET_COLOR}")
    print(end, end="")


def print_colored_text(text: str) -> None:
    """Print text with color codes."""
    parts = re.split(r"(\^[0-7])", text)
    current_color = RESET_COLOR
    for part in parts:
        if part in COLOR_MAP:
            current_color = COLOR_MAP[part]
        else:
            print(f"{current_color}{part}", end="")
    print(RESET_COLOR, end="")


def strip_color_codes(text: str) -> str:
    """Remove color codes from text for length calculation."""
    return re.sub(r"\^[0-7]", "", text)


def color_player_count(player_count: int, max_players: int) -> str:
    """Color the player count based on the number of players."""
    if player_count == max_players:
        color = COLOR_MAP["^1"]
    elif player_count >= max_players - 2:
        color = COLOR_MAP["^3"]
    elif player_count >= max_players * 0.7:
        color = COLOR_MAP["^2"]
    elif player_count >= max_players * 0.4:
        color = COLOR_MAP["^5"]
    elif player_count > 0:
        color = COLOR_MAP["^6"]
    else:
        color = COLOR_MAP["^7"]
    return f"{color}[{player_count:2}/{max_players:2}]{RESET_COLOR}"


def play_sound(sound_name: str = "Ring08") -> None:
    """Play a system sound for notifications."""
    try:
        winsound.PlaySound(r"C:\Windows\Media\Ring08.wav", winsound.SND_FILENAME)
    except Exception as e:
        tprint(f"Failed to play sound: {e}", RED)


def get_servers(api_url: str) -> Optional[List[Dict[str, Any]]]:
    """Fetch the server list from the API and filter for IW5 games."""
    try:
        response = requests.get(api_url, timeout=10)
        response.raise_for_status()
        servers = response.json()
        return [server for server in servers if server["game"] == "IW5"]
    except requests.RequestException as e:
        tprint(f"Error fetching server status: {e}", RED)
    return None


def find_plutonium_window() -> Optional[str]:
    """Find the Plutonium window with a title matching 'Plutonium r[any 4 numbers]'."""
    for title in gw.getAllTitles():
        if re.match(r"Plutonium r\d{4}", title):
            return title
    return None


def open_console() -> None:
    """Open and activate the Plutonium console window."""
    console_title = find_plutonium_window()
    if not console_title:
        tprint("No Plutonium console window found", RED)
        input("Open the console and press enter to continue...")
        open_console()
        return

    console_list = gw.getWindowsWithTitle(console_title)
    if not console_list:
        tprint(f"No '{console_title}' (console) found", RED)
        input("Open the console and press enter to continue...")
        open_console()
        return

    console = console_list[0]
    if console.isMinimized:
        console.restore()
    app = Application(backend="uia").connect(handle=console._hWnd)
    dlg = app.window(handle=console._hWnd)
    dlg.set_focus()


def connect_to_server(server: Dict[str, Any]) -> None:
    """Connect to the server using the Plutonium console."""
    open_console()
    time.sleep(0.1)
    pyautogui.typewrite(f"connect {server['listenAddress']}:{server['listenPort']}")
    pyautogui.press("enter")


def find_terminal_control(dlg: Any) -> Optional[Any]:
    """Find the control containing the terminal content."""
    for control in dlg.descendants():
        try:
            if (
                control.friendly_class_name() == "Static"
                and control.window_text().startswith("----------------------")
            ):
                return control
        except Exception as e:
            tprint(f"Error accessing control properties: {e}", RED)
    return None


def find_game_window() -> Optional[str]:
    """Find the Plutonium IW5: Multiplayer window."""
    for title in gw.getAllTitles():
        if "Plutonium IW5: Multiplayer" in title:
            return title
    return None


def capture_terminal_content() -> str:
    """Capture and monitor the terminal content for specific messages."""
    console_title = find_plutonium_window()
    if not console_title:
        raise Exception("Plutonium console window not found")

    window = gw.getWindowsWithTitle(console_title)[0]
    app = Application(backend="uia").connect(handle=window._hWnd)
    dlg = app.window(handle=window._hWnd)

    terminal_control = find_terminal_control(dlg)
    if not terminal_control:
        raise Exception("Terminal control not found")

    previous_content = terminal_control.window_text()
    try:
        while True:
            if msvcrt.kbhit():
                key = msvcrt.getch().decode("utf-8").lower()
                if key == "r":
                    return "return"

            current_content = terminal_control.window_text()
            if current_content != previous_content:
                new_lines = current_content[len(previous_content) :]
                previous_content = current_content

                if "not allowing server to override saved dvar" in new_lines:
                    tprint("✅ Successfully connected to the server!", GREEN)
                    game_window = find_game_window()
                    if game_window:
                        gw.getWindowsWithTitle(game_window)[0].activate()
                    play_sound("Ring08")
                    return "connected"
                elif "Com_ERROR: EXE_SERVERISFULL" in new_lines:
                    tprint(
                        f"🔒 Server is full, retrying in {CHECK_INTERVAL} seconds...",
                        YELLOW,
                    )
                    time.sleep(CHECK_INTERVAL)
                    return "retry"
            time.sleep(1)
    except Exception as e:
        tprint(f"❌ Error reading terminal content: {e}", RED)
        return "error"


def print_server_list(servers: List[Dict[str, Any]]) -> None:
    """Print the list of servers with their names colored and additional information aligned."""
    tprint("Available servers:", BLUE)
    max_name_length = max(
        len(strip_color_codes(server["serverName"])) for server in servers
    )
    max_map_length = max(len(server["currentMap"]["name"]) for server in servers)
    max_alias_length = max(len(server["currentMap"]["alias"]) for server in servers)

    for i, server in enumerate(servers, 1):
        player_count = server["clientNum"]
        max_players = server["maxClients"]
        map_name = server["currentMap"]["name"]
        map_alias = server["currentMap"]["alias"]

        index = f"{i:2}. "
        name = format_colored_text(server["serverName"], max_name_length)
        players = color_player_count(player_count, max_players)
        map_info = f"{map_name:{max_map_length}} - {map_alias:{max_alias_length}}"

        server_info = f"{index}{name} {players} - {map_info}"
        tprint(server_info, add_timestamp=False)


def format_colored_text(text: str, max_length: int) -> str:
    """Format colored text to a specific length."""
    stripped = strip_color_codes(text)
    padding = max_length - len(stripped)
    return text + (" " * padding)


def select_api_url() -> Optional[str]:
    """Prompt the user to select the API URL."""
    while True:
        print("Select the API URL:")
        for i, url in enumerate(SERVER_API_URLS, 1):
            tprint(f"{i}. {url}", add_timestamp=False)

        choice = input("Enter the number of the server API you want to use: ").lower()
        if choice == "r":
            return None
        try:
            choice = int(choice)
            if 1 <= choice <= len(SERVER_API_URLS):
                return SERVER_API_URLS[choice - 1]
            tprint("❗ Invalid choice. Please enter a valid number.", RED)
        except ValueError:
            tprint("❗ Invalid input. Please enter a number.", RED)
        print()


def select_server(servers: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Prompt the user to select a server from the list."""
    while True:
        tprint(
            "Enter the number of the server you want to connect to: ", YELLOW, end=""
        )
        choice = input().lower()
        if choice == "r":
            return None
        try:
            choice = int(choice)
            if 1 <= choice <= len(servers):
                return servers[choice - 1]
            tprint("❗ Invalid choice. Please enter a valid number.", RED)
        except ValueError:
            if choice != "r":
                tprint("❗ Invalid input. Please enter a number or 'r'.", RED)


def monitor_server(api_url: str, chosen_server: Dict[str, Any]) -> None:
    """Monitor the chosen server and attempt to connect when a slot is available."""
    global current_map

    player_count = chosen_server["clientNum"]
    max_players = chosen_server["maxClients"]
    map_name = chosen_server["currentMap"]["name"]
    map_alias = chosen_server["currentMap"]["alias"]
    current_map = map_name

    print_server_list([chosen_server])
    tprint(f"Selected server: {chosen_server['serverName']}", MAGENTA)
    tprint(f"[{player_count}/{max_players}] - {map_name} - {map_alias}", BLUE)
    tprint("Press 'r' to return to server selection.", YELLOW)

    last_status = None

    while True:
        if msvcrt.kbhit():
            key = msvcrt.getch().decode("utf-8").lower()
            if key == "r":
                return

        server_status = get_servers(api_url)
        if server_status:
            server = next(
                (s for s in server_status if s["id"] == chosen_server["id"]), None
            )
            if server:
                player_count = server["clientNum"]
                max_players = server["maxClients"]
                new_map = server["currentMap"]["name"]
                map_alias = server["currentMap"]["alias"]

                current_status = (
                    f"{player_count}/{max_players} - {new_map} - {map_alias}"
                )

                if new_map != current_map:
                    current_map = new_map
                    tprint(f"Map changed to: {new_map}", YELLOW)

                if player_count < max_players:
                    tprint("🎮 Slot available! Connecting to the server...", GREEN)
                    connect_to_server(server)
                    status = capture_terminal_content()
                    if status == "connected":
                        tprint(
                            "✅ Successfully connected to the server. Exiting.", GREEN
                        )
                        return
                    elif status == "retry":
                        continue
                    elif status == "return":
                        return
                elif current_status != last_status:
                    tprint(f"🔒 Server is full - [{current_status}]", RED)
                    last_status = current_status
            else:
                tprint("❌ Chosen server not found in status update.", RED)
        else:
            tprint("❌ Failed to fetch server status.", RED)
        time.sleep(CHECK_INTERVAL)


def main() -> None:
    """Main function to run the Plutonium Server Monitor and Auto-Connector."""
    tprint("Press 'r' to return to the previous step at any time.", YELLOW)
    try:
        while True:
            api_url = select_api_url()
            if api_url is None:
                continue

            servers = get_servers(api_url)
            if not servers:
                tprint("❌ Failed to fetch server list. Exiting.", RED)
                return
            print_server_list(servers)

            chosen_server = select_server(servers)
            if chosen_server is None:
                continue

            monitor_server(api_url, chosen_server)

    except KeyboardInterrupt:
        print("\n")
        tprint("🛑 Operation cancelled by user. Exiting...", RED)
    except Exception as e:
        tprint(f"An unexpected error occurred: {e}", RED)
    finally:
        tprint("😊 Goodbye!", GREEN)


if __name__ == "__main__":
    main()
