import re
import requests
import time
import pyautogui
import pygetwindow as gw
from pywinauto import Application
import datetime
import os
import winsound
import sys

CHECK_INTERVAL = 0.5

# ANSI escape codes for colors
COLOR_MAP = {
    "^0": "\033[30m",  # Black
    "^1": "\033[31m",  # Red
    "^2": "\033[32m",  # Green
    "^3": "\033[33m",  # Yellow
    "^4": "\033[34m",  # Blue
    "^5": "\033[36m",  # Cyan (Light Blue)
    "^6": "\033[35m",  # Magenta (Pink)
    "^7": "\033[37m",  # White
}

CYAN = "\033[36m"
YELLOW = "\033[33m"
GREEN = "\033[32m"
RED = "\033[31m"
MAGENTA = "\033[35m"
BLUE = "\033[34m"
RESET_COLOR = "\033[0m"

current_map = None


def play_sound(sound_name="Ring08"):
    try:
        winsound.PlaySound(r"C:\Windows\Media\Ring08.wav", winsound.SND_FILENAME)
    except:
        pass


def get_servers(api_url):
    """Fetch the server list from the API."""
    try:
        response = requests.get(api_url)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        tprint(f"Error fetching server status: {e}")
    return None


def find_plutonium_window():
    """Find the Plutonium window with a title matching 'Plutonium r[any 4 numbers]'."""
    windows = gw.getAllTitles()
    for title in windows:
        if re.match(r"Plutonium r\d{4}", title):
            return title
    return None


def open_console():
    """Open and activate the Plutonium console window."""
    console_title = find_plutonium_window()
    if console_title:
        console_list = gw.getWindowsWithTitle(console_title)
        if console_list:
            console = console_list[0]
            if console.isMinimized:
                console.restore()
            app = Application(backend="uia").connect(handle=console._hWnd)
            dlg = app.window(handle=console._hWnd)
            dlg.set_focus()
            # console.maximize()
        else:
            tprint(f"No '{console_title}' (console) found")
            input("Open the console and press enter to continue...")
            open_console()
    else:
        tprint("No Plutonium console window found")
        input("Open the console and press enter to continue...")
        open_console()


def connect_to_server(server):
    """Connect to the server using the Plutonium console."""
    open_console()
    time.sleep(0.1)
    pyautogui.typewrite(f"connect {server['listenAddress']}:{server['listenPort']}")
    pyautogui.press("enter")


def find_terminal_control(dlg):
    """Find the control containing the terminal content."""
    for control in dlg.descendants():
        try:
            control_text = control.window_text()
            control_class = control.friendly_class_name()
            if control_class == "Static" and control_text.startswith(
                "----------------------"
            ):
                return control
        except Exception as e:
            tprint(f"Error accessing control properties: {e}")
    return None


def find_game_window():
    """Find the Plutonium IW5: Multiplayer window."""
    windows = gw.getAllTitles()
    for title in windows:
        if "Plutonium IW5: Multiplayer" in title:
            return title
    return None


def capture_terminal_content():
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
            current_content = terminal_control.window_text()
            if current_content != previous_content:
                new_lines = current_content[len(previous_content) :]
                previous_content = current_content

                if "not allowing server to override saved dvar" in new_lines:
                    tprint("✅ Successfully connected to the server!", GREEN)
                    game_window = find_game_window()
                    if game_window:
                        game_window_obj = gw.getWindowsWithTitle(game_window)[0]
                        game_window_obj.activate()
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
    except KeyboardInterrupt:
        tprint("🛑 Stopped reading terminal content.", RED)
    except Exception as e:
        tprint(f"❌ Error reading terminal content: {e}", RED)


def print_server_list(servers):
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

        # Format each part separately
        index = f"{i:2}. "
        name = format_colored_text(server["serverName"], max_name_length)
        players = color_player_count(player_count, max_players)
        map_info = f"{map_name:{max_map_length}} - {map_alias:{max_alias_length}}"

        # Combine all parts
        server_info = f"{index}{name} {players} - {map_info}"

        # Print the formatted string
        tprint(server_info, add_timestamp=False)


def format_colored_text(text, max_length):
    """Format colored text to a specific length."""
    stripped = strip_color_codes(text)
    padding = max_length - len(stripped)
    return text + (" " * padding)


def tprint(text, color=RESET_COLOR, end="\n", add_timestamp=True):
    """Print text with time prefix, color support, and emojis."""
    if add_timestamp:
        now = datetime.datetime.now()
        current_time = now.strftime("[%H:%M:%S.") + f"{now.microsecond // 1000:03d}] "
        print(f"{CYAN}{current_time}{RESET_COLOR}", end="")
    print_colored_text(f"{color}{text}{RESET_COLOR}")
    print(end, end="")


def print_colored_text(text):
    """Print text with color codes."""
    parts = re.split(r"(\^[0-7])", text)
    current_color = RESET_COLOR
    for part in parts:
        if part in COLOR_MAP:
            current_color = COLOR_MAP[part]
        else:
            print(f"{current_color}{part}", end="")
    print(RESET_COLOR, end="")


def strip_color_codes(text):
    """Remove color codes from text for length calculation."""
    return re.sub(r"\^[0-7]", "", text)


def color_player_count(player_count, max_players):
    """Color the player count based on the number of players."""
    if player_count == max_players:
        color = COLOR_MAP["^1"]  # Red for full server
    elif player_count >= max_players - 2:
        color = COLOR_MAP["^3"]  # Yellow for nearly full (1-2 spots open)
    elif player_count >= max_players * 0.7:  # 13+ players in an 18-player server
        color = COLOR_MAP["^2"]  # Green for high player count (exciting games)
    elif player_count >= max_players * 0.4:  # 8-12 players in an 18-player server
        color = COLOR_MAP["^5"]  # Cyan for medium player count (decent games)
    elif player_count > 0:
        color = COLOR_MAP["^6"]  # Magenta for low player count (not very exciting)
    else:
        color = COLOR_MAP["^7"]  # White for empty server

    return f"{color}[{player_count:2}/{max_players:2}]{RESET_COLOR}"


def select_api_url():
    """Prompt the user to select the API URL."""
    print("Select the API URL:")
    print("1. https://hgmserve.rs/api/server")
    print("2. https://cod.gilletteclan.com/api/server")
    choice = input("Enter the number of the API URL you want to use: ")
    if choice == "1":
        return "https://hgmserve.rs/api/server"
    elif choice == "2":
        return "https://cod.gilletteclan.com/api/server"
    else:
        print("Invalid choice. Defaulting to https://cod.gilletteclan.com/api/server")
        return "https://cod.gilletteclan.com/api/server"


def get_servers(api_url):
    """Fetch the server list from the API and filter for IW5 games."""
    try:
        response = requests.get(api_url)
        response.raise_for_status()
        servers = response.json()
        return [server for server in servers if server["game"] == "IW5"]
    except requests.RequestException as e:
        tprint(f"Error fetching server status: {e}")
    return None


def main():
    global current_map

    api_url = select_api_url()
    servers = get_servers(api_url)
    if not servers:
        tprint("❌ Failed to fetch server list. Exiting.", RED)
        return
    print_server_list(servers)

    chosen_server = None
    while chosen_server is None:
        try:
            tprint(
                "Enter the number of the server you want to connect to: ",
                YELLOW,
                end="",
            )
            choice = int(input())
            if 1 <= choice <= len(servers):
                chosen_server = servers[choice - 1]
                player_count = chosen_server["clientNum"]
                max_players = chosen_server["maxClients"]
                map_name = chosen_server["currentMap"]["name"]
                map_alias = chosen_server["currentMap"]["alias"]
                current_map = map_name
            else:
                tprint("❗ Invalid choice. Please enter a valid number.", RED)
        except ValueError:
            tprint("❗ Invalid input. Please enter a number.", RED)
        except KeyboardInterrupt:
            print()  # Add a newline before the CTRL+C message
            tprint("🛑 CTRL+C detected. Exiting...", RED)
            tprint("😊 Goodbye!", GREEN)
            exit()

    # Clear the screen and move cursor to home position
    sys.stdout.write("\033[2J\033[H")
    sys.stdout.flush()

    # Reprint the server list and selection
    print_server_list(servers)
    tprint(f"Selected server: {chosen_server['serverName']}", MAGENTA)
    tprint(f"[{player_count}/{max_players}] - {map_name} - {map_alias}", BLUE)

    last_status = None

    try:
        while True:
            server_status = get_servers(api_url)
            if server_status:
                for server in server_status:
                    if server["id"] == chosen_server["id"]:
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
                            tprint(
                                "🎮 Slot available! Connecting to the server...", GREEN
                            )
                            connect_to_server(server)
                            status = capture_terminal_content()
                            if status == "connected":
                                tprint(
                                    "✅ Successfully connected to the server. Exiting.",
                                    GREEN,
                                )
                                return
                            elif status == "retry":
                                continue
                        else:
                            if current_status != last_status:
                                tprint(f"🔒 Server is full - [{current_status}]", RED)
                                last_status = current_status
                        break
            else:
                tprint("❌ Failed to fetch server status.", RED)
            time.sleep(CHECK_INTERVAL)
    except KeyboardInterrupt:
        print("\n")  # Move to a new line
        tprint("🛑 CTRL+C detected. Exiting...", RED)
        tprint("😊 Goodbye!", GREEN)
    finally:
        sys.stdout.write("\033[?25h")  # Show cursor
        sys.stdout.flush()


if __name__ == "__main__":
    main()
