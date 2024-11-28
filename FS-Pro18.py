"""
File Search Pro
Copyright (C) 2024 [Kristopher Sorensen]

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.
"""

import os
import json
import sys
import shutil
import win32com.client
import threading
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QListWidget, QPushButton, QProgressBar,
    QVBoxLayout, QWidget, QMessageBox, QFileDialog, QComboBox, QMenu, QInputDialog, QListWidgetItem, QTextBrowser, QDialog, QMenuBar
)
from PyQt5.QtCore import pyqtSignal, QObject, Qt
from PyQt5.QtGui import QIcon
from filelock import FileLock, Timeout
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import shutil

# Path to save the JSON index file in the same directory as the script or .exe
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

INDEX_FILE = os.path.join(APP_DIR, "file_index.json")
ICON_FILE = os.path.join(APP_DIR, "FS-ICO.ico")
TAGS_FILE = os.path.join(APP_DIR, "tags.json")  # File to store tags

DARK_MODE_STYLESHEET = """
    QMainWindow {
        background-color: #2b2b2b;
        color: #ffffff;
    }
    QLabel, QLineEdit, QListWidget, QPushButton, QComboBox {
        color: #ffffff;
        background-color: #3c3f41;
        border: 1px solid #555555;
    }
    QPushButton:hover {
        background-color: #555555;
    }
    QProgressBar {
        background-color: #3c3f41;
        color: #ffffff;
        border: 1px solid #555555;
        text-align: center;
    }
    QProgressBar::chunk {
        background-color: #85c1e9;
    }
"""

LIGHT_MODE_STYLESHEET = """
    QMainWindow {
        background-color: #f0f0f0;
        color: #000000;
    }
    QLabel, QLineEdit, QListWidget, QPushButton, QComboBox {
        color: #000000;
        background-color: #ffffff;
        border: 1px solid #cccccc;
    }
    QPushButton:hover {
        background-color: #e6e6e6;
    }
    QProgressBar {
        background-color: #ffffff;
        color: #000000;
        border: 1px solid #cccccc;
        text-align: center;
    }
    QProgressBar::chunk {
        background-color: #4caf50;
    }
"""
# Help menu items
class HelpDialog(QDialog):
    """Custom dialog for displaying the Help information."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Help")
        self.setGeometry(200, 200, 600, 400)
        # Remove the "?" button by adjusting the window flags
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)

        # Help content (update this content with the desired information)
        help_content = """
        <h2>File Search Pro - Help</h2>
        <ul>
            <li><b>Adding Directories:</b> Use the "Add Directory" button to select and index a directory.</li>
            <li><b>Dynamic Updates are performed:</b>  If a file is removed,renamed or added to the directory you are indexing, it is automatically updated.</li>
            <li><b>Refresh Index:</b> Use the "Refresh Index" button to refresh the index in the current directory to refresh your results list.</li>
            <li><b>Deleting Directories:</b> Use the "Delete Directory" button to remove the current directory from the index and monitoring.</li>
            <li><b>Filters:</b> Use the common file type and dev/engineering filters to narrow down your search results.</li>
            <li><b>Real-Time Search:</b> Start typing in the search bar to filter the results based on file names.</li>
            <li><b>Clear Search:</b> Use the "Clear Search" button to clear your search.</li>
            <li><b>Add Tags:</b> Right-click on a file to add or edit tags. Tagged files are displayed with a yellow highlight.</li>
            <li><b>Tags Info:</b> If someone renames a file in the original directory then your tag will disappear from that file in your index search results.</li>
            <li><b>Remove Tags:</b> Right-click on a file to add or edit tags. Make it blank then hit ok to remove the tag.</li>
            <li><b>Opening Files:</b> Double-click a file in the list to open it with the default application.</li>
            <li><b>Saving Files:</b> Right-click on a file and choose "Save As" to save it to a different location.</li>
            <li><b>Email Files:</b> Right-click on a file and choose "Send As Email" to open and attach in an outlook email. Only works with Outlook...</li>
            <li><b>Dark Mode:</b> Toggle dark mode using the "View Mode" menu.</li>
            <li><b>Exclusions:</b> The application automatically excludes certain system directories like C:\\Windows and file types like .ini, .exe, .dll, .reg, etc.</li>
        </ul>
        """

        # Add help content
        help_text = QTextBrowser(self)
        help_text.setHtml(help_content)
        help_text.setReadOnly(True)
        layout.addWidget(help_text)

        # Close button
        close_button = QPushButton("Close", self)
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

# Tag features
class TagManager:
    def __init__(self):
        self.tags = {}
        self.load_tags()

    def load_tags(self):
        """Loads tags from the JSON file."""
        if os.path.exists(TAGS_FILE):
            with open(TAGS_FILE, "r") as f:
                self.tags = json.load(f)
        else:
            self.tags = {}

    def save_tags(self):
        """Saves tags to the JSON file."""
        with open(TAGS_FILE, "w") as f:
            json.dump(self.tags, f, indent=4)

    def add_tag(self, file_path, tag):
        """Adds a tag to a file."""
        if file_path not in self.tags:
            self.tags[file_path] = []
        if tag not in self.tags[file_path]:
            self.tags[file_path].append(tag)
            self.save_tags()

    def remove_tag(self, file_path, tag):
        """Removes a tag from a file."""
        if file_path in self.tags and tag in self.tags[file_path]:
            self.tags[file_path].remove(tag)
            if not self.tags[file_path]:  # Remove file if no tags left
                del self.tags[file_path]
            self.save_tags()

    def get_tags(self, file_path):
        """Gets tags for a file."""
        return self.tags.get(file_path, [])
    
# Main application window event handler
class FileMonitorHandler(FileSystemEventHandler):
    def __init__(self, update_callback, remove_callback, rename_callback):
        self.update_callback = update_callback
        self.remove_callback = remove_callback
        self.rename_callback = rename_callback

    def on_created(self, event):
        if not event.is_directory:
            self.update_callback(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self.update_callback(event.src_path)

    def on_deleted(self, event):
        if not event.is_directory:
            self.remove_callback(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self.rename_callback(event.src_path, event.dest_path)

# Signal method to update progress
class WorkerSignals(QObject):
    progress = pyqtSignal(int)  # To update progress bar
    indexing_complete = pyqtSignal()  # To notify when indexing is complete

# Main window layout, features
class FileSearcherApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Search Pro Version 1.18")
        self.setGeometry(100, 100, 900, 500)
        self.tag_manager = TagManager()
        # Initialize dark mode tracking
        self.dark_mode_enabled = False
        self.directories = []  # Store up to 7 directory paths
        self.current_directory = None  # Currently selected directory
        self.files = set()
        self.last_modified_time = 0
        self.lock = FileLock(f"{INDEX_FILE}.lock")
        self.signals = WorkerSignals()
        self.observer = None  # Watchdog observer for monitoring
        self.files_lock = threading.Lock()  # Thread-safe lock for self.files

        # Connect signals to GUI update methods
        self.signals.progress.connect(self.update_progress_bar)
        self.signals.indexing_complete.connect(self.on_indexing_complete)

        if os.path.exists(ICON_FILE):
            self.setWindowIcon(QIcon(ICON_FILE))

        # Central widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Menu bar
        menu_bar = self.menuBar()
        # Add "View" menu for toggling themes
        view_menu = QMenu("View Mode", self)
        menu_bar.addMenu(view_menu)

        # Add "Toggle Dark Mode" action
        toggle_dark_mode_action = view_menu.addAction("Toggle Dark Mode")
        toggle_dark_mode_action.triggered.connect(self.toggle_dark_mode)

        # Help menu
        help_menu = QMenu("Help", self)
        menu_bar.addMenu(help_menu)

        # Add "View Help" action
        view_help_action = help_menu.addAction("View Help")
        view_help_action.triggered.connect(self.show_help)

        # Add Filter Dropdown
        self.filter_dropdown = QComboBox(self)
        self.filter_dropdown.addItem("Common Files Filter")  # Default option to show all files
        self.filter_dropdown.addItems([".pdf", ".jpg", ".jpeg", ".doc", ".docx", ".xls", ".xlsx", ".png", ".ppt", ".pptx"])
        self.filter_dropdown.currentIndexChanged.connect(self.apply_filter)  # Connect dropdown change to filter logic
        self.layout.addWidget(self.filter_dropdown)

        # Add Dev/Eng Filter Dropdown
        self.dev_filter_dropdown = QComboBox(self)
        self.dev_filter_dropdown.addItem("Dev/Eng Files Filter")  # Default option to show all files
        self.dev_filter_dropdown.addItems([
            ".js", ".py", ".java", ".cpp", ".c", ".cs", ".rb", ".php", ".html", ".css",
            ".scss", ".ts", ".go", ".swift", ".kt", ".rs", ".sql", ".sh", ".bat", ".pl",
            ".xml", ".json", ".yaml", ".yml", ".lua", ".asp", ".jsp", ".md", ".vue", ".jsx", ".tsx",
            ".dwg", ".dxf", ".step", ".stp", ".iges", ".igs", ".prt", ".asm", ".sldprt",
            ".sldasm", ".slddrw", ".stl", ".sch", ".brd", ".pcb", ".sp", ".dip", ".vsm", ".dsn", ".gbr"
        ])
        self.dev_filter_dropdown.currentIndexChanged.connect(self.apply_filter)
        self.layout.addWidget(self.dev_filter_dropdown)


        # Dropdown for directory selection
        self.directory_dropdown = QComboBox(self)
        self.directory_dropdown.setPlaceholderText("Select a directory")
        self.directory_dropdown.currentIndexChanged.connect(self.change_directory)
        self.layout.addWidget(self.directory_dropdown)

        # Add Directory Button
        self.add_directory_button = QPushButton("Add Directory", self)
        self.add_directory_button.clicked.connect(self.add_directory)
        self.layout.addWidget(self.add_directory_button)

        # Add Delete Directory Button
        self.delete_directory_button = QPushButton("Delete Directory", self)
        self.delete_directory_button.clicked.connect(self.delete_directory)
        self.layout.addWidget(self.delete_directory_button)


        # Header
        self.label_folder = QLabel("Monitoring Folder: None", self)
        self.label_folder.setStyleSheet("color: green;")
        self.layout.addWidget(self.label_folder)

        # Search bar
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("Real-Time-Search")
        self.search_bar.textChanged.connect(self.filter_files)
        self.layout.addWidget(self.search_bar)

        # Clear Search button
        self.clear_button = QPushButton("Clear Search", self)
        self.clear_button.clicked.connect(self.clear_search)
        self.clear_button.setStyleSheet("color: green;")
        self.layout.addWidget(self.clear_button)

        # Results list
        self.result_list = QListWidget(self)
        self.result_list.setContextMenuPolicy(Qt.CustomContextMenu)  # Enable custom context menu
        self.result_list.customContextMenuRequested.connect(self.show_context_menu)
        self.result_list.itemDoubleClicked.connect(self.open_file)
        self.layout.addWidget(self.result_list)

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar)

        # Refresh button
        self.refresh_button = QPushButton("Refresh Index", self)
        self.refresh_button.setStyleSheet("color: green;")
        self.refresh_button.clicked.connect(self.refresh_files)
        self.layout.addWidget(self.refresh_button)

        # Initialize the app
        self.load_or_index_files()
        self.start_monitoring()


    # Opens a selected file and attaches it in an emal automatically. OUTLOOK ONLY
    def send_email(self, item):
        """Sends the selected file as an email attachment using Outlook."""
        # Retrieve the full file path from the item's data
        selected_file = item.data(Qt.UserRole)
        
        if not selected_file or not os.path.exists(selected_file):
            QMessageBox.critical(self, "Error", f"File not found: {selected_file}")
            print(f"Error: File '{selected_file}' not found or no longer exists.")
            return

        try:
            # Initialize Outlook application
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # Create a new email item

            # Add details to the email
            mail.Subject = "File Attachment"
            mail.Body = "Please find the attached file."
            mail.Attachments.Add(selected_file)  # Attach the selected file

            # Display the email (for user to edit before sending)
            mail.Display()
            print(f"Email composed successfully with attachment: {selected_file}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to compose email: {e}")
            print(f"Error sending email: {e}")

    # Opens help menu
    def show_help(self):
        """Displays the Help dialog."""
        help_dialog = HelpDialog(self)
        help_dialog.exec_()

    # These are the mouse right click functions
    def show_context_menu(self, position):
        """Displays the context menu when right-clicking on a file."""
        item = self.result_list.itemAt(position)
        if item:
            context_menu = QMenu(self)
            
            save_as_action = context_menu.addAction("Save As...")
            save_as_action.triggered.connect(lambda: self.save_file_as(item))

            tag_action = context_menu.addAction("Add/Edit Tags")
            tag_action.triggered.connect(lambda: self.manage_tags(item))
            
            send_email_action = context_menu.addAction("Send as Email...")
            send_email_action.triggered.connect(lambda: self.send_email(item))

            context_menu.exec_(self.result_list.mapToGlobal(position))


    # Add or edit tags
    def manage_tags(self, item):
        """Opens a dialog to add or remove tags for a file and updates the results list."""
        # Retrieve the full file path from the item's data
        selected_file = item.data(Qt.UserRole)

        if not selected_file or not os.path.exists(selected_file):
            QMessageBox.critical(self, "Error", f"File not found: {selected_file}")
            print(f"Error: File '{selected_file}' not found or no longer exists.")
            return

        current_tags = self.tag_manager.get_tags(selected_file)

        # Input dialog for managing tags
        new_tags, ok = QInputDialog.getText(
            self,
            "Manage Tags",
            "Enter tags separated by commas:",
            text=", ".join(current_tags)
        )

        if ok:
            # Update tags in TagManager
            updated_tags = [tag.strip() for tag in new_tags.split(",") if tag.strip()]
            self.tag_manager.tags[selected_file] = updated_tags
            self.tag_manager.save_tags()

            # Refresh the results list to display the updated tags
            self.apply_filter()

            QMessageBox.information(self, "Tags Updated", f"Tags for {os.path.basename(selected_file)} updated.")


    # Saves a selected file to any location you want using a windows field dialog
    def save_file_as(self, item):
        """Opens a dialog for the user to save the selected file."""
        # Retrieve the full file path from the item's data
        selected_file = item.data(Qt.UserRole)

        if not selected_file or not os.path.exists(selected_file):
            QMessageBox.critical(self, "Error", f"File not found: {selected_file}")
            print(f"Error: File '{selected_file}' not found or no longer exists.")
            return

        # Open a "Save As" dialog
        target_path, _ = QFileDialog.getSaveFileName(self, "Save File As", os.path.basename(selected_file))
        if target_path:
            try:
                shutil.copy(selected_file, target_path)  # Copy the file to the new location
                QMessageBox.information(self, "File Saved", f"File saved to:\n{target_path}")
                print(f"File saved to: {target_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {e}")
                print(f"Error: Failed to save file '{selected_file}' with error: {e}")

    # Toggles light & dark mode
    def toggle_dark_mode(self):
        """Toggles between dark mode and light mode."""
        self.dark_mode_enabled = not self.dark_mode_enabled  # Toggle the mode
        if self.dark_mode_enabled:
            self.setStyleSheet(DARK_MODE_STYLESHEET)
        else:
            self.setStyleSheet(LIGHT_MODE_STYLESHEET)

    # User slected file type filtering
    def apply_filter(self):
        """Filters the displayed results based on the selected file type and updates the list."""
        file_type_filter = self.filter_dropdown.currentText()  # Existing file type filter
        dev_filter = self.dev_filter_dropdown.currentText() if hasattr(self, 'dev_filter_dropdown') else "Dev/Eng Files Filter"
        query = self.search_bar.text().strip().lower()
        self.result_list.clear()

        with self.files_lock:  # Safely access self.files
            files_snapshot = list(self.files)

        for file_path in files_snapshot:
            file_name = os.path.basename(file_path).lower()
            tags = ", ".join(self.tag_manager.get_tags(file_path))  # Get tags for the file
            display_text = f"{file_name} [Tags: {tags}]" if tags else file_name

            # Check if the file matches the filters
            matches_file_type = file_type_filter == "Common Files Filter" or file_name.endswith(file_type_filter)
            matches_dev_filter = dev_filter == "Dev/Eng Files Filter" or file_name.endswith(dev_filter)

            # Apply the search query and both filters
            if query in file_name and matches_file_type and matches_dev_filter:
                # Create the QListWidgetItem
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, file_path)  # Store the full file path as data

                # Highlight tags in light yellow if tags are present
                if tags:
                    item.setForeground(Qt.black)  # Default text color
                    item.setBackground(Qt.yellow)  # Highlight tags in light yellow

                # Add the item to the results list
                self.result_list.addItem(item)


    # Initially add your directory with system directory exclusions
    def add_directory(self):
        """Allows the user to add a directory and automatically index it."""
        # Define excluded directories
        excluded_directories = {
    "C:\\Windows",
    "C:\\Windows\\System32",
    "C:\\Windows\\SysWOW64",
    "C:\\Windows\\Temp",
    "C:\\Windows\\Hyper-V",
    "C:\\Program Files",
    "C:\\Program Files (x86)",
    "C:\\Recovery",
    "C:\\System Volume Information",
    "C:\\Perflogs",
    "C:\\Users",
    "C:\\Users\\AppData",
    "C:\\$Recycle.Bin",
    "C:\\Boot",
    "C:\\EFI",
    "Z:\\"
}


        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            if os.path.abspath(directory) in excluded_directories:
                QMessageBox.warning(
                    self,
                    "Invalid Directory",
                    f"The selected directory '{directory}' cannot be added as it is a protected system directory."
                )
                print(f"Attempted to add excluded directory: {directory}")
                return

            if directory in self.directories:
                QMessageBox.warning(self, "Duplicate Directory", "This directory is already added.")
                return

            if len(self.directories) >= 10:
                QMessageBox.warning(self, "Limit Reached", "You can only monitor up to 10 directories.")
                return

            # Add the new directory to the list and dropdown
            self.directories.append(directory)
            self.directory_dropdown.addItem(directory)
            print(f"Added directory: {directory}")

            # Automatically select and index the new directory
            self.directory_dropdown.setCurrentIndex(self.directory_dropdown.count() - 1)
            self.change_directory(self.directory_dropdown.currentIndex())

    # Changes the working directory
    def change_directory(self, index):
        """Switches to the selected directory and updates the results list."""
        if index >= 0 and index < len(self.directories):
            self.current_directory = self.directories[index]
            self.label_folder.setText(f"Monitoring Folder: {self.current_directory}")
            self.files.clear()  # Clear the files list
            self.result_list.clear()  # Clear the UI results list
            print(f"Switched to directory: {self.current_directory}")


            # Start indexing and update results after completion
            def on_indexing_complete():
                self.apply_filter()  # Apply filter to update results list
                self.signals.indexing_complete.disconnect(on_indexing_complete)

            self.signals.indexing_complete.connect(on_indexing_complete)
            self.index_files()  # Start reindexing the new directory
            self.start_monitoring()  # Restart monitoring for the new directory

    # Monitors the current working directoies for any changes and dynamically updates them
    def start_monitoring(self):
        """Starts or restarts monitoring the folder for changes."""
        if self.observer:
            self.observer.stop()
            self.observer.join()

        if self.current_directory:
            self.event_handler = FileMonitorHandler(
                self.handle_file_event, self.handle_file_removed, self.handle_file_renamed
            )
            self.observer = Observer()
            self.observer.schedule(self.event_handler, self.current_directory, recursive=True)
            self.observer.start()
            print(f"Monitoring started for directory: {self.current_directory}")

    # Clear the current search and results list
    def clear_search(self):
        """Clears the search bar and refreshes the results list based on the selected filter."""
        self.search_bar.clear()  # Clear the search query
        self.result_list.clear()  # Clear the results list in the UI
        self.apply_filter()  # Reapply the filter to reset the results list
        self.result_list.clear()  # Clear the results list in the UI

    def update_progress_bar(self, value):
        """Update the progress bar safely from the signal."""
        self.progress_bar.setValue(value)

    def on_indexing_complete(self):
        """Reset the progress bar when indexing is complete."""
        self.progress_bar.setValue(0) 

    def load_or_index_files(self):
        """Loads the file index from disk or indexes the folder if needed."""
        try:
            with self.lock.acquire(timeout=10):  # Wait up to 10 seconds to acquire the lock
                print("File lock acquired for reading index.")
                if os.path.exists(INDEX_FILE):
                    with open(INDEX_FILE, "r") as f:
                        saved_data = json.load(f)
                        self.directories = saved_data.get("directories", [])
                        self.files = set(saved_data.get("files", []))
                        self.last_modified_time = saved_data.get("last_modified_time", 0)

                        # Populate the directory dropdown with saved directories
                        self.directory_dropdown.clear()
                        for directory in self.directories:
                            self.directory_dropdown.addItem(directory)

                        # Set the first directory as the current directory (if available)
                        if self.directories:
                            self.current_directory = self.directories[0]
                            self.directory_dropdown.setCurrentIndex(0)
                            self.label_folder.setText(f"Monitoring Folder: {self.current_directory}")
                            print(f"Loaded directories: {self.directories}")
                            print(f"Loaded {len(self.files)} files.")
                        else:
                            print("No directories found in saved index.")
                else:
                    print("No saved index file found.")
        except Timeout:
            QMessageBox.critical(self, "Error", "Could not acquire file lock for index file.")
        finally:
            print("File lock released after reading index.")


    def delete_directory(self):
        """Deletes the selected directory from the list after user confirmation."""
        if not self.directories:
            QMessageBox.warning(self, "No Directory", "No directories to delete.")
            return

        current_index = self.directory_dropdown.currentIndex()
        if current_index < 0 or current_index >= len(self.directories):
            QMessageBox.warning(self, "Invalid Selection", "Please select a valid directory to delete.")
            return

        directory_to_remove = self.directories[current_index]
        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete the directory:\n{directory_to_remove}?\n"
            "This action will remove the directory from the index but will not delete its contents from the directory folder.",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Remove the directory from the list
            self.directories.pop(current_index)
            self.directory_dropdown.removeItem(current_index)

            # Clear the files if the current directory is deleted
            if self.current_directory == directory_to_remove:
                self.current_directory = None
                self.files.clear()
                self.result_list.clear()
                self.label_folder.setText("Monitoring Folder: None")

            # Save the updated state
            self.save_index()
            print(f"Deleted directory: {directory_to_remove}")
        else:
            print(f"Deletion canceled for directory: {directory_to_remove}")


    def save_index(self):
        """Saves the current index to disk."""
        try:
            with self.lock.acquire(timeout=10):  # Wait up to 10 seconds to acquire the lock
                print("Saving index to JSON.")
                with open(INDEX_FILE, "w") as f:
                    json.dump(
                        {
                            "directories": self.directories,
                            "last_modified_time": self.last_modified_time,
                            "files": list(self.files),
                        },
                        f,
                        indent=4  # Pretty print for easier debugging
                    )
                print(f"Index saved: {INDEX_FILE}")
        except Timeout:
            QMessageBox.critical(self, "Error", "Could not acquire file lock for saving index file.")
        finally:
            print("File lock released after saving index.")


    def index_files(self):
        """Indexes all files in the folder with real-time progress updates, excluding specific directories and file types."""
        if not self.current_directory:
            QMessageBox.warning(self, "No Directory", "Please select or add a directory to monitor.")
            return

        # Set the progress bar to 0% when indexing starts
        self.progress_bar.setValue(0)
        self.statusBar().setStyleSheet("color: red;")
        self.statusBar().showMessage("Indexing Started, Please Wait...")

        def scan_folder():
            print(f"Starting indexing for directory: {self.current_directory}")
            self.files.clear()
            files = []

            # Define excluded directories and file types
            excluded_directories = {"C:\\Windows", "C:\\Program Files", "C:\\Program Files (x86)", "Z:\\"}
            excluded_file_types = {".ini", ".tmp", ".bak", ".log", ".sys", ".dll", ".reg", ".cab", ".msi", ".drv", ".inf", ".db", ".ink", ".exe", ".scr"}

            # Scan the directory and collect all files
            for root, _, file_list in os.walk(self.current_directory):
                # Skip excluded directories
                if any(os.path.abspath(root).startswith(excluded) for excluded in excluded_directories):
                    print(f"Skipping excluded directory: {root}")
                    continue

                for file in file_list:
                    full_path = os.path.join(root, file)
                    _, file_ext = os.path.splitext(full_path)

                    # Skip excluded file types
                    if file_ext.lower() in excluded_file_types:
                        print(f"Skipping excluded file: {full_path}")
                        continue

                    files.append(full_path)

            total_files = len(files)
            if total_files == 0:
                print("No valid files found in the directory.")
                self.signals.progress.emit(100)  # Complete the progress bar
                self.signals.indexing_complete.emit()
                return

            # Emit progress bar updates as files are indexed
            current_progress = 0
            for i, file_path in enumerate(files, start=1):
                self.files.add(file_path)

                # Update progress incrementally
                new_progress = int((i / total_files) * 100)
                if new_progress > current_progress:  # Emit only if progress has increased
                    current_progress = new_progress
                    self.signals.progress.emit(current_progress)
                    self.statusBar().showMessage(f"Indexing progress: {current_progress}%")

                print(f"Indexed: {file_path}")

            # Save the index and signal completion
            self.save_index()
            self.signals.progress.emit(100)  # Ensure the progress bar reaches 100%
            self.signals.indexing_complete.emit()
            print(f"Indexing complete. Total files indexed: {len(self.files)}")
            self.statusBar().setStyleSheet("color: green;")
            self.statusBar().showMessage(f"Indexing Completed. Total files: {len(self.files)}")

        # Start the background thread
        self.indexing_thread = threading.Thread(target=scan_folder, daemon=True)
        self.indexing_thread.start()



    def is_safe_path(self, base_path, target_path):
        """Ensure the target path is within the base path."""
        base_path = os.path.abspath(base_path)
        target_path = os.path.abspath(target_path)
        return os.path.commonpath([base_path]) == os.path.commonpath([base_path, target_path])


    def handle_file_event(self, file_path):
        """Handles new or modified files detected by watchdog."""
        if not self.is_safe_path(self.current_directory, file_path):
            print(f"Blocked access to unsafe file path: {file_path}")
            return

        # Add the new or modified file to the index
        if file_path not in self.files:
            self.files.add(file_path)
            self.save_index()  # Save the updated index
            print(f"File added: {file_path}")


    def handle_file_removed(self, file_path):
        """Handles file removals detected by watchdog."""
        if self.is_safe_path(self.current_directory, file_path):
            if file_path in self.files:
                self.files.discard(file_path)  # Remove the file from the index
                self.save_index()  # Save the updated index
                print(f"File removed: {file_path}")



    def handle_file_renamed(self, old_path, new_path):
        """Handles file renames or moves detected by watchdog."""
        if not self.is_safe_path(self.current_directory, old_path) or not self.is_safe_path(self.current_directory, new_path):
            print(f"Blocked unsafe rename/move: {old_path} -> {new_path}")
            return

        # Remove the old file and add the new one
        if old_path in self.files:
            self.files.remove(old_path)
            self.result_list.clear()  # Clear the UI list to refresh

        if new_path not in self.files:
            self.files.add(new_path)
            self.save_index()  # Save the updated index
            print(f"File renamed: {old_path} -> {new_path}")


    def filter_files(self):
        """Filters the files based on the search query and selected file types."""
        query = self.search_bar.text().strip().lower()
        file_type_filter = self.filter_dropdown.currentText()  # Get the file type filter
        dev_filter = self.dev_filter_dropdown.currentText() if hasattr(self, 'dev_filter_dropdown') else "Dev Files Filter"
        self.result_list.clear()

        with self.files_lock:  # Safely access self.files
            files_snapshot = list(self.files)

        for file_path in files_snapshot:
            file_name = os.path.basename(file_path).lower()
            tags = ", ".join(self.tag_manager.get_tags(file_path))  # Get tags for the file
            display_text = f"{file_name} [Tags: {tags}]" if tags else file_name

            # Check if the file matches both filters
            matches_file_type = file_type_filter == "Common Files Filter" or file_name.endswith(file_type_filter)
            matches_dev_filter = dev_filter == "Dev/Eng Files Filter" or file_name.endswith(dev_filter)

            # Apply the search query and both filters
            if query in file_name and matches_file_type and matches_dev_filter:
                # Create the QListWidgetItem with the display text
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, file_path)  # Store the full file path as data

                # Highlight tags in light yellow if tags are present
                if tags:
                    item.setForeground(Qt.black)  # Default text color
                    item.setBackground(Qt.yellow)  # Highlight tags in light yellow

                self.result_list.addItem(item)



    def refresh_files(self):
        """Refreshes the file list for the current directory and includes tags in the display."""
        if not self.current_directory:
            QMessageBox.warning(self, "No Directory", "Please select or add a directory to monitor.")
            return

        self.result_list.clear()  # Clear the results list in the UI

        # Reindex the files
        self.index_files()

        # Apply the selected filter directly after reindexing
        self.apply_filter()

        # Use a snapshot of files to avoid threading issues
        files_snapshot = list(self.files)  # Create a list copy of the set

        for file_path in files_snapshot:
            file_name = os.path.basename(file_path)
            tags = ", ".join(self.tag_manager.get_tags(file_path))  # Get tags for the file
            display_text = f"{file_name} [Tags: {tags}]" if tags else file_name
            self.result_list.addItem(display_text)

        print(f"File list refreshed for directory: {self.current_directory}")


    def open_file(self, item):
        """Opens a file with the default application."""
        # Retrieve the full file path from the item's data
        selected_file = item.data(Qt.UserRole)  # Fetch the stored file path

        if not selected_file or not os.path.exists(selected_file):
            QMessageBox.critical(self, "Error", f"File not found: {selected_file}")
            print(f"Error: File '{selected_file}' not found or no longer exists.")
            return

        try:
            # Use os.startfile to open the file with the default application
            os.startfile(selected_file)
            print(f"Opened file: {selected_file}")
        except FileNotFoundError:
            QMessageBox.critical(self, "Error", f"File no longer exists: {selected_file}")
            print(f"Error: File '{selected_file}' does not exist.")
        except OSError as e:
            QMessageBox.critical(self, "Error", f"Failed to open file: {e}")
            print(f"Error: Failed to open file '{selected_file}' with error: {e}")



    def closeEvent(self, event):
        """Handles application close event."""
        # Check if indexing is in progress
        if hasattr(self, "indexing_thread") and self.indexing_thread.is_alive():
            reply = QMessageBox.question(
                self,
                "Indexing in Progress",
                "Indexing is currently in progress. Do you want to cancel indexing and close the application?",
                QMessageBox.Yes | QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                # Attempt to stop the indexing thread
                try:
                    self.indexing_thread.join(timeout=2)  # Allow the thread to terminate gracefully
                    if self.indexing_thread.is_alive():
                        QMessageBox.warning(self, "Indexing Interrupted", "Indexing thread could not be stopped.")
                    else:
                        print("Indexing thread stopped.")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to stop indexing thread: {e}")
                    print(f"Error stopping indexing thread: {e}")
            else:
                event.ignore()  # Cancel the close event
                return

        # Stop monitoring
        if hasattr(self, "observer") and self.observer is not None:
            try:
                self.observer.stop()
                self.observer.join()
                print("Observer stopped.")
            except Exception as e:
                print(f"Error stopping observer: {e}")

        # Final cleanup
        print("Application closing.")
        event.accept()  # Accept the close event



if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    window = FileSearcherApp()
    window.show()
    sys.exit(app.exec_())


