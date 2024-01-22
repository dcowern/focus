import sys
from PyQt6.QtWidgets import QApplication, QMenu, QSystemTrayIcon, QVBoxLayout, QDialog, QSlider, QPushButton, QColorDialog, QLabel
from PyQt6.QtGui import QIcon, QColor
from PyQt6.QtCore import Qt, pyqtSignal
import win32gui
import win32api
import win32con
import ctypes
import ctypes.wintypes
import os
import json
import sys

WinEventProcType = ctypes.WINFUNCTYPE(
    None, 
    ctypes.wintypes.HANDLE,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND,
    ctypes.wintypes.LONG,
    ctypes.wintypes.LONG,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD
)

class ConfigDialog(QDialog):
    """
    A dialog for configuring the settings of the Focus application.
    """

    config_changed = pyqtSignal()
    preview_changed = pyqtSignal()

    def __init__(self, tint, transparency, parent=None):
        """
        Initialize the ConfigDialog.

        Args:
            tint (int): The tint color.  Note that this is not currently impemented because SetLayeredWindowAttributes
                        does not work the way I thought it did.  I thought it would tint the window, but it actually
                        just sets the color key for the window.  I'm leaving this in here in case I want to implement
                        it later.
            transparency (float): The transparency value.
            parent (QWidget): The parent widget.
        """
        super().__init__(parent)

        self.transparency_config = transparency
        self.tint_config = tint

        # Create the dialog without a close button
        self.setWindowTitle("Focus Configuration")
        #self.resize(400, 300)
        self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, False)

        # Create the layout
        self.layout = QVBoxLayout()

        # Create the transparency slider label
        self.transparency_slider_label = QLabel("Transparency")

        # Create the transparency slider
        self.transparency_slider = QSlider(Qt.Orientation.Horizontal)
        self.transparency_slider.setMinimum(0)
        self.transparency_slider.setMaximum(255)
        self.transparency_slider.setValue(int(self.transparency_config))

        # Create the tint color picker label
        self.tint_color_picker_label = QLabel("Tint Color")

        # Create the tint color picker without OK and cancel buttons
        #self.tint_color_picker = QColorDialog()
        #self.tint_color_picker.setOption(QColorDialog.ColorDialogOption.NoButtons)
        #self.tint_color_picker.setCurrentColor(QColor(self.tint_config))
        
        # Create the OK button
        self.ok_button = QPushButton("OK")

        # Create the cancel button
        self.cancel_button = QPushButton("Cancel")

        # Add the widgets to the layout
        self.layout.addWidget(self.transparency_slider_label)
        self.layout.addWidget(self.transparency_slider)
        #self.layout.addWidget(self.tint_color_picker_label)
        #self.layout.addWidget(self.tint_color_picker)
        self.layout.addWidget(self.ok_button)
        self.layout.addWidget(self.cancel_button)

        # Add the layout to the dialog
        self.setLayout(self.layout)

        # Connect the OK and cancel buttons to their functions
        self.ok_button.clicked.connect(self.ok_button_clicked)
        self.cancel_button.clicked.connect(self.cancel_button_clicked)

        # Connect the transparency slider to the preview changed signal
        self.transparency_slider.valueChanged.connect(self.preview_changed)

        # Connect the tint color picker to the preview changed signal
        #self.tint_color_picker.currentColorChanged.connect(self.preview_changed)
    
    def ok_button_clicked(self):
        """
        Function called when the OK button is clicked.

        Saves the configuration to memory, gets the transparency value,
        gets the tint value, hides the dialog, and emits the config_changed signal.

        Args:
            None

        Returns:
            None
        """
        print("Saving config to memory")

        # Get the transparency value
        self.transparency_config = self.transparency_slider.value()

        # Get the tint value as an RGB value without the alpha channel
        #self.tint_config = self.tint_color_picker.currentColor().rgb() & 0x00FFFFFF

        # Hide the dialog
        self.hide()

        # Emit the config changed signal to let the main app know the config has changed
        self.config_changed.emit()

    def cancel_button_clicked(self):
        """
        Handle the click event of the Cancel button.

        This method is called when the Cancel button is clicked. It prints a message
        indicating that the configuration changes are being canceled and hides the dialog.

        Args:
            None

        Returns:
            None
        """
        print("Canceling config changes")

        # Hide the dialog
        self.hide()

    def set_transparency(self, transparency):
        """
        Set the transparency value.

        Args:
            transparency (float): The transparency value.  (Range: 0 to 255)
        """
        self.transparency_config = transparency

    def set_tint(self, tint):
        """
        Set the tint color. Note that window tinting is not currently implemented.

        Args:
            tint (int): The tint color. As an RGB value without the alpha channel. (Range: 0x000000 to 0xFFFFFF)

        Returns:
            None
        """
        self.tint_config = tint

    def get_transparency(self):
        """
        Get the saved transparency value.

        Args:
            None

        Returns:
            float: The transparency value.
        """
        return self.transparency_config
    
    def get_preview_transparency(self):
        """
        Get the transparency value directly from the transparency slider.

        Args:
            None

        Returns:
            float: The preview transparency value from the transparency slider. (Range: 0 to 255)
        """
        return self.transparency_slider.value()
    
    def get_tint(self):
        """
        Get the saved tint color. Note that window tinting is not currently implemented.

        Args:
            None

        Returns:
            int: The tint color. As an RGB value without the alpha channel. (Range: 0x000000 to 0xFFFFFF)
        """
        return self.tint_config
    
    def get_preview_tint(self):
        """
        Get the tint color directly from the color picker. Note that window tinting is not currently implemented.

        Args:
            None

        Returns:
            int: The tint color from the color picker. As an RGB value without the alpha channel. (Range: 0x000000 to 0xFFFFFF)
        """
        # Return the tint value as an RGB value without the alpha channel
        # return self.tint_color_picker.currentColor().rgb() & 0x00FFFFFF
        return self.tint_config

    def showEvent(self, event):
        """
        Override the showEvent function to update the slider and color picker values every time the dialog is shown.

        Args:
            event (QShowEvent): The show event.

        Returns:
            None
        """
        self.transparency_slider.setValue(int(self.transparency_config))
        #self.tint_color_picker.setCurrentColor(QColor(self.tint_config))

    def closeEvent(self, event):
        """
        Override the closeEvent function to hide the dialog instead of closing it so that the dialog maintains state between uses.

        Args:
            event (QCloseEvent): The close event.

        Returns:
            None
        """
        event.ignore()
        self.hide()

class FocusApp(QApplication):
    """
    The main application class for the Focus application.
    """

    def __init__(self, *args):
        """
        Initialize the FocusApp.

        Args:
            args: The arguments passed to the application.

        Returns:
            None
        """
        super().__init__(*args)

        # Set default transparency values
        self.transparency_max = 255
        self.transparency_dim = 0.5 * self.transparency_max
        self.transparency_default = 1.0 * self.transparency_max

        # Set default tint values
        self.tint_color = 0x00000080
        self.tint_default = 0x00000000

        # Set the config file path to the users home directory
        self.config_file_path = os.path.join(os.path.expanduser("~"), ".focus_config.json")

        # If the config file exists, read it. Otherwise, create it.
        if os.path.exists(self.config_file_path):
            print("Config file exists - reading config file")
            self.read_config()
        else:
            print("Config file does not exist - creating config file")
            self.save_config()
        
        # Create the system tray icon
        # Icon source: https://icons8.com/icon/50274/aperture
        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self)

        # Set the app to dim windows by default
        self.bDim = True

        # Create the menu
        self.menu = QMenu()

        # Create the dim and undim menu options
        self.dim_option = self.menu.addAction("Dim")
        self.undim_option = self.menu.addAction("Undim")
        
        # Make the dim and undim options checkable
        self.dim_option.setCheckable(True)
        self.undim_option.setCheckable(True)

        # Put a check next to the selected dim/undim option
        if self.bDim:
            self.dim_option.setChecked(True)
            self.undim_option.setChecked(False)
        else: 
            self.dim_option.setChecked(False)
            self.undim_option.setChecked(True)

        # Add a divider line
        self.menu.addSeparator()

        # Create the configure option
        self.config_option = self.menu.addAction("Configure")

        # Add a divider line
        self.menu.addSeparator()

        # Create the exit option
        self.exit_option = self.menu.addAction("Exit")

        # Connect the actions to their respective functions
        self.dim_option.triggered.connect(self.dim_action)
        self.undim_option.triggered.connect(self.undim_action)
        self.config_option.triggered.connect(self.config_action)
        self.exit_option.triggered.connect(self.exit_action)

        # Add the actions to the menu
        self.menu.addAction(self.dim_option)
        self.menu.addAction(self.undim_option)
        self.menu.addAction(self.config_option)
        self.menu.addAction(self.exit_option)

        # Set the menu for the system tray icon
        self.tray_icon.setContextMenu(self.menu)

        # Create the config dialog and hide it
        self.config_dialog = ConfigDialog(self.tint_color, self.transparency_dim)
        self.config_dialog.hide()

        # Connect the config_changed signal to the update_config slot
        self.config_dialog.config_changed.connect(self.update_config)

        # Connect the preview_changed signal to the preview_changes slot
        self.config_dialog.preview_changed.connect(self.preview_changes)

        # Show the system tray icon
        self.tray_icon.show()

        # Use WinEventProcType to create a callback function that receives notifications
        self.WinEventProc = WinEventProcType(self.active_window_change_callback)

        # Use user32.SetWinEventHook to hook to the active window change callback
        ctypes.windll.user32.SetWinEventHook(win32con.EVENT_OBJECT_FOCUS, win32con.EVENT_OBJECT_FOCUS, None, self.WinEventProc, 0, 0, win32con.WINEVENT_OUTOFCONTEXT | win32con.WINEVENT_SKIPOWNPROCESS)

    def dim_action(self):
        """
        This method is called when the "Dim" option is selected.

        It sets the class variable bDim to True, checks the dim_option checkbox,
        unchecks the undim_option checkbox, and dims all windows except the active window.

        Args:
            None

        Returns:
            None
        """
        print("Dim option selected")

        # Set the class variable bDim to true
        self.bDim = True
        
        # Set the Menu Checkboxes
        self.dim_option.setChecked(True)
        self.undim_option.setChecked(False)

        # Dim all windows except the active window
        self.dim_inactive_windows()

    def dim_inactive_windows(self):
        """
        Dim the inactive windows on the screen.

        This method gets the handle of the active window and then dims all the other visible windows
        except for the taskbar, start menu, and the active window itself. It sets the WS_EX_LAYERED
        extended style for each window and applies transparency and tint color to achieve the dimming effect.

        Args:
            None

        Returns:
            None
        """

        # Get the handle of the active window
        active_window = win32gui.GetForegroundWindow()

        print("Dimming inactive windows")

        if self.bDim:
            self.undim_active_window()

            # Enumerate all top-level windows
            def enum_dim_callback(hwnd, _):
                # Skip the active window
                if hwnd != active_window:

                    # Make sure hwnd is not null
                    if hwnd is not None:
                        
                        # Make sure the window is visible
                        if win32gui.IsWindowVisible(hwnd) != 0 and win32gui.IsIconic(hwnd) == 0:

                            # Make sure the window is not the taskbar or the start menu
                            if win32gui.GetClassName(hwnd) != "Shell_TrayWnd" and win32gui.GetClassName(hwnd) != "Button" and win32gui.GetClassName(hwnd) != "Windows.UI.Core.CoreWindow":

                                # Check whether WS_EX_LAYERED is set
                                if win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_LAYERED == 0:
                                    # Use ctypes.windll.user32.SetWindowLongPtrW to make sure WS_EX_LAYERED is set
                                    ctypes.windll.user32.SetWindowLongPtrW(hwnd, win32con.GWL_EXSTYLE,
                                                                    ctypes.windll.user32.GetWindowLongPtrW(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
                                
                                # Check whether WS_EX_LAYERED is set
                                if win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_LAYERED != 0:
                                    
                                    #print("Dimming window (hwnd: " + str(hwnd) + ")")
                                    #print("Transparency: " + str(self.transparency_dim) + ", Tint: " + str(hex(self.tint_color)))
                                    flags = win32con.LWA_COLORKEY | win32con.LWA_ALPHA

                                    if win32gui.SetLayeredWindowAttributes(hwnd, self.tint_color, int(self.transparency_dim), flags) == 0:
                                        print("Error setting layered window attributes (hwnd: " + str(hwnd) + ")")
                                        print("Error: " + str(win32api.GetLastError()))
                                    win32gui.RedrawWindow(hwnd, None, None, win32con.RDW_ERASE | win32con.RDW_INVALIDATE | win32con.RDW_FRAME | win32con.RDW_ALLCHILDREN)

            win32gui.EnumWindows(enum_dim_callback, None)

    def undim_action(self):
        """
        Undim the windows and update the menu checkboxes.

        This method sets the `bDim` flag to False, indicating that the windows should not be dimmed.
        It also updates the menu checkboxes to reflect the undim option being selected.
        Finally, it calls the `undim_all_windows` method to undim all windows.

        Args:
            None

        Returns:
            None
        """
        print("Undim option selected")
        
        # Set the Dim flag to false
        self.bDim = False

        # Set the Menu Checkboxes
        self.dim_option.setChecked(False)
        self.undim_option.setChecked(True)

        # Undim all windows
        self.undim_all_windows()

    def config_action(self):
        """
        This method is called when the config option is selected.
        It prints a message, reads the config file, sets the config dialog values,
        and shows the config dialog.

        Args:
            None

        Returns:
            None
        """
        print("Config option selected")

        # Read the config file
        self.read_config()

        # Set the config dialog values
        self.config_dialog.set_tint(self.tint_color)
        self.config_dialog.set_transparency(self.transparency_dim)

        # Show the config dialog
        self.config_dialog.show()

    def update_config(self):
        """
        Update the configuration settings.

        This method retrieves the values from the config dialog, saves the configuration,
        and applies the new configuration settings if the dim flag is set.

        Args:
            None

        Returns:
            None
        """
        print("Updating config")

        # Get the config dialog values
        self.tint_color = self.config_dialog.get_tint()
        self.transparency_dim = self.config_dialog.get_transparency()

        # Save the config
        self.save_config()

        # If the dim flag is set, re-dim all inactive windows using the new config values
        if self.bDim:
            self.dim_inactive_windows()

    def preview_changes(self):
        """
        Preview the changes made in the config dialog in real time.

        This method retrieves the tint color and transparency dim values from the config dialog.
        If the dim flag is set, it re-dims all inactive windows using the new config values.

        Args:
            None

        Returns:
            None
        """
        print("Previewing changes")

        # Get the config dialog values
        self.tint_color = self.config_dialog.get_preview_tint()
        self.transparency_dim = self.config_dialog.get_preview_transparency()

        # If the dim flag is set, re-dim all inactive windows using the new config values
        if self.bDim:
            self.dim_inactive_windows()

    def read_config(self):
        """
        Read the saved values from the config file.

        This method reads the values from the config file, if it exists, and assigns them to the corresponding attributes of the class.
        If the config file does not exist, default values are used instead.

        Args:
            None

        Returns:
            None
        """
        print("Reading config")

        # Read the config file if it exists
        try:
            config_file = open(self.config_file_path, "r")
            config_data = json.load(config_file)

            # Read the transparency value
            try:
                self.transparency_dim = config_data["transparency"]
            except:
                print("Error reading transparency value - keeping default value")

            # Read the tint value
            try:
                self.tint_color = config_data["tint"]
            except:
                print("Error reading tint value - keeping default value")

            # Close the config file
            config_file.close()

        # Don't read anything - just use the default values if the config file doesn't exist.
        except:
            print("====================================")
            print("Error reading config file at path: " + self.config_file_path)
            print("Error: " + str(sys.exc_info()[0]))
            print("Using default values instead.")
            print("====================================")


    def save_config(self):
        """
        Saves the configuration settings to a file.

        This function deletes the existing config file (if it exists),
        creates a new config file, and writes the current configuration
        settings to the file.

        Raises:
            OSError: If there is an error deleting the existing config file
            or creating/writing to the new config file.

        Args:
            self (object): The instance of the class.

        Returns:
            None
        """
        
        print("Saving config")

        try:
            # Delete the config file if it exists
            if os.path.exists(self.config_file_path):
                os.remove(self.config_file_path)

            # Create the config file
            config_file = open(self.config_file_path, "w")
            config_file.write("{\n")
            config_file.write("\t\"transparency\": " + str(self.transparency_dim) + ",\n")
            config_file.write("\t\"tint\": " + str(self.tint_color) + "\n")
            config_file.write("}")
            config_file.close()
        except:
            print("====================================")
            print("Error saving config file to path: " + self.config_file_path)
            print("Error: " + str(sys.exc_info()[0]))
            print("====================================")
 
    def undim_all_windows(self):
        """
        Undims all windows by setting the WS_EX_LAYERED extended window style and adjusting the transparency.

        This method enumerates all top-level windows and checks if they are visible and not minimized.
        If a window does not have the WS_EX_LAYERED style set, it sets the style using SetWindowLongPtrW.
        Then, it adjusts the transparency of the window using SetLayeredWindowAttributes and redraws the window.

        Args:
            self (object): The instance of the class.

        Returns:
            None
        """
        print("Undimming all windows")

        if not self.bDim:

            # Enumerate all top-level windows
            def enum_undim_callback(hwnd, _):

                # Make sure hwnd is not null
                if hwnd is not None:

                    # Make sure the window is visible
                    if win32gui.IsWindowVisible(hwnd) != 0 and win32gui.IsIconic(hwnd) == 0:

                        # Check whether WS_EX_LAYERED is set
                        if win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_LAYERED == 0:

                            print("Setting WS_EX_LAYERED (hwnd: " + str(hwnd) + ")" )

                            # Use ctypes.windll.user32.SetWindowLongPtrW to make sure WS_EX_LAYERED is set
                            ctypes.windll.user32.SetWindowLongPtrW(hwnd, win32con.GWL_EXSTYLE,
                                                                ctypes.windll.user32.GetWindowLongPtrW(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
                            
                        # Verify WS_EX_LAYERED is set
                        if win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_LAYERED != 0:
                            
                            #print("Undimming window (hwnd: " + str(hwnd) + ")")
                            
                            win32gui.SetLayeredWindowAttributes(hwnd, self.tint_color, int(self.transparency_default), win32con.LWA_ALPHA | win32con.LWA_COLORKEY)
                            win32gui.RedrawWindow(hwnd, None, None, win32con.RDW_ERASE | win32con.RDW_INVALIDATE | win32con.RDW_FRAME | win32con.RDW_ALLCHILDREN)

            win32gui.EnumWindows(enum_undim_callback, None)

    def undim_active_window(self):
        """
        Undims the active window by setting the transparency level to the default value.

        This function gets the handle of the active window, sets the WS_EX_LAYERED style using SetWindowLongPtrW,
        and then sets the transparency level using SetLayeredWindowAttributes.

        Args:
            None

        Returns:
            None
        """
        # Get the hwnd of the active window
        hwnd = win32gui.GetForegroundWindow()

        # Use ctypes.windll.user32.SetWindowLongPtrW to make sure WS_EX_LAYERED is set
        ctypes.windll.user32.SetWindowLongPtrW(hwnd, win32con.GWL_EXSTYLE,
                        ctypes.windll.user32.GetWindowLongPtrW(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)

        win32gui.SetLayeredWindowAttributes(hwnd, self.tint_default, int(self.transparency_default), win32con.LWA_ALPHA)

    def active_window_change_callback(self, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
        """
        Callback function triggered when the active window changes.

        Args:
            hWinEventHook (int): The handle to the event hook.
            event (int): The event that occurred.
            hwnd (int): The handle to the window that triggered the event.
            idObject (int): The identifier of the object associated with the event.
            idChild (int): The identifier of the child object associated with the event.
            dwEventThread (int): The identifier of the thread that triggered the event.
            dwmsEventTime (int): The timestamp of the event.

        Returns:
            None
        """

        #print("Active window changed (hwnd: " + str(hwnd) + ")")

        # Dim all windows except the active window
        if self.bDim:
            self.dim_inactive_windows()


    def exit_action(self):
        """
        Perform the exit action.

        This method is called when the exit option is selected. It undims all windows to clean up
        and then exits the program.

        Args:
            None

        Returns:
            None
        """
        print("Exit option selected")
        # Undim all Windows to clean up
        self.undim_action()
        sys.exit()

if __name__ == "__main__":
    app = FocusApp(sys.argv)
    sys.exit(app.exec())
