import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

# Import the main selection window class from your existing file
from selection_window import SelectionWindow


import sys
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class SplashScreen:
    def __init__(self, root, on_close):
        """
        Initialize the splash screen.
        
        Args:
            root (tk.Tk): The main Tkinter root window.
            on_close (function): The function to call when the splash screen timer finishes.
        """
        self.root = root
        self.on_close = on_close
        
        # Hide the main window until the splash screen is done
        self.root.withdraw()
        
        # Create a Toplevel window for the splash screen
        self.splash_window = tk.Toplevel(self.root)
        self.splash_window.overrideredirect(True) # Make it borderless

        # --- Load and Display Logo ---
        try:
            # Use Pillow to open the image and create a PhotoImage
            logo_image = Image.open(resource_path("OrbitandSkyline.jpg"))
            # Resize the image if necessary (e.g., to a width of 300px)
            logo_image = logo_image.resize((300, 150), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_image)
            
            # Use a Label to display the image
            logo_label = ttk.Label(self.splash_window, image=self.logo, anchor="center")
            logo_label.pack(pady=20, padx=20)

        except FileNotFoundError:
            # Fallback if the image is not found
            logo_label = ttk.Label(self.splash_window, text="Orbit and Skyline", font=("Helvetica", 24, "bold"), anchor="center")
            logo_label.pack(pady=(40, 20), padx=40)
        
        # --- Add a Loading Message ---
        loading_label = ttk.Label(self.splash_window, text="Loading Match Network Analyzer...", font=("Helvetica", 12))
        loading_label.pack(pady=(0, 20))

        # --- Center the Splash Screen on the Screen ---
        self.splash_window.update_idletasks() # Update window info
        screen_width = self.splash_window.winfo_screenwidth()
        screen_height = self.splash_window.winfo_screenheight()
        splash_width = self.splash_window.winfo_width()
        splash_height = self.splash_window.winfo_height()

        x = (screen_width // 2) - (splash_width // 2)
        y = (screen_height // 2) - (splash_height // 2)

        self.splash_window.geometry(f'{splash_width}x{splash_height}+{x}+{y}')

        # --- Set a Timer to Close the Splash Screen ---
        # Close after 3000 milliseconds (3 seconds)
        self.splash_window.after(3000, self.close_splash)

    def close_splash(self):
        """Destroys the splash screen and calls the callback function."""
        self.splash_window.destroy()
        self.on_close()

def launch_main_application(root):
    """Initializes the main application window."""
    # The SelectionWindow now takes control of the root window
    app = SelectionWindow(root)
    # Make the main window visible
    root.deiconify()

if __name__ == "__main__":
    # 1. Create the main root window for the entire application
    root = tk.Tk()
    
    # --- MODIFICATION START ---
    # Set the window icon for the entire application.
    # This removes the default feather icon. This line affects the root window
    # and all subsequent Toplevel windows.
    try:
        # NOTE: You must provide a '.ico' file for this to work.
        root.iconbitmap(resource_path("logo.ico"))

    except tk.TclError:
        print("Warning: 'logo.ico' not found. The application will use the default icon.")
    # --- MODIFICATION END ---
    
    # 2. Define the function that will launch the main app
    #    We pass it 'root' so it knows which window to build on
    main_app_launcher = lambda: launch_main_application(root)

    # 3. Create the splash screen, passing it the root window and the launcher function
    #    The splash screen will call 'main_app_launcher' when it's done.
    splash = SplashScreen(root, on_close=main_app_launcher)

    # 4. Start the Tkinter event loop. This will show the splash screen first.
    root.mainloop()