import os
import time
from tkinter import filedialog
from PIL import ImageGrab, Image, ImageTk
from fpdf import FPDF
import customtkinter as ctk
from pathlib import Path
from pynput import keyboard
from fpdf.enums import XPos, YPos
from customtkinter import CTkImage


class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot Documentation Tool")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 400
        window_height = 300
        position_x = int(((screen_width // 2) - (window_width // 2)) * 2)
        position_y = int(((screen_height // 2) - (window_height // 2)) * 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        # Initialize variables
        self.is_capturing = False
        self.screenshots = []
        self.selected_screenshots = []
        self.listener = None
        self.current_keys = set()  # Track currently pressed keys

        # Configure CustomTkinter appearance
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # UI Elements
        self.start_button = ctk.CTkButton(root, text="Start Capturing", command=self.start_capturing)
        self.start_button.pack(pady=20)

        self.stop_button = ctk.CTkButton(root, text="Stop Capturing", command=self.stop_capturing, state="disabled")
        self.stop_button.pack(pady=20)

        self.preview_button = ctk.CTkButton(root, text="Preview Screenshots", command=self.preview_screenshots, state="disabled")
        self.preview_button.pack(pady=20)

        self.export_button = ctk.CTkButton(root, text="Export to PDF", command=self.export_to_pdf, state="disabled")
        self.export_button.pack(pady=20)

        self.status_label = ctk.CTkLabel(root, text="Status: Idle", text_color="gray", font=("Helvetica", 13), anchor="w")
        self.status_label.pack(side="bottom", fill="x", padx=10, pady=6)

    def start_capturing(self):
        self.is_capturing = True
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.preview_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self.status_label.configure(text="Status: Capturing...")
        self.start_keyboard_listener()

    def stop_capturing(self):
        self.is_capturing = False
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.preview_button.configure(state="normal")
        self.export_button.configure(state="normal")
        self.status_label.configure(text="Status: Stopped")
        self.stop_keyboard_listener()

    def start_keyboard_listener(self):
        # Directory to save screenshots
        self.screenshots_dir = Path.home() / "Documents" / "Screenshots"
        if not self.screenshots_dir.exists():
            os.makedirs(self.screenshots_dir)

        # Keyboard key press event handlers
        def on_press(key):
            if self.is_capturing:
                try:
                    # Add the key to the set of currently pressed keys
                    self.current_keys.add(key)

                    # Check if both CTRL and SHIFT are pressed
                    if keyboard.Key.ctrl_l in self.current_keys or keyboard.Key.ctrl_r in self.current_keys:
                        if keyboard.Key.shift in self.current_keys:
                            # Take a screenshot
                            timestamp = time.strftime("%Y%m%d_%H%M%S")
                            screenshot_path = self.screenshots_dir / f"step_{timestamp}.png"
                            screenshot = ImageGrab.grab()
                            screenshot.save(screenshot_path)
                            self.screenshots.append(screenshot_path)
                            self.status_label.configure(text=f"Screenshot taken: {screenshot_path.name}", text_color="green")

                except Exception as e:
                    self.status_label.configure(text=f"Error: {str(e)}", text_color="red")

        def on_release(key):
            # Remove the key from the set of currently pressed keys
            if key in self.current_keys:
                self.current_keys.remove(key)

        # Start the listener in a non-blocking way
        self.listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.listener.start()

    def stop_keyboard_listener(self):
        if self.listener:
            self.listener.stop()
            self.listener = None

    def preview_screenshots(self):
        if not self.screenshots:
            self.status_label.configure(text="No screenshots to preview", text_color="orange")
            return

        preview_window = ctk.CTkToplevel(self.root)
        preview_window.transient(self.root)    # Keep it on top of main window
        preview_window.lift()                  # Bring it to front
        preview_window.focus_force()          # Grab keyboard focus
        preview_window.title("Screenshot Viewer")
        
        # Get screen width and height
        preview_window.update_idletasks()  # Ensure dimensions are correct before placing
        screen_width = preview_window.winfo_screenwidth()
        screen_height = preview_window.winfo_screenheight()

        # Desired window size (same as your geometry)
        win_width = 1000
        win_height = 700

        # Calculate position
        x = int(((screen_width // 2) - (win_width // 2)) * 2)
        y = int(((screen_height // 2) - (win_height // 2)) * 2)

        # Set centered position
        preview_window.geometry(f"{win_width}x{win_height}+{x}+{y}")


        current_index = [0]  # Mutable index for inner scope

        # Display image area
        img_label = ctk.CTkLabel(preview_window, text="")
        img_label.pack(expand=True, fill="both", padx=10, pady=10)

        def show_image(index):
            if 0 <= index < len(self.screenshots):
                try:
                    img = Image.open(self.screenshots[index])
                    img.thumbnail((950, 650))  # Resize to fit
                    img_ctk = CTkImage(light_image=img, size=img.size)
                    img_label.configure(image=img_ctk)
                    img_label.image = img_ctk
                    img.close()
                except Exception as e:
                    self.status_label.configure(text=f"Error loading image: {str(e)}", text_color="red")

        def next_image():
            if current_index[0] < len(self.screenshots) - 1:
                current_index[0] += 1
                show_image(current_index[0])

        def prev_image():
            if current_index[0] > 0:
                current_index[0] -= 1
                show_image(current_index[0])

        def delete_current():
            if self.screenshots:
                to_delete = self.screenshots[current_index[0]]
                try:
                    os.remove(to_delete)
                    self.screenshots.remove(to_delete)
                    self.status_label.configure(text=f"Deleted: {Path(to_delete).name}", text_color="orange")
                except Exception as e:
                    self.status_label.configure(text=f"Error deleting: {str(e)}", text_color="red")

            if not self.screenshots:
                preview_window.destroy()
                return

            if current_index[0] >= len(self.screenshots):
                current_index[0] = len(self.screenshots) - 1
            show_image(current_index[0])

        nav_frame = ctk.CTkFrame(preview_window)
        nav_frame.pack(pady=10)

        prev_button = ctk.CTkButton(nav_frame, text="Previous", command=prev_image)
        prev_button.grid(row=0, column=0, padx=10)

        delete_button = ctk.CTkButton(nav_frame, text="Delete", command=delete_current)
        delete_button.grid(row=0, column=1, padx=10)

        next_button = ctk.CTkButton(nav_frame, text="Next", command=next_image)
        next_button.grid(row=0, column=2, padx=10)

        show_image(current_index[0])

    def export_to_pdf(self):
        screenshots_to_export = self.selected_screenshots or self.screenshots
        if not screenshots_to_export:
            self.status_label.configure(text="No screenshots to export!", text_color="red")
            return

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        for i, screenshot_path in enumerate(screenshots_to_export, start=1):
            pdf.add_page()
            pdf.set_font("Helvetica", size=12)
            pdf.cell(200, 10, text=f"Step {i}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
            pdf.image(str(screenshot_path), x=10, y=20, w=190)

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save PDF")
        if save_path:
            pdf.output(save_path)
    
            # Delete screenshots after export
            for screenshot_path in screenshots_to_export:
                try:
                    os.remove(screenshot_path)
                except Exception as e:
                    self.status_label.configure(text=f"Error deleting: {screenshot_path.name}", text_color="red")
    
            self.screenshots.clear()
            self.selected_screenshots.clear()

            self.status_label.configure(text="PDF Exported and Screenshots Deleted!", text_color="green")
            
    def on_closing(self):
        # Delete any remaining screenshots
        for screenshot_path in self.screenshots:
            try:
                os.remove(screenshot_path)
            except Exception as e:
                print(f"Failed to delete {screenshot_path}: {e}")
        self.screenshots.clear()

        # Destroy the main window
        self.root.destroy()

# Run the app
if __name__ == "__main__":
    import time
    root = ctk.CTk()
    app = ScreenshotApp(root)
    root.mainloop()
