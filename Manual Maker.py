import os
from tkinter import filedialog
from PIL import ImageGrab, Image, ImageTk
from fpdf import FPDF
import customtkinter as ctk
from pathlib import Path
from pynput import keyboard


class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot Documentation Tool")
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

        self.status_label = ctk.CTkLabel(root, text="Status: Idle", text_color="gray")
        self.status_label.pack(pady=10)

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
        self.screenshots_dir = Path.home() / "Desktop" / "screenshots"
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
        # Create a new window for the preview using customtkinter
        preview_window = ctk.CTkToplevel(self.root)
        preview_window.title("Preview Screenshots")
        preview_window.geometry("800x600")

        # Display each screenshot with a delete button
        for i, screenshot_path in enumerate(self.screenshots):
            img = Image.open(screenshot_path)
            img.thumbnail((150, 150))  # Resize for thumbnail view
            img = ImageTk.PhotoImage(img)

            img_label = ctk.CTkLabel(preview_window, image=img, text="")
            img_label.image = img  # Keep a reference to avoid garbage collection
            img_label.grid(row=i // 4, column=(i % 4) * 2, padx=10, pady=10)

            delete_button = ctk.CTkButton(preview_window, text="Delete", command=lambda path=screenshot_path: self.delete_screenshot(path))
            delete_button.grid(row=i // 4, column=(i % 4) * 2 + 1, padx=10, pady=10)

        # Confirm selection button
        confirm_button = ctk.CTkButton(preview_window, text="Confirm Selection", command=lambda: self.finalize_selection(preview_window))
        confirm_button.pack(pady=20)

    def delete_screenshot(self, screenshot_path):
        # Remove screenshot from list and delete the file
        if screenshot_path in self.screenshots:
            self.screenshots.remove(screenshot_path)
            os.remove(screenshot_path)
            self.status_label.configure(text=f"Deleted: {screenshot_path.name}", text_color="orange")

    def finalize_selection(self, preview_window):
        # Save the final selection and close the preview
        self.selected_screenshots = self.screenshots.copy()
        preview_window.destroy()
        self.status_label.configure(text="Selection finalized", text_color="green")

    def export_to_pdf(self):
        if not self.selected_screenshots:
            self.status_label.configure(text="No screenshots to export!", text_color="red")
            return

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        for i, screenshot_path in enumerate(self.selected_screenshots, start=1):
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt=f"Step {i}", ln=True, align="C")
            pdf.image(str(screenshot_path), x=10, y=20, w=190)

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save PDF")
        if save_path:
            pdf.output(save_path)
            self.status_label.configure(text="PDF Exported Successfully!", text_color="green")
        else:
            self.status_label.configure(text="Export Cancelled", text_color="orange")


# Run the app
if __name__ == "__main__":
    import time
    root = ctk.CTk()
    app = ScreenshotApp(root)
    root.mainloop()
