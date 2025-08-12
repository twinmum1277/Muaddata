# Muad'Data v17 - Fully Functional Element Viewer + RGB Overlay Tabs
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, colorchooser, simpledialog
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import re

class MathExpressionDialog:
    def __init__(self, parent, title="Enter Mathematical Expression"):
        self.result = None
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("500x300")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))
        
        # Build the dialog interface immediately
        self.build_dialog()
        
        # Make dialog modal
        self.dialog.focus_set()
        self.dialog.wait_window()
    
    def build_dialog(self):
        # Main frame
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Map Math - Mathematical Expression", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions
        instructions = tk.Label(main_frame, text="Enter a mathematical expression using 'x' as the variable.\nExample: x * 0.001 (to convert CPS to ppm)", 
                              font=("Arial", 11), justify=tk.LEFT)
        instructions.pack(pady=(0, 15))
        
        # Expression entry
        tk.Label(main_frame, text="Expression:", font=("Arial", 12)).pack(anchor='w')
        self.expression_entry = tk.Entry(main_frame, font=("Arial", 12), width=50)
        self.expression_entry.pack(fill=tk.X, pady=(5, 15))
        self.expression_entry.insert(0, "x * 0.001")
        self.expression_entry.focus()
        
        # Common expressions frame
        common_frame = tk.Frame(main_frame)
        common_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(common_frame, text="Common expressions:", font=("Arial", 11, "bold")).pack(anchor='w')
        
        # Common expression buttons
        common_expressions = [
            ("Square root", "np.sqrt(x)"),
            ("Log base 10", "np.log10(x)"),
            ("Natural log", "np.log(x)"),
            ("Square", "x ** 2")
        ]
        
        for label, expr in common_expressions:
            btn = tk.Button(common_frame, text=label, command=lambda e=expr: self.expression_entry.delete(0, tk.END) or self.expression_entry.insert(0, e),
                           font=("Arial", 10))
            btn.pack(side=tk.LEFT, padx=(0, 5), pady=2)
        
        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Apply button
        apply_btn = tk.Button(button_frame, text="Apply Expression", command=self.apply_expression, 
                             font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", padx=20)
        apply_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Cancel button
        cancel_btn = tk.Button(button_frame, text="Cancel", command=self.cancel, 
                              font=("Arial", 12), padx=20)
        cancel_btn.pack(side=tk.RIGHT)
        
        # Bind Enter key to apply
        self.expression_entry.bind("<Return>", lambda e: self.apply_expression())
        self.expression_entry.bind("<Escape>", lambda e: self.cancel())
        
        # Bind window close button
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
    
    def apply_expression(self):
        expression = self.expression_entry.get().strip()
        if not expression:
            messagebox.showerror("Error", "Please enter a mathematical expression.")
            return
        
        # Validate expression
        try:
            # Test with a sample value
            x = 1.0
            eval(expression, {"__builtins__": {}}, {"x": x, "np": np})
            self.result = expression
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("Invalid Expression", f"The expression contains an error:\n{str(e)}\n\nPlease check your syntax.")
    
    def cancel(self):
        self.dialog.destroy()

class MuadDataViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Muad'Data - Elemental Map Viewer")

        # Single Element Viewer state
        self.single_matrix = None
        self.single_colormap = tk.StringVar(value='viridis')
        self.single_min = tk.DoubleVar()
        self.single_max = tk.DoubleVar()
        self.max_constraint = tk.IntVar()  # For constraining max slider value (integer only)
        self.show_colorbar = tk.IntVar()
        self.show_scalebar = tk.IntVar()
        self.pixel_size = tk.IntVar(value=1)  # Integer only for pixel size
        self.scale_length = tk.DoubleVar(value=50)
        self.single_file_label = None  # For displaying loaded file info
        self.single_file_name = None   # Store loaded file name
        self._single_colorbar = None   # Store the colorbar object for removal
        self.original_matrix = None    # Store original matrix for math operations

        # RGB Overlay state
        self.rgb_data = {'R': None, 'G': None, 'B': None}
        self.rgb_sliders = {}
        self.rgb_labels = {}
        self.rgb_colors = {'R': '#ff0000', 'G': '#00ff00', 'B': '#0000ff'}  # Default colors
        self.rgb_color_buttons = {}
        self.rgb_gradient_canvases = {}
        self.file_root_label = None
        self.normalize_var = tk.IntVar()

        # Tabs
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill=tk.BOTH, expand=True)

        self.single_tab = tk.Frame(self.tabs)
        self.rgb_tab = tk.Frame(self.tabs)

        self.tabs.add(self.single_tab, text="Element Viewer")
        self.tabs.add(self.rgb_tab, text="RGB Overlay")

        self.build_single_tab()
        self.build_rgb_tab()

    def build_single_tab(self):
        control_frame = tk.Frame(self.single_tab, padx=10, pady=10)
        control_frame.pack(side=tk.LEFT, fill=tk.Y)

        display_frame = tk.Frame(self.single_tab)
        display_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tk.Button(control_frame, text="Load Matrix File", command=self.load_single_file, font=("Arial", 13)).pack(fill=tk.X, pady=(6, 2))

        tk.Label(control_frame, text="Colormap", font=("Arial", 13)).pack()
        cmap_menu = ttk.Combobox(control_frame, textvariable=self.single_colormap, values=plt.colormaps(), font=("Arial", 12))
        cmap_menu.pack(fill=tk.X)
        cmap_menu.bind("<<ComboboxSelected>>", lambda e: self.view_single_map())

        tk.Label(control_frame, text="Min Value", font=("Arial", 13)).pack()
        self.min_slider = tk.Scale(control_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, variable=self.single_min, font=("Arial", 13))
        self.min_slider.pack(fill=tk.X)
        # Update plot when min slider is released (not during dragging)
        self.min_slider.bind("<ButtonRelease-1>", lambda e: self.update_histogram_and_view())

        # Add histogram frame above Max Value
        histogram_frame = tk.Frame(control_frame)
        histogram_frame.pack(fill=tk.X, pady=(5, 0))
        tk.Label(histogram_frame, text="Data Distribution", font=("Arial", 11)).pack()
        self.histogram_canvas = tk.Canvas(histogram_frame, height=60, width=200, bg='white', relief='sunken', bd=1)
        self.histogram_canvas.pack(fill=tk.X, pady=(2, 5))

        tk.Label(control_frame, text="Max Value", font=("Arial", 13)).pack()
        self.max_slider = tk.Scale(control_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, variable=self.single_max, font=("Arial", 13))
        self.max_slider.pack(fill=tk.X)
        # Update plot when max slider is released (not during dragging)
        self.max_slider.bind("<ButtonRelease-1>", lambda e: self.update_histogram_and_view())

        tk.Label(control_frame, text="Set Max", font=("Arial", 13)).pack(pady=(10, 0))
        max_constraint_entry = tk.Entry(control_frame, textvariable=self.max_constraint, font=("Arial", 13))
        max_constraint_entry.pack(fill=tk.X)
        max_constraint_entry.bind("<Return>", lambda e: self.apply_max_constraint())

        tk.Checkbutton(control_frame, text="Show Color Bar", variable=self.show_colorbar, command=self.view_single_map, font=("Arial", 13)).pack(anchor='w')
        tk.Checkbutton(control_frame, text="Show Scale Bar", variable=self.show_scalebar, command=self.view_single_map, font=("Arial", 13)).pack(anchor='w')

        tk.Label(control_frame, text="Pixel size (µm)", font=("Arial", 13)).pack(pady=(10, 0))
        tk.Entry(control_frame, textvariable=self.pixel_size, font=("Arial", 13)).pack(fill=tk.X)

        tk.Label(control_frame, text="Scale bar length (µm)", font=("Arial", 13)).pack(pady=(10, 0))
        tk.Entry(control_frame, textvariable=self.scale_length, font=("Arial", 13)).pack(fill=tk.X)

        tk.Button(control_frame, text="View Map", command=self.view_single_map, font=("Arial", 13)).pack(fill=tk.X, pady=(10, 2))
        
        # Add Map Math button
        tk.Button(control_frame, text="Map Math", command=self.open_map_math, font=("Arial", 13), bg="#FF8C00", fg="black", relief="raised", bd=2).pack(fill=tk.X, pady=(5, 2))
        
        # Add Reset to Original button
        tk.Button(control_frame, text="Reset to Original", command=self.reset_to_original, font=("Arial", 13), bg="#4169E1", fg="black", relief="raised", bd=2).pack(fill=tk.X, pady=(2, 2))
        
        tk.Button(control_frame, text="Save PNG", command=self.save_single_image, font=("Arial", 13)).pack(fill=tk.X)

        # Add a label at the bottom left to display loaded file info
        self.single_file_label = tk.Label(control_frame, text="Loaded file: None", font=("Arial", 11, "italic"), anchor="w", justify="left", wraplength=200)
        self.single_file_label.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        self.single_figure, self.single_ax = plt.subplots(constrained_layout=True)
        self.single_ax.axis('off')
        self.single_canvas = FigureCanvasTkAgg(self.single_figure, master=display_frame)
        self.single_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def build_rgb_tab(self):
        control_frame = tk.Frame(self.rgb_tab, padx=10, pady=10)
        control_frame.pack(side=tk.LEFT, fill=tk.Y)

        display_frame = tk.Frame(self.rgb_tab)
        display_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.file_root_label = tk.Label(control_frame, text="Dataset: None", font=("Arial", 13, "italic"))
        self.file_root_label.pack(pady=(0, 10))

        # Channel mapping for cleaner labels
        channel_labels = {'R': 'Channel 1', 'G': 'Channel 2', 'B': 'Channel 3'}
        default_colors = {'R': '#ff0000', 'G': '#00ff00', 'B': '#0000ff'}

        for ch in ['R', 'G', 'B']:
            channel_label = channel_labels[ch]
            tk.Button(control_frame, text=f"Load {channel_label}", command=lambda c=ch: self.load_rgb_file(c), font=("Arial", 13)).pack(fill=tk.X, pady=(6, 2))
            elem_label = tk.Label(control_frame, text=f"Loaded Element: None", font=("Arial", 13, "italic"))
            elem_label.pack()
            
            # Color picker with channel color as button background
            color_picker_frame = tk.Frame(control_frame)
            color_picker_frame.pack(fill=tk.X, padx=5, pady=2)
            color_btn = tk.Button(color_picker_frame, text=f"Color {ch}", bg=self.rgb_colors[ch], fg='white', font=("Arial", 10, "bold"),
                                  command=lambda c=ch: self.pick_channel_color(c))
            color_btn.pack(side=tk.LEFT, padx=(0, 5))
            
            # Gradient preview
            gradient_canvas = tk.Canvas(color_picker_frame, height=10, width=200)
            gradient_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.draw_gradient(gradient_canvas, self.rgb_colors[ch])
            self.rgb_color_buttons[ch] = color_btn
            self.rgb_gradient_canvases[ch] = gradient_canvas

            max_slider = tk.Scale(control_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, label=f"{channel_label} Max", font=("Arial", 13))
            max_slider.set(1)
            max_slider.pack(fill=tk.X)
            max_slider.bind("<B1-Motion>", lambda e, c=ch: self.view_rgb_overlay())
            self.rgb_sliders[ch] = {'max': max_slider}
            self.rgb_labels[ch] = {'elem': elem_label}

        # Add space for future color bar
        tk.Label(control_frame, text="Color Scale", font=("Arial", 13)).pack(pady=(10, 0))
        self.color_scale_canvas = tk.Canvas(control_frame, height=80, width=200, bg='white', relief='sunken', bd=1)
        self.color_scale_canvas.pack(fill=tk.X, pady=(2, 10))

        tk.Checkbutton(control_frame, text="Normalize to 99th Percentile", variable=self.normalize_var, font=("Arial", 13)).pack(anchor='w', pady=(10, 5))
        tk.Button(control_frame, text="View Overlay", command=self.view_rgb_overlay, font=("Arial", 13)).pack(fill=tk.X, pady=(10, 2))
        tk.Button(control_frame, text="Save RGB Image", command=self.save_rgb_image, font=("Arial", 13)).pack(fill=tk.X)

        self.rgb_figure, self.rgb_ax = plt.subplots(constrained_layout=True)
        self.rgb_ax.axis('off')
        self.rgb_canvas = FigureCanvasTkAgg(self.rgb_figure, master=display_frame)
        self.rgb_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def pick_channel_color(self, channel):
        # Open color chooser and update color for the channel
        channel_labels = {'R': 'Channel 1', 'G': 'Channel 2', 'B': 'Channel 3'}
        initial_color = self.rgb_colors[channel]
        color_code = colorchooser.askcolor(title=f"Pick color for {channel_labels[channel]}", color=initial_color)
        if color_code and color_code[1]:
            self.rgb_colors[channel] = color_code[1]
            # Update button color
            self.rgb_color_buttons[channel].configure(bg=color_code[1])
            # Redraw gradient
            self.draw_gradient(self.rgb_gradient_canvases[channel], color_code[1])
            # Update overlay if visible
            self.view_rgb_overlay()
            # Update color scale
            self.update_color_scale()

    def draw_gradient(self, canvas, color):
        # Accepts either a color name ('red', 'green', 'blue') or a hex color
        canvas.delete("all")
        # If color is a hex string, interpolate from black to that color
        if isinstance(color, str) and color.startswith('#') and len(color) == 7:
            # Get RGB values
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            for i in range(256):
                frac = i / 255.0
                rr = int(r * frac)
                gg = int(g * frac)
                bb = int(b * frac)
                c = f'#{rr:02x}{gg:02x}{bb:02x}'
                canvas.create_line(i, 0, i, 10, fill=c)
        else:
            # Fallback to old behavior for 'red', 'green', 'blue'
            for i in range(256):
                c = {'red': f'#{i:02x}0000', 'green': f'#00{i:02x}00', 'blue': f'#0000{i:02x}'}[color]
                canvas.create_line(i, 0, i, 10, fill=c)

    def load_single_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not path:
            return
        
        # Validate file path
        if not os.path.exists(path):
            messagebox.showerror("Error", f"File not found: {path}")
            return
        
        try:
            if path.endswith('.xlsx'):
                # Try different Excel engines to handle various formats
                df = None
                engines_to_try = ['openpyxl', 'xlrd', 'odf']
                
                for engine in engines_to_try:
                    try:
                        df = pd.read_excel(path, header=None, engine=engine)
                        break  # If successful, break out of the loop
                    except Exception as e:
                        continue  # Try next engine
                
                if df is None:
                    # If all engines failed, try with explicit sheet name
                    try:
                        df = pd.read_excel(path, header=None, sheet_name=0, engine='openpyxl')
                    except Exception:
                        try:
                            df = pd.read_excel(path, header=None, sheet_name=0, engine='xlrd')
                        except Exception as e:
                            raise Exception(f"Could not read Excel file with any engine. Please ensure the file is a valid Excel format. Error: {str(e)}")
            else:
                df = pd.read_csv(path, header=None)
            
            df = df.apply(pd.to_numeric, errors='coerce').dropna(how='all').dropna(axis=1, how='all')
            mat = df.to_numpy()
            self.single_matrix = mat
            # Store original matrix for math operations
            self.original_matrix = np.array(mat, copy=True)
            # Update min/max values and sliders
            min_val = np.nanmin(mat)
            max_val = np.nanmax(mat)
            self.single_min.set(min_val)
            self.single_max.set(max_val)
            self.min_slider.config(from_=min_val, to=max_val)
            self.max_slider.config(from_=min_val, to=max_val)
            self.min_slider.set(min_val)
            self.max_slider.set(max_val)
            # Initialize max constraint with actual max value (as integer)
            self.max_constraint.set(int(max_val))
            # Update loaded file label
            self.single_file_name = os.path.basename(path)
            self.update_file_label()
            # Update histogram
            self.update_histogram()
            self.view_single_map()
        except Exception as e:
            error_msg = f"Failed to load matrix file:\n{e}\n\nFile path: {path}\nFile exists: {os.path.exists(path) if path else 'No path'}"
            messagebox.showerror("Error", error_msg)
            self.single_file_name = None
            self.update_file_label()

    def apply_max_constraint(self):
        """Apply the max constraint to limit the max slider value."""
        if self.single_matrix is None:
            messagebox.showwarning("No Data", "Please load a matrix file first.")
            return
        
        constraint_value = self.max_constraint.get()
        if constraint_value <= 0:
            messagebox.showerror("Invalid Value", "Max constraint must be greater than 0.")
            return
        
        # Get current data range
        min_val = np.nanmin(self.single_matrix)
        max_val = np.nanmax(self.single_matrix)
        
        # Apply constraint
        constrained_max = min(max_val, constraint_value)
        
        # Update slider range and current value
        self.max_slider.config(to=constrained_max)
        if self.single_max.get() > constrained_max:
            self.single_max.set(constrained_max)
        
        # Update histogram to show new range
        self.update_histogram_and_view()

    def update_histogram(self):
        """Update the data distribution histogram."""
        if self.single_matrix is None:
            return
        
        # Clear the canvas
        self.histogram_canvas.delete("all")
        
        # Get data and create histogram (only in current slider range)
        data = self.single_matrix.flatten()
        data = data[~np.isnan(data)]  # Remove NaN values
        current_min = self.single_min.get()
        current_max = self.single_max.get()
        data = data[(data >= current_min) & (data <= current_max)]
        
        if len(data) == 0:
            return
        
        # Create histogram
        hist, bin_edges = np.histogram(data, bins=50)
        
        # Get canvas dimensions
        canvas_width = self.histogram_canvas.winfo_width()
        canvas_height = self.histogram_canvas.winfo_height()
        
        if canvas_width <= 1:  # Canvas not yet drawn
            canvas_width = 200
            canvas_height = 60
        
        # Normalize histogram to fit canvas
        max_hist = np.max(hist)
        if max_hist == 0:
            return
        
        # Draw smooth curve like photo editors
        x_coords = []
        y_coords = []
        for i, count in enumerate(hist):
            if count > 0:
                x = (i / len(hist)) * canvas_width
                y = canvas_height - 5 - (count / max_hist) * (canvas_height - 10)
                x_coords.append(x)
                y_coords.append(y)
        
        # Create smooth curve with more points for better smoothing
        if len(x_coords) >= 2:
            # Create filled area under the curve (like photo editors)
            fill_points = []
            
            # Start at bottom-left
            fill_points.extend([0, canvas_height - 5])
            
            # Add curve points
            for i in range(len(x_coords)):
                fill_points.extend([x_coords[i], y_coords[i]])
            
            # End at bottom-right
            fill_points.extend([canvas_width, canvas_height - 5])
            
            # Draw filled area with light gray
            if len(fill_points) >= 6:
                self.histogram_canvas.create_polygon(fill_points, fill='lightgray', outline='', smooth=True)
            
            # Draw the curve line on top
            curve_points = []
            for i in range(len(x_coords)):
                curve_points.extend([x_coords[i], y_coords[i]])
            
            # Draw the smooth curve with anti-aliasing
            if len(curve_points) >= 4:
                self.histogram_canvas.create_line(curve_points, fill='darkgray', width=2, smooth=True, capstyle='round')
        
        # Add min/max indicators (always at left/right of canvas now)
        self.histogram_canvas.create_line(0, 0, 0, canvas_height, fill='red', width=2)
        self.histogram_canvas.create_line(canvas_width, 0, canvas_width, canvas_height, fill='blue', width=2)

    def update_color_scale(self):
        """Update the color scale based on loaded channels."""
        # Clear the canvas
        self.color_scale_canvas.delete("all")
        
        # Count loaded channels
        loaded_channels = [ch for ch in ['R', 'G', 'B'] if self.rgb_data[ch] is not None]
        num_channels = len(loaded_channels)
        
        if num_channels == 0:
            # No channels loaded - show empty canvas
            return
        elif num_channels == 1:
            # Single channel - show single color bar
            ch = loaded_channels[0]
            color = self.rgb_colors[ch]
            self.color_scale_canvas.create_rectangle(10, 20, 190, 60, fill=color, outline='black')
            self.color_scale_canvas.create_text(100, 70, text=f"Channel {ch}", font=("Arial", 10))
            
        elif num_channels == 2:
            # Two channels - create linear gradient
            ch1, ch2 = loaded_channels[0], loaded_channels[1]
            color1 = self.rgb_colors[ch1]
            color2 = self.rgb_colors[ch2]
            
            # Create gradient bar
            bar_width = 180
            bar_height = 40
            x_start, y_start = 10, 20
            
            # Draw gradient by interpolating colors
            for i in range(bar_width):
                frac = i / (bar_width - 1)
                # Interpolate between the two colors
                r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
                r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)
                
                r = int(r1 * (1 - frac) + r2 * frac)
                g = int(g1 * (1 - frac) + g2 * frac)
                b = int(b1 * (1 - frac) + b2 * frac)
                
                gradient_color = f'#{r:02x}{g:02x}{b:02x}'
                self.color_scale_canvas.create_line(x_start + i, y_start, x_start + i, y_start + bar_height, 
                                                  fill=gradient_color, width=1)
            
            # Add labels
            self.color_scale_canvas.create_text(10, 70, text=f"Ch{ch1}", font=("Arial", 8), anchor='w')
            self.color_scale_canvas.create_text(190, 70, text=f"Ch{ch2}", font=("Arial", 8), anchor='e')
            
        elif num_channels == 3:
            # Three channels - create triangular color space
            ch1, ch2, ch3 = loaded_channels[0], loaded_channels[1], loaded_channels[2]
            color1 = self.rgb_colors[ch1]
            color2 = self.rgb_colors[ch2]
            color3 = self.rgb_colors[ch3]
            
            # Create triangular color space
            # Triangle vertices (equilateral triangle)
            center_x, center_y = 100, 40
            radius = 30
            
            # Draw triangle outline
            points = [
                center_x, center_y - radius,  # Top
                center_x - radius * 0.866, center_y + radius * 0.5,  # Bottom left
                center_x + radius * 0.866, center_y + radius * 0.5   # Bottom right
            ]
            self.color_scale_canvas.create_polygon(points, outline='black', width=2)
            
            # Fill triangle with interpolated colors
            # Create a grid of points and interpolate colors
            for y in range(int(center_y - radius), int(center_y + radius + 1), 2):
                for x in range(int(center_x - radius), int(center_x + radius + 1), 2):
                    # Check if point is inside triangle
                    if self.point_in_triangle(x, y, points):
                        # Calculate barycentric coordinates
                        bary = self.barycentric_coords(x, y, points)
                        if bary:
                            # Interpolate colors using barycentric coordinates
                            color = self.interpolate_colors([color1, color2, color3], bary)
                            self.color_scale_canvas.create_oval(x-1, y-1, x+1, y+1, fill=color, outline='')
            
            # Add channel labels at vertices
            self.color_scale_canvas.create_text(center_x, center_y - radius - 5, text=f"Ch{ch1}", font=("Arial", 8))
            self.color_scale_canvas.create_text(center_x - radius * 0.866, center_y + radius * 0.5 + 10, text=f"Ch{ch2}", font=("Arial", 8))
            self.color_scale_canvas.create_text(center_x + radius * 0.866, center_y + radius * 0.5 + 10, text=f"Ch{ch3}", font=("Arial", 8))

    def point_in_triangle(self, x, y, triangle_points):
        """Check if a point is inside a triangle."""
        x1, y1, x2, y2, x3, y3 = triangle_points
        
        # Calculate barycentric coordinates
        denominator = (y2 - y3) * (x1 - x3) + (x3 - x2) * (y1 - y3)
        if denominator == 0:
            return False
            
        u = ((x - x3) * (y2 - y3) - (y - y3) * (x2 - x3)) / denominator
        v = ((x - x1) * (y3 - y1) - (y - y1) * (x3 - x1)) / denominator
        w = 1 - u - v
        
        return u >= 0 and v >= 0 and w >= 0

    def barycentric_coords(self, x, y, triangle_points):
        """Calculate barycentric coordinates of a point relative to a triangle."""
        x1, y1, x2, y2, x3, y3 = triangle_points
        
        denominator = (y2 - y3) * (x1 - x3) + (x3 - x2) * (y1 - y3)
        if denominator == 0:
            return None
            
        u = ((x - x3) * (y2 - y3) - (y - y3) * (x2 - x3)) / denominator
        v = ((x - x1) * (y3 - y1) - (y - y1) * (x3 - x1)) / denominator
        w = 1 - u - v
        
        return [u, v, w]

    def interpolate_colors(self, colors, barycentric_coords):
        """Interpolate colors using barycentric coordinates."""
        r, g, b = 0, 0, 0
        
        for i, color in enumerate(colors):
            weight = barycentric_coords[i]
            r += int(color[1:3], 16) * weight
            g += int(color[3:5], 16) * weight
            b += int(color[5:7], 16) * weight
        
        return f'#{int(r):02x}{int(g):02x}{int(b):02x}'

    def update_histogram_and_view(self):
        """Update both histogram and the main view."""
        self.update_histogram()
        self.view_single_map()

    def view_single_map(self):
        if self.single_matrix is None:
            return
        mat = np.array(self.single_matrix, dtype=float)
        mat[np.isnan(mat)] = 0
        # Update min/max values from sliders in case they changed
        vmin = self.single_min.get()
        vmax = self.single_max.get()
        self.single_ax.clear()
        im = self.single_ax.imshow(mat, cmap=self.single_colormap.get(), vmin=vmin, vmax=vmax)
        self.single_ax.axis('off')
        
        # Remove previous colorbar if it exists
        if hasattr(self, '_single_colorbar') and self._single_colorbar is not None:
            try:
                self._single_colorbar.remove()
            except Exception:
                pass
            self._single_colorbar = None
        
        # Add colorbar with custom formatting
        if self.show_colorbar.get():
            # Create colorbar with reduced height and custom ticks
            self._single_colorbar = self.single_figure.colorbar(im, ax=self.single_ax, fraction=0.023, pad=0.04, aspect=20)
            # Set custom ticks (max, middle, min)
            ticks = [vmin, (vmin + vmax) / 2, vmax]
            self._single_colorbar.set_ticks(ticks)
            # Format with 1-2 decimal places based on value range
            if vmax - vmin > 100:
                format_str = '{:.1f}'
            else:
                format_str = '{:.2f}'
            self._single_colorbar.set_ticklabels([format_str.format(vmin), format_str.format((vmin + vmax) / 2), format_str.format(vmax)])
            # Set font properties
            self._single_colorbar.ax.tick_params(labelsize=8)
            for label in self._single_colorbar.ax.get_yticklabels():
                label.set_fontname('Arial')
        
        # Add scale bar underneath the colorbar (outside the image area)
        if self.show_scalebar.get():
            bar_length = self.scale_length.get() / self.pixel_size.get()
            # Position scale bar below the image and colorbar with more padding
            x_start = 10
            x_end = x_start + bar_length
            y_pos = mat.shape[0] + 25  # More padding below the image area
            
            # Draw scale bar
            self.single_ax.plot([x_start, x_end], [y_pos, y_pos], color='black', lw=2, solid_capstyle='butt')
            # Add scale bar label to the right of the bar
            label_offset = 10
            self.single_ax.text(x_end + label_offset, y_pos, f"{int(self.scale_length.get())} µm", 
                               color='black', fontsize=6, ha='left', va='center', fontname='Arial')
        
        # Use constrained_layout instead of tight_layout to prevent image shifting
        self.single_figure.set_constrained_layout(True)
        self.single_canvas.draw()

    def save_single_image(self):
        if self.single_matrix is None:
            return
        out_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
        if out_path:
            self.single_figure.savefig(out_path, dpi=300, bbox_inches='tight')

    def load_rgb_file(self, channel):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not path:
            return
        
        # Validate file path
        if not os.path.exists(path):
            messagebox.showerror("Error", f"File not found: {path}")
            return
        
        try:
            if path.endswith('.xlsx'):
                # Try different Excel engines to handle various formats
                df = None
                engines_to_try = ['openpyxl', 'xlrd', 'odf']
                
                for engine in engines_to_try:
                    try:
                        df = pd.read_excel(path, header=None, engine=engine)
                        break  # If successful, break out of the loop
                    except Exception as e:
                        continue  # Try next engine
                
                if df is None:
                    # If all engines failed, try with explicit sheet name
                    try:
                        df = pd.read_excel(path, header=None, sheet_name=0, engine='openpyxl')
                    except Exception:
                        try:
                            df = pd.read_excel(path, header=None, sheet_name=0, engine='xlrd')
                        except Exception as e:
                            raise Exception(f"Could not read Excel file with any engine. Please ensure the file is a valid Excel format. Error: {str(e)}")
            else:
                df = pd.read_csv(path, header=None)
            
            df = df.apply(pd.to_numeric, errors='coerce').dropna(how='all').dropna(axis=1, how='all')
            mat = df.to_numpy()
            self.rgb_data[channel] = mat
            file_name = os.path.basename(path)
            root_name = file_name.split()[0]
            elem = next((part for part in file_name.split() if any(e in part for e in ['ppm', 'CPS'])), 'Unknown')
            self.rgb_labels[channel]['elem'].config(text=f"Loaded Element: {elem.split('_')[0]}")
            if self.file_root_label.cget("text") == "Dataset: None":
                self.file_root_label.config(text=f"Dataset: {root_name}")
            max_val = float(np.nanmax(mat))
            if np.isfinite(max_val):
                self.rgb_sliders[channel]['max'].config(from_=0, to=max_val)
                self.rgb_sliders[channel]['max'].set(max_val)
            messagebox.showinfo("Loaded", f"{channel} channel loaded with shape {mat.shape}")
            # Update color scale
            self.update_color_scale()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load {channel} channel:\n{e}")

    def view_rgb_overlay(self, event=None):
        def rescale(mat, vmax):
            return np.clip(mat / (vmax + 1e-6), 0, 1)
        def get_scaled_matrix(channel):
            mat = self.rgb_data[channel]
            vmax = self.rgb_sliders[channel]['max'].get()
            if self.normalize_var.get():
                p99 = np.nanpercentile(mat, 99)
                vmax = min(vmax, p99)
            scaled = rescale(mat, vmax)
            scaled[np.isnan(scaled)] = 0
            return scaled
        shape = None
        for ch in 'RGB':
            if self.rgb_data[ch] is not None:
                shape = self.rgb_data[ch].shape
                break
        if shape is None:
            messagebox.showwarning("No Data", "Please load at least one channel.")
            return
        composite = []
        for ch in 'RGB':
            mat = self.rgb_data[ch]
            if mat is None:
                composite.append(np.zeros(shape))
            else:
                composite.append(get_scaled_matrix(ch))
        # Now, instead of stacking as RGB, use the selected color for each channel
        rgb = np.zeros((shape[0], shape[1], 3), dtype=float)
        for idx, ch in enumerate('RGB'):
            color_hex = self.rgb_colors[ch]
            r = int(color_hex[1:3], 16) / 255.0
            g = int(color_hex[3:5], 16) / 255.0
            b = int(color_hex[5:7], 16) / 255.0
            # Add the channel's scaled matrix times the color
            rgb[..., 0] += composite[idx] * r
            rgb[..., 1] += composite[idx] * g
            rgb[..., 2] += composite[idx] * b
        rgb = np.clip(rgb, 0, 1)
        rgb[np.isnan(rgb)] = 0
        black_mask = np.all(rgb == 0, axis=2)
        rgb[black_mask] = [0, 0, 0]
        self.rgb_ax.clear()
        self.rgb_ax.imshow(rgb)
        self.rgb_ax.axis('off')
        self.rgb_figure.tight_layout()
        self.rgb_canvas.draw()

    def save_rgb_image(self):
        if all(self.rgb_data[c] is None for c in 'RGB'):
            return
        out_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
        if out_path:
            self.rgb_figure.savefig(out_path, dpi=300, bbox_inches='tight')

    def open_map_math(self):
        """Open the map math dialog and apply mathematical expressions to the loaded matrix."""
        if self.single_matrix is None:
            messagebox.showwarning("No Data", "Please load a matrix file first.")
            return
        
        # Create and show the math expression dialog
        dialog = MathExpressionDialog(self.root)
        
        if dialog.result:
            try:
                # Store original matrix if not already stored
                if self.original_matrix is None:
                    self.original_matrix = np.array(self.single_matrix, copy=True)
                
                # Create a copy of the current matrix for processing
                mat = np.array(self.single_matrix, dtype=float)
                
                # Create a mask for non-empty cells (where there are actual values)
                # We'll consider cells with values > 0 as non-empty
                non_empty_mask = (mat > 0) & ~np.isnan(mat)
                
                # Apply the expression only to non-empty cells
                result_mat = np.array(mat, copy=True)
                
                # For each non-empty cell, apply the expression
                for i in range(mat.shape[0]):
                    for j in range(mat.shape[1]):
                        if non_empty_mask[i, j]:
                            x = mat[i, j]
                            try:
                                # Safely evaluate the expression for this cell
                                result = eval(dialog.result, {"__builtins__": {}}, {"x": x, "np": np})
                                result_mat[i, j] = result
                            except Exception as e:
                                messagebox.showerror("Evaluation Error", f"Error evaluating expression for cell [{i},{j}]:\n{str(e)}")
                                return
                
                # Update the current matrix with the result
                self.single_matrix = result_mat
                
                # Update min/max values and sliders
                min_val = np.nanmin(result_mat)
                max_val = np.nanmax(result_mat)
                self.single_min.set(min_val)
                self.single_max.set(max_val)
                self.min_slider.config(from_=min_val, to=max_val)
                self.max_slider.config(from_=min_val, to=max_val)
                self.min_slider.set(min_val)
                self.max_slider.set(max_val)
                
                # Update max constraint
                self.max_constraint.set(int(max_val))
                
                # Update histogram and view
                self.update_histogram()
                self.view_single_map()
                
                # Update file label to show modification status
                self.update_file_label()
                
                # Ask user if they want to save the result
                save_result = messagebox.askyesno("Save Result", 
                                                "Expression applied successfully!\n\nWould you like to save the result to a file?")
                
                if save_result:
                    self.save_math_result(result_mat, dialog.result)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to apply expression:\n{str(e)}")
    
    def save_math_result(self, result_matrix, expression):
        """Save the math result to a file with automatic naming."""
        if self.single_file_name is None:
            # Fallback if no original filename
            default_name = "math_result.xlsx"
        else:
            # Create filename with _math suffix
            name_without_ext = os.path.splitext(self.single_file_name)[0]
            default_name = f"{name_without_ext}_math.xlsx"
        
        # Ask user for save location with pre-filled name
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv")
            ],
            initialfile=default_name
        )
        
        if save_path:
            try:
                if save_path.endswith('.xlsx'):
                    # Save as Excel
                    df = pd.DataFrame(result_matrix)
                    df.to_excel(save_path, header=False, index=False)
                elif save_path.endswith('.csv'):
                    # Save as CSV
                    df = pd.DataFrame(result_matrix)
                    df.to_csv(save_path, header=False, index=False)
                
                messagebox.showinfo("Saved", 
                                  f"Math result saved successfully!\n\n"
                                  f"File: {os.path.basename(save_path)}\n"
                                  f"Expression: {expression}\n"
                                  f"Matrix shape: {result_matrix.shape}")
                
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save the result:\n{str(e)}")
    
    def reset_to_original(self):
        """Reset the current matrix to the original loaded matrix."""
        if self.original_matrix is not None:
            self.single_matrix = np.array(self.original_matrix, copy=True)
            
            # Update min/max values and sliders
            min_val = np.nanmin(self.single_matrix)
            max_val = np.nanmax(self.single_matrix)
            self.single_min.set(min_val)
            self.single_max.set(max_val)
            self.min_slider.config(from_=min_val, to=max_val)
            self.max_slider.config(from_=min_val, to=max_val)
            self.min_slider.set(min_val)
            self.max_slider.set(max_val)
            
            # Update max constraint
            self.max_constraint.set(int(max_val))
            
            # Update histogram and view
            self.update_histogram()
            self.view_single_map()
            
            # Update file label
            self.update_file_label()
            
            messagebox.showinfo("Reset", "Matrix reset to original values.")
        else:
            messagebox.showwarning("No Original", "No original matrix to reset to.")

    def is_matrix_modified(self):
        """Check if the current matrix has been modified from the original."""
        if self.original_matrix is None or self.single_matrix is None:
            return False
        return not np.array_equal(self.single_matrix, self.original_matrix)
    
    def update_file_label(self):
        """Update the file label to show current status."""
        if self.single_file_name is None:
            self.single_file_label.config(text="Loaded file: None")
        else:
            base_text = f"Loaded file: {self.single_file_name}"
            if self.is_matrix_modified():
                base_text += " (Modified)"
            self.single_file_label.config(text=base_text)

def main():
    root = tk.Tk()
    root.geometry("1100x700")
    app = MuadDataViewer(root)
    root.mainloop()

if __name__ == '__main__':
    main() 