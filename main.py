"""
Created by MattFor on 15/05/2025

Copyright © 2025 by MattFor.

This software is provided under a **Proprietary License**.
Please refer to the `LICENSE` file in this repository for the full terms.
"""

import os
import json
import winreg
import ctypes
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


def is_admin():
	"""Checks if the program is running with administrator privileges"""
	try:
		return ctypes.windll.shell32.IsUserAnAdmin()
	except:
		return False


def get_user_path():
	"""Gets the value of the user PATH variable"""
	key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ)
	try:
		value, _ = winreg.QueryValueEx(key, "Path")
		return value
	except:
		return ""
	finally:
		winreg.CloseKey(key)


def get_system_path():
	"""Gets the value of the system PATH variable"""
	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", 0, winreg.KEY_READ)
	try:
		value, _ = winreg.QueryValueEx(key, "Path")
		return value
	except:
		return ""
	finally:
		winreg.CloseKey(key)


def set_system_path(path_value):
	"""Sets the value of the system PATH variable"""
	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", 0, winreg.KEY_WRITE | winreg.KEY_READ)
	winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, path_value)
	winreg.CloseKey(key)


def set_user_path(path_value):
	"""Sets the value of the user PATH variable"""
	key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_WRITE | winreg.KEY_READ)
	winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, path_value)
	winreg.CloseKey(key)


def broadcast_environment_change():
	"""Notifies the system about environment variable changes"""
	# Constants for sending WM_SETTINGCHANGE
	HWND_BROADCAST = 0xFFFF
	WM_SETTINGCHANGE = 0x001A
	SMTO_ABORTIFHUNG = 0x0002

	try:
		ctypes.windll.user32.SendMessageTimeoutW(
			HWND_BROADCAST, WM_SETTINGCHANGE, 0,
			ctypes.create_unicode_buffer("Environment"),
			SMTO_ABORTIFHUNG, 5000, ctypes.byref(ctypes.c_ulong())
		)
	except:
		pass


class PathEditor:
	def __init__(self, root):
		self.root = root
		root.title("PATH Editor")
		root.geometry("700x500")
		root.minsize(600, 400)

		if not is_admin():
			messagebox.showerror("Error", "The program is running without administrator privileges. Some changes may not be saved.")

		selector_frame = ttk.Frame(root, padding="10")
		selector_frame.pack(fill=tk.X)

		ttk.Label(selector_frame, text="Select PATH variable type:").pack(side=tk.LEFT, padx=(0, 10))

		self.var_type = tk.StringVar(value="user")
		ttk.Radiobutton(selector_frame, text="User", variable=self.var_type, value="user", command=self.load_path).pack(side=tk.LEFT, padx=(0, 10))
		ttk.Radiobutton(selector_frame, text="System", variable=self.var_type, value="system", command=self.load_path).pack(side=tk.LEFT)

		main_frame = ttk.Frame(root, padding="10")
		main_frame.pack(fill=tk.BOTH, expand=True)

		path_frame = ttk.LabelFrame(main_frame, text="Paths in PATH Variable", padding="10")
		path_frame.pack(fill=tk.BOTH, expand=True)

		list_frame = ttk.Frame(path_frame)
		list_frame.pack(fill=tk.BOTH, expand=True)

		scrollbar = ttk.Scrollbar(list_frame)
		scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

		self.path_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Consolas", 10))
		self.path_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

		self.path_listbox.config(yscrollcommand=scrollbar.set)
		scrollbar.config(command=self.path_listbox.yview)

		button_frame = ttk.Frame(path_frame)
		button_frame.pack(fill=tk.X, pady=(10, 0))

		ttk.Button(button_frame, text="Add", command=self.add_path).pack(side=tk.LEFT, padx=(0, 5))
		ttk.Button(button_frame, text="Edit", command=self.edit_path).pack(side=tk.LEFT, padx=(0, 5))
		ttk.Button(button_frame, text="Delete", command=self.delete_path).pack(side=tk.LEFT, padx=(0, 5))
		ttk.Button(button_frame, text="Up", command=lambda: self.move_path(-1)).pack(side=tk.LEFT, padx=(0, 5))
		ttk.Button(button_frame, text="Down", command=lambda: self.move_path(1)).pack(side=tk.LEFT)

		backup_frame = ttk.Frame(root, padding="10")
		backup_frame.pack(fill=tk.X)

		ttk.Button(backup_frame, text="Save PATH to File", command=self.export_path).pack(side=tk.LEFT, padx=(0, 5))
		ttk.Button(backup_frame, text="Restore from File", command=self.import_path).pack(side=tk.LEFT)

		bottom_frame = ttk.Frame(root, padding="10")
		bottom_frame.pack(fill=tk.X)

		ttk.Button(bottom_frame, text="Save", command=self.save_path).pack(side=tk.RIGHT, padx=(5, 0))
		ttk.Button(bottom_frame, text="Cancel", command=root.destroy).pack(side=tk.RIGHT)

		self.load_path()

	def load_path(self):
		"""Loads the current value of the PATH variable"""
		self.path_listbox.delete(0, tk.END)

		try:
			if self.var_type.get() == "user":
				path_value = get_user_path()
			else:
				path_value = get_system_path()

			# Split the PATH string into individual paths
			path_entries = [p for p in path_value.split(';') if p]

			# Add paths to the listbox
			for path in path_entries:
				self.path_listbox.insert(tk.END, path)

		except Exception as e:
			messagebox.showerror("Error", f"Failed to load PATH variable: {str(e)}")

	def add_path(self):
		"""Adds a new path to the PATH variable"""
		# Open the directory selection dialogue
		new_path = filedialog.askdirectory(title="Select directory to add to PATH")

		if new_path:
			new_path = new_path.replace('/', '\\')

			# Checks if the path already exists
			path_exists = False
			for i in range(self.path_listbox.size()):
				if self.path_listbox.get(i) == new_path:
					path_exists = True
					break

			if path_exists:
				messagebox.showinfo("Information", "The specified path is already added to the PATH variable")
			else:
				self.path_listbox.insert(tk.END, new_path)

	def edit_path(self):
		"""Edits the selected path"""
		selected = self.path_listbox.curselection()

		if not selected:
			messagebox.showinfo("Information", "Select a path to edit")
			return

		current_path = self.path_listbox.get(selected[0])

		# Create an editing dialogue window
		edit_window = tk.Toplevel(self.root)
		edit_window.title("Edit Path")
		edit_window.geometry("600x100")
		edit_window.resizable(True, False)
		edit_window.transient(self.root)
		edit_window.grab_set()

		# Add input field and buttons
		ttk.Label(edit_window, text="Path:").pack(side=tk.LEFT, padx=(10, 5), pady=10)

		path_var = tk.StringVar(value=current_path)
		path_entry = ttk.Entry(edit_window, width=50, textvariable=path_var)
		path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=10)

		def browse_path():
			new_path = filedialog.askdirectory(title="Select Directory")
			if new_path:
				path_var.set(new_path)

		ttk.Button(edit_window, text="Browse...", command=browse_path).pack(side=tk.LEFT, padx=(0, 10), pady=10)

		button_frame = ttk.Frame(edit_window)
		button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

		def save_edit():
			new_path = path_var.get()
			if new_path:
				self.path_listbox.delete(selected[0])
				self.path_listbox.insert(selected[0], new_path)
			edit_window.destroy()

		ttk.Button(button_frame, text="OK", command=save_edit).pack(side=tk.RIGHT, padx=(5, 0))
		ttk.Button(button_frame, text="Cancel", command=edit_window.destroy).pack(side=tk.RIGHT)

	def delete_path(self):
		"""Deletes the selected path"""
		selected = self.path_listbox.curselection()

		if not selected:
			messagebox.showinfo("Information", "Select a path to delete")
			return

		confirm = messagebox.askyesno("Confirmation", "Are you sure you want to delete the selected path?")

		if confirm:
			self.path_listbox.delete(selected[0])

	def move_path(self, direction):
		"""Moves the selected path up or down"""
		selected = self.path_listbox.curselection()

		if not selected:
			messagebox.showinfo("Information", "Select a path to move")
			return

		index = selected[0]

		# Determine a new position
		new_index = index + direction

		# Check if the new position is outside the list bounds
		if new_index < 0 or new_index >= self.path_listbox.size():
			return

		# Save the current path
		path = self.path_listbox.get(index)

		# Delete the current path and insert it at the new position
		self.path_listbox.delete(index)
		self.path_listbox.insert(new_index, path)

		# Select the path at the new position
		self.path_listbox.selection_clear(0, tk.END)
		self.path_listbox.selection_set(new_index)
		self.path_listbox.see(new_index)

	def save_path(self):
		"""Saves changes to the PATH variable"""
		paths = []
		for i in range(self.path_listbox.size()):
			paths.append(self.path_listbox.get(i))

		# Join paths into a single string
		path_value = ';'.join(paths)

		try:
			if self.var_type.get() == "user":
				set_user_path(path_value)
			else:
				set_system_path(path_value)

			# Broadcast environment variable change
			broadcast_environment_change()

			messagebox.showinfo("Success", "PATH variable successfully updated")
		except Exception as e:
			messagebox.showerror("Error", f"Failed to save PATH variable: {str(e)}")

	def export_path(self):
		"""Saves the current PATH value to a file"""
		# Get the current date and time
		now = datetime.datetime.now()
		date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")

		# Determine prefix based on variable type
		if self.var_type.get() == "user":
			# Get current username
			username = os.environ.get('USERNAME', 'user')
			prefix = username
		else:
			prefix = "sys"

		# Format default filename
		default_filename = f"PATH_backup_{prefix}_{date_time_str}.json"

		# Open save file dialog
		filepath = filedialog.asksaveasfilename(
			title="Save PATH to file",
			defaultextension=".json",
			initialfile=default_filename,
			filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
		)

		if not filepath:
			return

		try:
			# Collect paths from the listbox
			paths = []
			for i in range(self.path_listbox.size()):
				paths.append(self.path_listbox.get(i))

			# Create data dictionary
			data = {
				"path_value": ';'.join(paths),
				"path_entries": paths,
				"backup_date": now.isoformat(),
				"path_type": self.var_type.get(),
				"version": "1.0"
			}

			# Save to JSON file
			with open(filepath, 'w', encoding='utf-8') as f:
				json.dump(data, f, indent=4, ensure_ascii=False)

			messagebox.showinfo("Success", f"PATH backup successfully saved to file:\n{filepath}")
		except Exception as e:
			messagebox.showerror("Error", f"Failed to save PATH to file: {str(e)}")

	def import_path(self):
		"""Loads the PATH value from a file"""
		# Open load file dialogue
		filepath = filedialog.askopenfilename(
			title="Load PATH from file",
			filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
		)

		if not filepath:
			return

		try:
			# Load data from a file
			with open(filepath, 'r', encoding='utf-8') as f:
				data = json.load(f)

			# Check file format
			if "path_entries" not in data:
				raise ValueError("Invalid PATH backup file format")

			# Show backup information
			backup_date = datetime.datetime.fromisoformat(data["backup_date"]).strftime("%d.%m.%Y %H:%M:%S")
			path_type = "User" if data.get("path_type") == "user" else "System"

			confirm_msg = f"Backup creation date: {backup_date}\nPATH type: {path_type}\n\nRestore this PATH copy?"
			confirm = messagebox.askyesno("Restoration Confirmation", confirm_msg)

			if not confirm:
				return

			# Set PATH type
			if "path_type" in data:
				self.var_type.set(data["path_type"])

			# Clear the current list
			self.path_listbox.delete(0, tk.END)

			# Add paths from a file
			for path in data["path_entries"]:
				self.path_listbox.insert(tk.END, path)

			messagebox.showinfo("Success", "PATH backup successfully loaded from file")
		except Exception as e:
			messagebox.showerror("Error", f"Failed to load PATH from file: {str(e)}")


if __name__ == "__main__":
	root = tk.Tk()
	app = PathEditor(root)
	root.mainloop()
