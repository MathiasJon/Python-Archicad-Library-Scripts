"""
================================================================================
SCRIPT: Get All Layers 
================================================================================

Version:        1.0
Date:           2025-11-01
API Type:       Official Archicad API
Category:       Data Retrieval - GUI

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Optimized version with simplified interface. Displays layers in a TreeView
organized by folder structure with minimal overhead.

Features:
- Fast TreeView with folder structure
- Simple layer selection
- Minimal interface

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad-api package installed (pip install archicad)
- tkinter (included with Python)

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. Select a layer
4. Click "Close" to exit

================================================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox
from archicad import ACConnection


class LayerTreeViewApp:
    """
    Optimized GUI application for viewing Archicad layers
    """

    def __init__(self, root):
        self.root = root
        self.root.title("Layer Selector")
        self.root.geometry("600x500")

        # Connect and load
        try:
            conn = ACConnection.connect()
            acc = conn.commands
            self.folder_structure = acc.GetAttributeFolderStructure('Layer')
        except Exception as e:
            messagebox.showerror("Error", f"Connection failed:\n{str(e)}")
            self.root.quit()
            return

        # Create UI
        self.create_widgets()
        self.populate_tree()

    def create_widgets(self):
        """Create simplified UI"""

        # Title
        tk.Label(
            self.root,
            text="Select a Layer",
            font=("Arial", 12, "bold")
        ).pack(pady=5)

        # TreeView frame
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Scrollbar and TreeView
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(frame, yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.tree.yview)

        # Bind selection
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # Selected layer label
        self.label = tk.Label(
            self.root,
            text="No selection",
            font=("Arial", 10),
            fg="gray"
        )
        self.label.pack(pady=5)

        # Close button
        tk.Button(
            self.root,
            text="Close",
            command=self.root.quit,
            width=10
        ).pack(pady=5)

    def populate_tree(self):
        """Populate tree with optimized recursive function"""
        self.add_folder(self.folder_structure, "")

    def add_folder(self, folder, parent):
        """
        Optimized recursive folder population

        Args:
            folder: AttributeFolderStructure
            parent: Parent tree item ID
        """
        # Create folder node (skip for root)
        if parent or folder.name != "Root":
            folder_id = self.tree.insert(
                parent,
                'end',
                text=f"📁 {folder.name}",
                tags=('folder',),
                open=True
            )
        else:
            folder_id = parent

        # Add layers
        if folder.attributes:
            for attr in folder.attributes:
                self.tree.insert(
                    folder_id,
                    'end',
                    text=f"  {attr.attribute.name}",
                    tags=('layer',),
                    values=(attr.attribute.name,)
                )

        # Add subfolders
        if folder.subfolders:
            for sub in folder.subfolders:
                self.add_folder(sub.attributeFolder, folder_id)

    def on_select(self, event):
        """Handle selection"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        tags = self.tree.item(item, 'tags')

        if 'layer' in tags:
            # Get layer name from values
            values = self.tree.item(item, 'values')
            if values:
                self.label.config(text=f"Selected: {values[0]}", fg="black")
        elif 'folder' in tags:
            self.label.config(text="Folder selected", fg="gray")


def main():
    root = tk.Tk()
    app = LayerTreeViewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
