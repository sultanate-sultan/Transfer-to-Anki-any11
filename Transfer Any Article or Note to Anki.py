import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import requests
import json
import pyperclip
import os
import pickle
import base64
import io
import re
import html
from datetime import datetime
from PIL import Image, ImageGrab
import tempfile
import win32clipboard
from io import BytesIO

class AnkiDeckManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Anki Deck Manager & Note Creator")
        self.root.geometry("600x500")
        self.root.configure(bg='#f0f0f0')
        
        # File to store deck data
        self.deck_data_file = "anki_decks.pkl"
        self.saved_decks = self.load_deck_data()
        
        # AnkiConnect settings
        self.anki_url = "http://localhost:8765"
        
        # Track selected deck
        self.selected_deck = None
        
        self.create_widgets()
        self.refresh_deck_list()
        
    def load_deck_data(self):
        """Load saved deck data from file"""
        try:
            if os.path.exists(self.deck_data_file):
                with open(self.deck_data_file, 'rb') as f:
                    return pickle.load(f)
            return {}
        except Exception as e:
            print(f"Error loading deck data: {e}")
            return {}
    
    def save_deck_data(self):
        """Save deck data to file"""
        try:
            with open(self.deck_data_file, 'wb') as f:
                pickle.dump(self.saved_decks, f)
        except Exception as e:
            print(f"Error saving deck data: {e}")
    
    def create_widgets(self):
        """Create the main UI"""
        # Title
        title_label = tk.Label(self.root, text="Anki Deck Manager & Note Creator", 
                              font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack(pady=10)
        
        # Deck management frame
        deck_frame = tk.LabelFrame(self.root, text="Deck Management", 
                                  font=('Arial', 12, 'bold'), bg='#f0f0f0')
        deck_frame.pack(padx=20, pady=10, fill='both', expand=True)
        
        # Deck list
        self.deck_listbox = tk.Listbox(deck_frame, height=8, font=('Arial', 10))
        self.deck_listbox.pack(padx=10, pady=10, fill='both', expand=True)
        
        # Deck management buttons
        button_frame = tk.Frame(deck_frame, bg='#f0f0f0')
        button_frame.pack(padx=10, pady=5, fill='x')
        
        self.add_deck_btn = tk.Button(button_frame, text="Add Deck", 
                                     command=self.add_deck, bg='#4CAF50', fg='white')
        self.add_deck_btn.pack(side='left', padx=5)
        
        self.edit_deck_btn = tk.Button(button_frame, text="Edit Deck", 
                                      command=self.edit_deck, bg='#FF9800', fg='white')
        self.edit_deck_btn.pack(side='left', padx=5)
        
        self.delete_deck_btn = tk.Button(button_frame, text="Delete Deck", 
                                        command=self.delete_deck, bg='#f44336', fg='white')
        self.delete_deck_btn.pack(side='left', padx=5)
        
        self.refresh_btn = tk.Button(button_frame, text="Refresh from Anki", 
                                    command=self.refresh_from_anki, bg='#2196F3', fg='white')
        self.refresh_btn.pack(side='right', padx=5)
        
        # Note creation frame
        note_frame = tk.LabelFrame(self.root, text="Create Note from Clipboard", 
                                  font=('Arial', 12, 'bold'), bg='#f0f0f0')
        note_frame.pack(padx=20, pady=10, fill='x')
        
        # Selected deck display
        selected_frame = tk.Frame(note_frame, bg='#f0f0f0')
        selected_frame.pack(padx=10, pady=5, fill='x')
        
        tk.Label(selected_frame, text="Selected Deck:", 
                font=('Arial', 10, 'bold'), bg='#f0f0f0').pack(side='left')
        
        self.selected_deck_label = tk.Label(selected_frame, text="None", 
                                           font=('Arial', 10), bg='#f0f0f0', fg='red')
        self.selected_deck_label.pack(side='left', padx=5)
        
        # Create note button (single unified button)
        self.create_note_btn = tk.Button(note_frame, text="Create Note from Clipboard", 
                                        command=self.create_note_from_clipboard, 
                                        bg='#9C27B0', fg='white', font=('Arial', 12, 'bold'))
        self.create_note_btn.pack(padx=10, pady=10)
        
        # Status label
        self.status_label = tk.Label(self.root, text="Ready", 
                                    font=('Arial', 10), bg='#f0f0f0')
        self.status_label.pack(pady=5)
        
        # Bind listbox selection - using both events for better reliability
        self.deck_listbox.bind('<<ListboxSelection>>', self.on_deck_select)
        self.deck_listbox.bind('<Button-1>', self.on_deck_click)
    
    def anki_request(self, action, **params):
        """Send request to AnkiConnect"""
        payload = {
            "action": action,
            "version": 6,
            "params": params
        }
        
        try:
            response = requests.post(self.anki_url, json=payload, timeout=10)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            self.status_label.config(text=f"Error connecting to Anki: {e}", fg='red')
            return None
    
    def check_anki_connection(self):
        """Check if AnkiConnect is available"""
        result = self.anki_request("version")
        if result and result.get('result'):
            return True
        messagebox.showerror("Connection Error", 
                           "Cannot connect to Anki. Make sure:\n"
                           "1. Anki is running\n"
                           "2. AnkiConnect add-on is installed\n"
                           "3. AnkiConnect is enabled")
        return False
    
    def get_anki_decks(self):
        """Get all decks from Anki"""
        if not self.check_anki_connection():
            return []
        
        result = self.anki_request("deckNames")
        if result and result.get('result'):
            return result['result']
        return []
    
    def refresh_from_anki(self):
        """Refresh deck list from Anki"""
        anki_decks = self.get_anki_decks()
        if anki_decks:
            # Add new decks from Anki to saved decks
            for deck in anki_decks:
                if deck not in self.saved_decks:
                    self.saved_decks[deck] = {
                        'name': deck,
                        'created': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'source': 'anki'
                    }
            
            self.save_deck_data()
            self.refresh_deck_list()
            self.status_label.config(text="Refreshed from Anki", fg='green')
        else:
            self.status_label.config(text="Failed to refresh from Anki", fg='red')
    
    def refresh_deck_list(self):
        """Refresh the deck listbox"""
        self.deck_listbox.delete(0, tk.END)
        for deck_name in sorted(self.saved_decks.keys()):
            self.deck_listbox.insert(tk.END, deck_name)
        
        # Clear selection when refreshing
        self.selected_deck = None
        self.update_selected_deck_display()
    
    def add_deck(self):
        """Add a new deck"""
        deck_name = simpledialog.askstring("Add Deck", "Enter deck name:")
        if deck_name and deck_name.strip():
            deck_name = deck_name.strip()
            
            if deck_name in self.saved_decks:
                messagebox.showwarning("Duplicate Deck", "Deck already exists!")
                return
            
            # Try to create deck in Anki
            if self.check_anki_connection():
                result = self.anki_request("createDeck", deck=deck_name)
                if result and result.get('error') is None:
                    source = 'created'
                else:
                    messagebox.showwarning("Warning", 
                                         f"Deck saved locally but couldn't create in Anki: {result.get('error', 'Unknown error')}")
                    source = 'local'
            else:
                source = 'local'
            
            self.saved_decks[deck_name] = {
                'name': deck_name,
                'created': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'source': source
            }
            
            self.save_deck_data()
            self.refresh_deck_list()
            self.status_label.config(text=f"Deck '{deck_name}' added", fg='green')
    
    def edit_deck(self):
        """Edit selected deck"""
        selection = self.deck_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a deck to edit")
            return
        
        old_name = self.deck_listbox.get(selection[0])
        new_name = simpledialog.askstring("Edit Deck", f"Edit deck name:", initialvalue=old_name)
        
        if new_name and new_name.strip() and new_name != old_name:
            new_name = new_name.strip()
            
            if new_name in self.saved_decks:
                messagebox.showwarning("Duplicate Deck", "Deck name already exists!")
                return
            
            # Update saved decks
            self.saved_decks[new_name] = self.saved_decks[old_name]
            self.saved_decks[new_name]['name'] = new_name
            del self.saved_decks[old_name]
            
            # Update selected deck if it was the one being edited
            if self.selected_deck == old_name:
                self.selected_deck = new_name
            
            self.save_deck_data()
            self.refresh_deck_list()
            self.update_selected_deck_display()
            self.status_label.config(text=f"Deck renamed to '{new_name}'", fg='green')
    
    def delete_deck(self):
        """Delete selected deck"""
        selection = self.deck_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a deck to delete")
            return
        
        deck_name = self.deck_listbox.get(selection[0])
        
        if messagebox.askyesno("Confirm Delete", 
                              f"Are you sure you want to delete '{deck_name}' from the list?\n"
                              f"(This won't delete the deck from Anki)"):
            del self.saved_decks[deck_name]
            
            # Clear selection if deleted deck was selected
            if self.selected_deck == deck_name:
                self.selected_deck = None
                
            self.save_deck_data()
            self.refresh_deck_list()
            self.update_selected_deck_display()
            self.status_label.config(text=f"Deck '{deck_name}' removed from list", fg='green')
    
    def on_deck_click(self, event):
        """Handle deck click (for immediate selection)"""
        # Use after_idle to ensure the selection has been processed
        self.root.after_idle(self.update_deck_selection)
    
    def on_deck_select(self, event):
        """Handle deck selection"""
        self.update_deck_selection()
    
    def update_deck_selection(self):
        """Update the selected deck based on current listbox selection"""
        selection = self.deck_listbox.curselection()
        if selection:
            deck_name = self.deck_listbox.get(selection[0])
            self.selected_deck = deck_name
        else:
            self.selected_deck = None
        
        self.update_selected_deck_display()
    
    def update_selected_deck_display(self):
        """Update the selected deck label display"""
        if self.selected_deck:
            self.selected_deck_label.config(text=self.selected_deck, fg='green')
        else:
            self.selected_deck_label.config(text="None", fg='red')
    
    def get_clipboard_html(self):
        """Get HTML content from clipboard if available"""
        try:
            win32clipboard.OpenClipboard()
            try:
                # Try to get HTML format
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.RegisterClipboardFormat("HTML Format")):
                    html_data = win32clipboard.GetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"))
                    if html_data:
                        # Parse HTML from clipboard
                        html_str = html_data.decode('utf-8', errors='ignore')
                        # Extract the HTML content between StartHTML and EndHTML
                        start_match = re.search(r'StartHTML:(\d+)', html_str)
                        end_match = re.search(r'EndHTML:(\d+)', html_str)
                        if start_match and end_match:
                            start_pos = int(start_match.group(1))
                            end_pos = int(end_match.group(1))
                            html_content = html_str[start_pos:end_pos]
                            return html_content
                        return html_str
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"Error getting HTML from clipboard: {e}")
        return None
    
    def get_clipboard_rtf(self):
        """Get RTF content from clipboard if available"""
        try:
            win32clipboard.OpenClipboard()
            try:
                # Try to get RTF format
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.RegisterClipboardFormat("Rich Text Format")):
                    rtf_data = win32clipboard.GetClipboardData(win32clipboard.RegisterClipboardFormat("Rich Text Format"))
                    if rtf_data:
                        return rtf_data.decode('utf-8', errors='ignore')
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"Error getting RTF from clipboard: {e}")
        return None
    
    def extract_images_from_html(self, html_content):
        """Extract and process images from HTML content"""
        if not html_content:
            return html_content, []
        
        images_stored = []
        
        # Find all img tags with base64 data
        img_pattern = re.compile(r'<img[^>]*src="data:image/([^;]+);base64,([^"]+)"[^>]*>', re.IGNORECASE)
        
        def replace_img(match):
            img_format = match.group(1)
            img_data = match.group(2)
            
            try:
                # Decode base64 image
                image_bytes = base64.b64decode(img_data)
                
                # Create PIL image
                image = Image.open(BytesIO(image_bytes))
                
                # Scale image to 125% of original size
                original_width, original_height = image.size
                new_width = int(original_width * 1.25)
                new_height = int(original_height * 1.25)
                image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                # Generate filename
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                filename = f"clipboard_img_{timestamp}.{img_format.lower()}"
                
                # Convert image back to base64 for Anki storage
                buffer = BytesIO()
                if image.mode in ('RGBA', 'LA', 'P'):
                    image = image.convert('RGB')
                image.save(buffer, format='PNG')
                img_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
                
                # Store in Anki media
                media_result = self.anki_request("storeMediaFile", filename=filename, data=img_base64)
                
                if media_result and not media_result.get('error'):
                    images_stored.append(filename)
                    original_tag = match.group(0)
                    style_match = re.search(r'style="([^"]*)"', original_tag)
                    if style_match:
                        style = style_match.group(1)
                        if "max-width" not in style.lower():
                            style += "; max-width: 140%; height: auto;"
                        else:
                            style = re.sub(r'max-width:\s*[^;]+', 'max-width: 140%', style)
                    else:
                        style = "max-width: 140%; height: auto;"
                    
                    return f'<img src="{filename}" style="{style}">'
                else:
                    print(f"Failed to store image: {media_result.get('error') if media_result else 'Unknown error'}")
                    return match.group(0)
                    
            except Exception as e:
                print(f"Error processing embedded image: {e}")
                return match.group(0)
        
        processed_html = img_pattern.sub(replace_img, html_content)
        return processed_html, images_stored
    
    def rtf_to_html(self, rtf_content):
        """Convert RTF to HTML (basic conversion)"""
        if not rtf_content:
            return ""
        html_content = rtf_content
        html_content = re.sub(r'\\rtf1\\ansi.*?\\f0\\fs\d+\\lang\d+', '', html_content)
        html_content = re.sub(r'\\b\s*([^\\}]+?)\\b0', r'<b>\1</b>', html_content)
        html_content = re.sub(r'\\i\s*([^\\}]+?)\\i0', r'<i>\1</i>', html_content)
        html_content = re.sub(r'\\ul\s*([^\\}]+?)\\ulnone', r'<u>\1</u>', html_content)
        html_content = re.sub(r'\\par\s*', '<br>', html_content)
        html_content = re.sub(r'\\[a-z]+\d*\s*', '', html_content)
        html_content = re.sub(r'[{}]', '', html_content)
        return html_content.strip()
    
    def preserve_formatting(self, content):
        """Enhanced formatting preservation with layered styling."""
        if not content:
            return ""
        
        html_content = self.get_clipboard_html()
        rtf_content = self.get_clipboard_rtf()
        
        if html_content:
            print("DEBUG: Found HTML content in clipboard")
            processed_html, stored_images = self.extract_images_from_html(html_content)
            if stored_images:
                print(f"DEBUG: Stored {len(stored_images)} images from HTML")
            
            # Styling pipeline
            # 1. Apply styles directly to semantic tags (b, strong, i, em).
            styled_html = self.apply_styles_to_semantic_tags(processed_html)
            # 2. Incrementally apply styles from inline CSS to spans without cascading.
            final_styled_html = self.apply_styles_incrementally(styled_html)
            
            return self.clean_html(final_styled_html)
        
        elif rtf_content:
            print("DEBUG: Found RTF content in clipboard")
            return self.rtf_to_html(rtf_content)
        
        else:
            print("DEBUG: Using plain text with enhanced formatting")
            return self.convert_to_html(content)

    def apply_styles_to_semantic_tags(self, html_content):
        """Injects custom CSS styles directly into <b>, <strong>, <i>, and <em> tags."""
        if not html_content:
            return ""

        def style_injector(match):
            tag_name_full = match.group(1)
            attributes = match.group(2)
            tag_name_lower = tag_name_full.lower()
            
            style_to_add = ""
            if tag_name_lower in ['b', 'strong']:
                style_to_add = "color: #facc15; font-weight: 600;"
            elif tag_name_lower in ['i', 'em']:
                style_to_add = "color: #4ade80; font-style: italic;"
            
            if not style_to_add:
                return match.group(0)

            style_match = re.search(r'style="([^"]*)"', attributes, re.IGNORECASE)
            
            if style_match:
                existing_styles = style_match.group(1).rstrip('; ')
                new_style_attr = f'style="{existing_styles}; {style_to_add}"'
                updated_attributes = attributes.replace(style_match.group(0), new_style_attr)
            else:
                updated_attributes = f'{attributes} style="{style_to_add}"'
            
            return f'<{tag_name_full}{updated_attributes}>'

        tag_pattern = re.compile(r'<((?:b|strong|i|em))([^>]*)>', re.IGNORECASE)
        return tag_pattern.sub(style_injector, html_content)

    def apply_styles_incrementally(self, html_content):
        """
        Applies styles based on inline CSS (e.g., font-weight) only to SPAN tags,
        avoiding inheritance issues with block-level elements.
        """
        if not html_content:
            return ""

        # Pattern to find all opening tags and their attributes
        tag_pattern = re.compile(r'(<([a-zA-Z0-9]+)\s*[^>]*>)')
        
        def style_enhancer(match):
            full_tag = match.group(1)
            tag_name = match.group(2).lower()
            
            # <<< KEY CHANGE: Only apply this logic to SPAN tags >>>
            # This prevents styling entire blocks like <p> or <div>.
            if tag_name != 'span':
                return full_tag

            # Find the style attribute within the tag
            style_attr_match = re.search(r'style=(["\'])(.*?)\1', full_tag, re.IGNORECASE | re.DOTALL)
            if not style_attr_match:
                return full_tag

            original_styles = style_attr_match.group(2)
            
            # Check if a color is already explicitly defined
            if 'color:' in original_styles.lower():
                return full_tag

            style_to_add = ""
            # Check for bold and italic indicators
            is_bold = re.search(r'font-weight:\s*(bold|[7-9]00)\b', original_styles, re.IGNORECASE)
            is_italic = re.search(r'font-style:\s*italic\b', original_styles, re.IGNORECASE)

            # Apply color based on weight/style, prioritizing bold
            if is_bold:
                style_to_add = "color: #facc15;"
            elif is_italic:
                style_to_add = "color: #4ade80;"

            if style_to_add:
                # Append the new color to the existing styles
                new_styles = original_styles.rstrip('; ') + '; ' + style_to_add
                # Replace the old style string with the new, enhanced one
                return full_tag.replace(original_styles, new_styles)
            
            return full_tag

        # Apply the enhancer function to all found tags in the HTML
        return tag_pattern.sub(style_enhancer, html_content)

    def clean_html(self, html_content):
        """Clean and optimize HTML for Anki"""
        if not html_content:
            return ""
        
        html_content = re.sub(r'<meta[^>]*>', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</?html[^>]*>', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</?head[^>]*>', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</?body[^>]*>', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'<title[^>]*>.*?</title>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        html_content = re.sub(r'\n\s*\n', '\n', html_content)
        html_content = html_content.strip()
        
        return html_content
    
    def create_note_from_clipboard(self):
        """Create note from clipboard content with full formatting and image preservation"""
        if not self.selected_deck:
            messagebox.showwarning("No Deck Selected", "Please select a deck first")
            return
        
        if not self.check_anki_connection():
            return
        
        try:
            clipboard_content = ""
            try:
                clipboard_content = pyperclip.paste()
            except:
                pass
            
            standalone_image = None
            try:
                standalone_image = ImageGrab.grabclipboard()
            except:
                pass
            
            if not clipboard_content.strip() and not standalone_image:
                messagebox.showwarning("Empty Clipboard", "Clipboard is empty")
                return
            
            if standalone_image and not clipboard_content.strip():
                print("DEBUG: Processing standalone image")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"clipboard_image_{timestamp}.png"
                
                original_width, original_height = standalone_image.size
                new_width = int(original_width * 1.25)
                new_height = int(original_height * 1.25)
                scaled_image = standalone_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                buffer = BytesIO()
                if scaled_image.mode in ('RGBA', 'LA', 'P'):
                    scaled_image = scaled_image.convert('RGB')
                scaled_image.save(buffer, format='PNG')
                img_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
                
                media_result = self.anki_request("storeMediaFile", filename=filename, data=img_base64)
                
                if media_result and not media_result.get('error'):
                    front_content = f'<img src="{filename}" style="max-width: 140%; height: auto;">'
                    image_info = f" (Image scaled from {original_width}x{original_height} to {new_width}x{new_height})"
                else:
                    messagebox.showerror("Error", "Failed to store image in Anki")
                    return
            else:
                print("DEBUG: Processing rich content")
                front_content = self.preserve_formatting(clipboard_content)
                image_info = ""
            
            note_data = {
                "deckName": self.selected_deck,
                "modelName": "Basic",
                "fields": {
                    "Front": front_content,
                    "Back": ""
                },
                "tags": ["clipboard-import", "front-only"]
            }
            
            result = self.anki_request("addNote", note=note_data)
            
            if result and result.get('error') is None:
                note_id = result.get('result')
                self.status_label.config(text=f"Note created successfully (ID: {note_id})", fg='green')
                messagebox.showinfo("Success", f"Note added to deck '{self.selected_deck}'!{image_info}")
            else:
                error_msg = result.get('error', 'Unknown error') if result else 'Connection failed'
                self.status_label.config(text=f"Failed to create note: {error_msg}", fg='red')
                messagebox.showerror("Error", f"Failed to create note: {error_msg}")
                
        except Exception as e:
            self.status_label.config(text=f"Error: {e}", fg='red')
            messagebox.showerror("Error", f"An error occurred: {e}")
    
    def convert_to_html(self, text):
        """Enhanced text to HTML conversion with custom styling."""
        if not text:
            return ""

        h_style = "font-weight: 700; line-height: 1.3; margin-top: 1.5em; margin-bottom: 0.5em;"
        h1_style = f"color: #60a5fa; {h_style}"
        h2_style = f"color: #a78bfa; border-bottom: 2px solid #374151; padding-bottom: 0.3em; {h_style}"
        h3_style = f"color: #f472b6; {h_style}"
        strong_style = "color: #facc15; font-weight: 600;"
        em_style = "color: #4ade80; font-style: italic;"
        
        html_content = html.escape(text)
        lines = html_content.split('\n')
        processed_lines = []
        
        for line in lines:
            leading_spaces = len(line) - len(line.lstrip())
            if leading_spaces > 0:
                indent = '&nbsp;' * leading_spaces
                line = indent + line.lstrip()
            
            if line.strip().startswith('### '):
                line = f'<h3 style="{h3_style}">{line.strip()[4:]}</h3>'
            elif line.strip().startswith('## '):
                line = f'<h2 style="{h2_style}">{line.strip()[3:]}</h2>'
            elif line.strip().startswith('# '):
                line = f'<h1 style="{h1_style}">{line.strip()[2:]}</h1>'
            elif re.match(r'^(&nbsp;)*[-*+]\s+', line):
                indent_level = line.count('&nbsp;') // 4
                content = re.sub(r'^(&nbsp;)*[-*+]\s+', '', line)
                line = f'<div style="margin-left: {indent_level * 20}px;">• {content}</div>'
            elif re.match(r'^(&nbsp;)*\d+\.\s+', line):
                indent_level = line.count('&nbsp;') // 4
                content = re.sub(r'^(&nbsp;)*\d+\.\s+', '', line)
                line = f'<div style="margin-left: {indent_level * 20}px;">1. {content}</div>'
            else:
                line = re.sub(r'\*\*(.*?)\*\*', rf'<strong style="{strong_style}">\1</strong>', line)
                line = re.sub(r'__(.*?)__', rf'<strong style="{strong_style}">\1</strong>', line)
                line = re.sub(r'\*(.*?)\*', rf'<em style="{em_style}">\1</em>', line)
                line = re.sub(r'_(.*?)_', rf'<em style="{em_style}">\1</em>', line)
                line = re.sub(r'<u>(.*?)</u>', r'<u>\1</u>', line)
                line = re.sub(r'`(.*?)`', r'<code>\1</code>', line)
                line = re.sub(r'~~(.*?)~~', r'<s>\1</s>', line)
            
            processed_lines.append(line)
        
        html_content = '<br>'.join(processed_lines)
        html_content = re.sub(r'```(.*?)```', r'<pre><code>\1</code></pre>', html_content, flags=re.DOTALL)
        html_content = re.sub(r'^&gt;\s*(.*?)$', r'<blockquote>\1</blockquote>', html_content, flags=re.MULTILINE)
        
        return html_content
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.root.quit()

if __name__ == "__main__":
    try:
        import pyperclip
        import requests
        from PIL import Image, ImageGrab
        import win32clipboard
    except ImportError as e:
        print(f"Missing required module: {e}")
        print("Please install required modules:")
        print("pip install pyperclip requests pillow pywin32")
        exit(1)
    
    app = AnkiDeckManager()
    app.run()
