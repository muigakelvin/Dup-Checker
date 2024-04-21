import os
import tkinter as tk
from tkinter import simpledialog, messagebox
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from ttkbootstrap import Style
from tkinter import ttk
from collections import Counter
import re
import shutil

# Define Google Drive API scope
#SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/drive.file']


# Memoization cache
memo = {}

# Global variable to store searched indices
searched_indices = set()

def authenticate():
    """Authenticate with Google APIs."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Construct the path to credentials.json dynamically
            script_dir = os.path.dirname(os.path.realpath(__file__))
            credentials_path = os.path.join(script_dir, 'credentials4.json')
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def list_folders(service):
    """List available folders in the Google Drive."""
    # Check if folders list is already cached
    if 'folders' in memo:
        return memo['folders']

    results = service.files().list(
        q="mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)").execute()
    folders = results.get('files', [])
    folder_count = len(folders)

    # Cache the folders list
    memo['folders'] = folders
    return memo['folders']

def get_files(service, folder_id, page_size=10, page_token=None):
    """Get files in the specified folder with pagination."""
    # Check if files list for this folder is already cached
    cache_key = folder_id + "_" + str(page_size) + "_" + str(page_token)
    if cache_key in memo:
        return memo[cache_key]

    results = service.files().list(
        q=f"'{folder_id}' in parents",
        fields="nextPageToken, files(id, name, modifiedTime, size)",
        pageSize=page_size,
        pageToken=page_token
    ).execute()

    files = results.get('files', [])
    next_page_token = results.get('nextPageToken')

    print(files)  # Print out files variable to debug

    # Cache the files list for this folder
    memo[cache_key] = (files, next_page_token)
    return files, next_page_token

def extract_index(file_name):
    """Extract index from file name."""
    match = re.search(r'CKS\s*(\d+)\s*(?:\(\d+\))?\s*\.pdf', file_name)
    if match:
        return match.group(1)
    return None

"""def display_files(service, folder_id):
    #Fetch and display files in the selected folder
    page_size = 10
    page_token = None
    total_files = 0
    
    # Initialize index_list
    index_list = []

    while True:
        files, next_page_token = get_files(service, folder_id, page_size, page_token)
        for file in files:
            index = extract_index(file['name'])  # Extract index from the file name
            tree.insert("", "end", text=file['name'], values=(index, file['modifiedTime'], file['size']))
            total_files += 1
            
            # Append index to index_list
            index_list.append(index)
        
        if not next_page_token:
            break  # No more pages to fetch
        else:
            page_token = next_page_token  # Move to the next page
    
    hits_label.config(text="Total Files: " + str(total_files))
    searched_label.config(text="Searched: " + ", ".join(searched_indices))  # Update the searched label
    
    # Return index_list
    return index_list"""
def display_files(service, folder_id):
    """Fetch and display files in the selected folder."""
    page_size = 10
    page_token = None
    total_files = 0
    
    # Initialize index_list
    index_list = []

    while True:
        files, next_page_token = get_files(service, folder_id, page_size, page_token)
        for file in files:
            index = extract_index(file['name'])  # Extract index from the file name
            if index is not None:  # Check if index extraction was successful
                # Check if 'size' key exists in the file dictionary
                file_size = file.get('size', 'Unknown')
                tree.insert("", "end", text=file['name'], values=(index, file['modifiedTime'], file_size))
                total_files += 1
                
                # Append index to index_list
                index_list.append(index)
        
        if not next_page_token:
            break  # No more pages to fetch
        else:
            page_token = next_page_token  # Move to the next page
    
    hits_label.config(text="Total Files: " + str(total_files))
    searched_label.config(text="Searched: " + ", ".join(searched_indices))  # Update the searched label
    
    # Return index_list
    return index_list


def show_duplicates(service, folder_id):
    # Show duplicate files in the selected folder
    files, _ = get_files(service, folder_id)  # Retrieve files list

    # Clear existing items in the treeview
    for item in tree.get_children():
        tree.delete(item)

    # Get the index_list
    index_list = display_files(service, folder_id)

    print("Index List:", index_list)  # Debug statement to print index_list

    # Find duplicate indices
    counter = Counter(index_list)
    duplicate_indices = [index for index, count in counter.items() if count > 1]

    print("Duplicate Indices:", duplicate_indices)  # Debug statement to print duplicate_indices
    
    # Initialize an array to store duplicate file names
    duplicate_files_array = []

    # Iterate through the items in the Treeview to find duplicate files
    for item in tree.get_children():
        values = tree.item(item, 'values')
        file_name = tree.item(item, 'text')
        index = values[0] if values else None
        if index in duplicate_indices:
            duplicate_files_array.append(file_name)

    # Print the array of duplicate file names
    print("Duplicate Files Array:", duplicate_files_array)

    # Calculate total occurrences of duplicates
    total_occurrences = sum(counter[index] for index in duplicate_indices)
    total_files_label.config(text="Total Hits: " + str(total_occurrences))

    # Return the array containing duplicate file names
    return duplicate_files_array


def select_folder(service, folders, selected_folder):
    """Callback for selecting a folder."""
    folder_info = [(folder['id'], folder['name']) for folder in folders if folder['name'] == selected_folder]
    if folder_info:
        folder_id, folder_name = folder_info[0]
        folder_id_label.config(text="Folder ID: " + folder_id)
        folder_name_label.config(text="Folder Name: " + folder_name)
        
        # Clear existing items in the treeview
        for item in tree.get_children():
            tree.delete(item)
        
        # Fetch and display files in the selected folder
        display_files(service, folder_id)
    else:
        messagebox.showerror("Error", "Selected folder information not found.")

import logging

def remove_duplicates(service, folder_id):
    try:
        # Retrieve the list of duplicate file names
        duplicate_files_array = show_duplicates(service, folder_id)

        # Check if the "Duplicates" folder already exists within the selected folder
        folder_exists = False
        files = service.files().list(q=f"'{folder_id}' in parents").execute().get('files', [])
        for file in files:
            if file['name'] == 'Duplicates' and file.get('mimeType') == 'application/vnd.google-apps.folder':
                folder_exists = True
                duplicates_folder_id = file['id']
                break

        if not folder_exists:
            # Create the "Duplicates" folder within the selected folder
            file_metadata = {
                'name': 'Duplicates',
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [folder_id]
            }
            duplicates_folder = service.files().create(body=file_metadata, fields='id').execute()
            duplicates_folder_id = duplicates_folder.get('id')
            print("Created 'Duplicates' folder with ID:", duplicates_folder_id)  # Debugging: Print folder ID

        # Move duplicate files to the "Duplicates" folder
        for file_name in duplicate_files_array:
            # Search for files with the same name in the folder
            query = f"name='{file_name}' and '{folder_id}' in parents"
            print("Query:", query)  # Debugging: Print the query
            files = service.files().list(q=query).execute().get('files', [])
            for file in files:
                try:
                    # Move the file to the "Duplicates" folder and remove original parent
                    service.files().update(
                        fileId=file['id'],
                        addParents=duplicates_folder_id,
                        removeParents=folder_id
                    ).execute()
                    logging.info(f"File '{file['name']}' moved to 'Duplicates' folder successfully.")
                except Exception as e:
                    logging.error(f"Error moving file '{file['name']}' to 'Duplicates' folder: {e}")

        # After moving duplicate files, show a message box indicating success or failure
        messagebox.showinfo("Duplicates Removed", "Duplicate files have been moved to the 'Duplicates' folder.")
    except Exception as e:
        # Display error message if an exception occurs
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def select_google_drive():
    """Select Google Drive folder."""
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)
    folders = list_folders(service)
    if folders:
        selected_folder_var = tk.StringVar()
        selected_folder_combo = ttk.Combobox(root, textvariable=selected_folder_var, values=[folder['name'] for folder in folders], state="readonly",width=50)
        selected_folder_combo.current(0)  # Set initial value
        selected_folder_combo.pack()
        selected_folder_combo.focus()

        confirm_button = ttk.Button(root, text="Select", command=lambda: select_folder(service, folders, selected_folder_var.get()))
        confirm_button.pack()

        show_duplicates_button = ttk.Button(root, text="Show Duplicates", command=lambda: show_duplicates(service, folders[int(selected_folder_combo.current())]['id']))
        show_duplicates_button.pack()
        remove_duplicates_button = ttk.Button(root, text="Remove Duplicates", command=lambda: remove_duplicates(service, folders[int(selected_folder_combo.current())]['id']))
        remove_duplicates_button.pack()

    else:
        messagebox.showerror("Error", "No folders found in Google Drive.")

# Create the main window
root = tk.Tk()
root.title("Google Drive Folder Selection")

# Set ttkbootstrap style
style = Style(theme="flatly")

# Create select folder button
select_folder_button = tk.Button(root, text="Select Google Drive Folder", command=select_google_drive)
select_folder_button.pack()

# Create labels for folder ID, folder name, total files, and searched
folder_id_label = tk.Label(root, text="Folder ID: ")
folder_id_label.pack()

folder_name_label = tk.Label(root, text="Folder Name: ")
folder_name_label.pack()

hits_label = tk.Label(root, text="Total Files: ")
hits_label.pack()

searched_label = tk.Label(root, text="Searched: ")
searched_label.pack()

# Create label for total files
total_files_label = tk.Label(root, text="Total Files: ")
total_files_label.pack()

# Create and configure ttk.Treeview
tree = ttk.Treeview(root, columns=("Index", "Last Modified", "Size"))
tree.heading("#0", text="Name")
tree.heading("Index", text="Index")
tree.heading("Last Modified", text="Last Modified")
tree.heading("Size", text="Size")
tree.column("#0", anchor="w", width=200)
tree.column("Index", anchor="w", width=100)
tree.column("Last Modified", anchor="w", width=200)
tree.column("Size", anchor="w", width=100)
tree.pack(fill="both", expand=True)

# Start the application
root.mainloop()
