import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser
import requests
import csv

# Dependency check and installation
def install_dependencies():
    try:
        import requests
    except ImportError:
        os.system(f"{sys.executable} -m pip install requests")

install_dependencies()

# Main Application Class
class WithSecureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WithSecure API Export Tool")
        self.root.geometry("480x480")
        self.root.resizable(False, False)

        self.has_acknowledged = False
        self.warning_popup_open = False

        # Main frame for centering
        main_frame = tk.Frame(root)
        main_frame.place(relx=0.5, rely=0.5, anchor="center")

        # Title and description
        tk.Label(main_frame, text="WithSecure API Export Tool", font=("Helvetica", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=(10, 5))
        tk.Label(main_frame, text="Retrieve device data via the WithSecure API and export it to a CSV file.").grid(row=1, column=0, columnspan=3, pady=(0, 10))

        # Input fields for API setup
        tk.Label(main_frame, text="Client ID:").grid(row=2, column=0, sticky="w")
        self.client_id = tk.Entry(main_frame, width=40)
        self.client_id.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(main_frame, text="Client Secret:").grid(row=3, column=0, sticky="w")
        self.client_secret = tk.Entry(main_frame, width=40, show="*")
        self.client_secret.grid(row=3, column=1, padx=5, pady=5)
        self.client_secret.bind("<FocusIn>", self.show_security_warning)

        tk.Label(main_frame, text="Export Folder:").grid(row=4, column=0, sticky="w")
        self.export_folder = tk.Entry(main_frame, width=40)
        self.export_folder.grid(row=4, column=1, padx=5, pady=5)

        tk.Button(main_frame, text="Browse", command=self.browse_folder).grid(row=5, column=1, pady=(5, 15))

        # API key explanation and link
        tk.Label(main_frame, text="Generate API Key:").grid(row=6, column=0, columnspan=3, pady=(10, 0))
        tk.Button(main_frame, text="Open API Key Page", command=self.open_api_key_page).grid(row=7, column=0, columnspan=3, pady=(5, 10))

        # Status label and progress bar
        self.status_label = tk.Label(main_frame, text="Status: Waiting for acknowledgment", anchor="center")
        self.status_label.grid(row=8, column=0, columnspan=3, pady=(10, 0))

        self.progress_bar = ttk.Progressbar(main_frame, length=300, mode="determinate")
        self.progress_bar.grid(row=9, column=0, columnspan=3, pady=(5, 10))

        # Start button
        self.export_button = tk.Button(main_frame, text="Export to CSV", command=self.start_export, state="disabled")
        self.export_button.grid(row=10, column=0, columnspan=3, pady=10)

        # Footer with link
        footer = tk.Label(main_frame, text="Made by Clément GHANEME", fg="blue", cursor="hand2")
        footer.grid(row=11, column=0, columnspan=3, pady=10)
        footer.bind("<Button-1>", lambda e: webbrowser.open("https://clement.business/"))

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.export_folder.delete(0, tk.END)
            self.export_folder.insert(0, folder)

    def open_api_key_page(self):
        webbrowser.open("https://elements.withsecure.com/apps/ccr/api_keys")

    def show_security_warning(self, event):
        if not self.has_acknowledged and not self.warning_popup_open:
            self.warning_popup_open = True
            warning_popup = tk.Toplevel(self.root)
            warning_popup.title("Security Warning")
            warning_popup.geometry("400x150")
            warning_popup.resizable(False, False)

            tk.Label(warning_popup, text="To ensure the security of your WithSecure database,", wraplength=380, justify="center").pack(pady=(10, 0))
            tk.Label(warning_popup, text="provide READ ONLY API access when generating your Client Secret.", wraplength=380, justify="center").pack(pady=(0, 10))

            acknowledge_button = tk.Button(warning_popup, text="Acknowledge", state="disabled", command=lambda: self.acknowledge_warning(warning_popup))
            acknowledge_button.pack(pady=10)

            # Enable button after 2 seconds
            self.root.after(2000, lambda: acknowledge_button.config(state="normal"))

    def acknowledge_warning(self, popup):
        self.has_acknowledged = True
        self.warning_popup_open = False
        self.status_label.config(text="Status: Ready")
        self.export_button.config(state="normal")
        popup.destroy()

    def start_export(self):
        if not self.has_acknowledged:
            messagebox.showerror("Acknowledgment Required", "Please acknowledge the security warning before proceeding.")
            return

        client_id = self.client_id.get().strip()
        client_secret = self.client_secret.get().strip()
        export_folder = self.export_folder.get().strip()

        if not client_id or not client_secret or not export_folder:
            messagebox.showerror("Input Error", "Please fill in all fields.")
            return

        self.status_label.config(text="Status: Authenticating...")
        self.root.update_idletasks()

        try:
            # Step 1: Authenticate
            token = self.authenticate(client_id, client_secret)
            self.status_label.config(text="Status: Retrieving organizations...")
            self.root.update_idletasks()

            # Step 2: Get organizations
            organizations = self.get_organizations(token)
            total_orgs = len(organizations)

            # Step 3: Get devices for each organization
            data = []
            for idx, org in enumerate(organizations, start=1):
                org_id = org['id']
                org_name = org['name']
                self.status_label.config(text=f"Status: Processing {org_name} ({idx}/{total_orgs})...")
                self.progress_bar['value'] = (idx / total_orgs) * 100
                self.root.update_idletasks()

                devices = self.get_devices(token, org_id)
                for device in devices:
                    data.append([org_name, device.get("name", "Unknown"), device.get("os", {}).get("name", "Unknown"), device.get("os", {}).get("version", "Unknown")])

            # Step 4: Export to CSV
            self.export_to_csv(data, export_folder)
            self.status_label.config(text="Status: Completed")
            messagebox.showinfo("Success", "Export completed successfully!")
        except Exception as e:
            self.status_label.config(text="Status: Error")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def authenticate(self, client_id, client_secret):
        url = "https://api.connect.withsecure.com/as/token.oauth2"
        payload = {
            "grant_type": "client_credentials",
            "scope": "connect.api.read"
        }
        auth = (client_id, client_secret)
        response = requests.post(url, data=payload, auth=auth)

        if response.status_code == 200:
            return response.json()["access_token"]
        else:
            raise Exception("Authentication failed.")

    def get_organizations(self, token):
        url = "https://api.connect.withsecure.com/organizations/v1/organizations"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            response.encoding = 'utf-8'  # Ensure response is treated as UTF-8
            return response.json()["items"]
        else:
            raise Exception("Failed to retrieve organizations.")

    def get_devices(self, token, organization_id):
        url = f"https://api.connect.withsecure.com/devices/v1/devices?organizationId={organization_id}"
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            response.encoding = 'utf-8'  # Ensure response is treated as UTF-8
            return response.json()["items"]
        else:
            raise Exception(f"Failed to retrieve devices for organization {organization_id}.")

    def export_to_csv(self, data, export_folder):
        output_path = os.path.join(export_folder, "withsecure_export.csv")
        with open(output_path, mode="w", newline="", encoding="utf-8") as file:
            file.write('\ufeff')  # Add BOM to ensure Excel opens the file as UTF-8
            writer = csv.writer(file)
            writer.writerow(["Organization", "Device Name", "OS Name", "OS Version"])
            writer.writerows(data)


# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = WithSecureApp(root)
    root.mainloop()
