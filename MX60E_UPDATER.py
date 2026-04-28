import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import requests
import threading
import random
import time
import os
import io
import xml.etree.ElementTree as ET

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ModernMX60EUpdater(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("MX60E-G Updater")
        self.geometry("1100x750")
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.session = requests.Session()
        self.device_ip = ""
        self.is_running = False
        self.is_logged_in = False

        # --- Sidebar (Connection & Login) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1)

        self.setup_sidebar()

        # --- Main Area (Tabs) ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        self.tab_lines = self.tabview.add("Bulk Line Update")
        self.tab_network = self.tabview.add("Network & SIP")
        self.tab_admin = self.tabview.add("Admin")

        self.setup_lines_tab()
        self.setup_network_tab()
        self.setup_admin_tab()

    def setup_sidebar(self):
        # Title
        ctk.CTkLabel(self.sidebar_frame, text="MX60E-G\nUpdater", font=ctk.CTkFont(size=26, weight="bold")).grid(row=0, column=0, padx=20, pady=(40, 20))

        # Separator
        ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="gray40").grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 20))

        # Connection Panel
        self.conn_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.conn_frame.grid(row=2, column=0, padx=15, sticky="ew")

        ctk.CTkLabel(self.conn_frame, text="DEVICE CONNECTION", font=ctk.CTkFont(size=13, weight="bold"), text_color="gray80").pack(anchor="w", pady=(0, 10))
        
        self.ip_entry = ctk.CTkEntry(self.conn_frame, placeholder_text="IP Address", height=35)
        self.ip_entry.pack(fill="x", pady=5)
        
        self.load_captcha_btn = ctk.CTkButton(self.conn_frame, text="Load Captcha", command=self.load_captcha, fg_color="#3B8ED0", height=35)
        self.load_captcha_btn.pack(fill="x", pady=5)

        self.captcha_label = ctk.CTkLabel(self.conn_frame, text="[Captcha]", height=50, fg_color="gray25", corner_radius=6)
        self.captcha_label.pack(fill="x", pady=10)

        self.pass_entry = ctk.CTkEntry(self.conn_frame, placeholder_text="Password", show="*", height=35)
        self.pass_entry.pack(fill="x", pady=5)

        self.captcha_entry = ctk.CTkEntry(self.conn_frame, placeholder_text="Captcha Code", height=35)
        self.captcha_entry.pack(fill="x", pady=5)

        self.login_btn = ctk.CTkButton(self.conn_frame, text="LOGIN", command=self.login, fg_color="#2CC985", hover_color="#229965", state="disabled", font=ctk.CTkFont(size=14, weight="bold"), height=40)
        self.login_btn.pack(fill="x", pady=(20, 5))

        self.status_label = ctk.CTkLabel(self.conn_frame, text="Disconnected", text_color="gray")
        self.status_label.pack(pady=5)

    def setup_lines_tab(self):
        tab = self.tab_lines
        tab.grid_columnconfigure(0, weight=1)

        # 1. Column Mapping
        ctk.CTkLabel(tab, text="1. Excel Column Mapping", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", pady=(10, 5))
        col_frame = ctk.CTkFrame(tab, fg_color="transparent")
        col_frame.grid(row=1, column=0, sticky="ew")
        
        self.col_ext = self.create_labeled_entry(col_frame, "Extension Column", 0, 0, "# Extension")
        self.col_user = self.create_labeled_entry(col_frame, "User Column", 0, 1, "Username Dispositivo")
        self.col_pass = self.create_labeled_entry(col_frame, "Password Column", 0, 2, "SIP Contrasenna")

        # 2. Global Options
        ctk.CTkLabel(tab, text="2. Global Line Options", font=ctk.CTkFont(size=16, weight="bold")).grid(row=2, column=0, sticky="w", pady=(20, 5))
        opt_frame = ctk.CTkFrame(tab, fg_color="transparent")
        opt_frame.grid(row=3, column=0, sticky="ew")

        self.chk_tls = ctk.CTkCheckBox(opt_frame, text="Enable TLS")
        self.chk_tls.pack(side="left", padx=10)
        
        self.chk_srtp = ctk.CTkCheckBox(opt_frame, text="Enable SRTP")
        self.chk_srtp.pack(side="left", padx=10)

        self.entry_volt = ctk.CTkEntry(opt_frame, placeholder_text="Ringing Voltage", width=150)
        self.entry_volt.pack(side="left", padx=10)

        # 3. File & Start
        ctk.CTkLabel(tab, text="3. Execution", font=ctk.CTkFont(size=16, weight="bold")).grid(row=4, column=0, sticky="w", pady=(20, 5))
        file_frame = ctk.CTkFrame(tab, fg_color="transparent")
        file_frame.grid(row=5, column=0, sticky="ew")
        
        self.file_path = tk.StringVar()
        ctk.CTkButton(file_frame, text="Browse Excel File...", command=self.browse_file).pack(side="left")
        ctk.CTkLabel(file_frame, textvariable=self.file_path, text_color="gray").pack(side="left", padx=10)

        self.start_btn = ctk.CTkButton(tab, text="START BULK UPDATE", command=self.start_bulk_update, fg_color="#2CC985", height=50, font=ctk.CTkFont(size=15, weight="bold"))
        self.start_btn.grid(row=6, column=0, pady=30, sticky="ew")

        # Logs
        self.log_box = ctk.CTkTextbox(tab, height=150)
        self.log_box.grid(row=7, column=0, sticky="nsew")
        tab.grid_rowconfigure(7, weight=1)

    def setup_network_tab(self):
        tab = self.tab_network
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=1)

        # Network
        ctk.CTkLabel(tab, text="Network Settings", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=10)
        
        self.net_ip = self.create_labeled_entry(tab, "IP Address", 1, 0)
        self.net_mask = self.create_labeled_entry(tab, "Mask", 1, 1)
        self.net_gw = self.create_labeled_entry(tab, "Gateway", 2, 0)
        self.net_dns = self.create_labeled_entry(tab, "DNS", 2, 1)

        # SIP
        ctk.CTkLabel(tab, text="SIP Settings", font=ctk.CTkFont(size=16, weight="bold")).grid(row=3, column=0, columnspan=2, sticky="w", pady=(20, 10))
        
        self.sip_proxy = self.create_labeled_entry(tab, "Proxy Server", 4, 0)
        self.sip_sub = self.create_labeled_entry(tab, "Subdomain", 4, 1)
        self.sip_tls = self.create_labeled_entry(tab, "TLS Server", 5, 0)
        
        self.proto_map = {"UDP": "0", "TCP": "1", "TLS": "2"}
        self.sip_proto = self.create_labeled_option(tab, "Protocol", 5, 1, list(self.proto_map.keys()), "TLS")

        self.srtp_map = {
            "RTP only": "0",
            "SRTP only": "1",
            "Both (RTP preferred)": "2",
            "Both (SRTP preferred)": "3",
            "Disable": "4",
            "Mandatory": "5"
        }
        self.sip_srtp_mode = self.create_labeled_option(tab, "SRTP Mode", 6, 0, list(self.srtp_map.keys()), "Mandatory")

        ctk.CTkButton(tab, text="APPLY NETWORK & SIP SETTINGS", command=self.apply_network_settings, fg_color="#2CC985", height=40, font=ctk.CTkFont(weight="bold")).grid(row=7, column=0, columnspan=2, pady=30, sticky="ew")

    def setup_admin_tab(self):
        tab = self.tab_admin
        tab.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(tab, text="Change Admin Password", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", pady=10)
        
        self.old_pass = ctk.CTkEntry(tab, placeholder_text="Old Password")
        self.old_pass.grid(row=1, column=0, pady=5, sticky="ew")
        
        self.new_pass = ctk.CTkEntry(tab, placeholder_text="New Password", show="*")
        self.new_pass.grid(row=2, column=0, pady=5, sticky="ew")
        
        self.conf_pass = ctk.CTkEntry(tab, placeholder_text="Confirm New Password", show="*")
        self.conf_pass.grid(row=3, column=0, pady=5, sticky="ew")

        ctk.CTkButton(tab, text="CHANGE PASSWORD", command=self.change_password, fg_color="#2CC985", height=40).grid(row=4, column=0, pady=20, sticky="ew")

        ctk.CTkLabel(tab, text="System Actions", font=ctk.CTkFont(size=16, weight="bold")).grid(row=5, column=0, sticky="w", pady=(30, 10))
        ctk.CTkButton(tab, text="REBOOT DEVICE", command=self.reboot_device, fg_color="#3B8ED0", height=40).grid(row=6, column=0, pady=10, sticky="ew")

    # --- Helper Methods ---
    def create_labeled_entry(self, parent, text, row, col, default=""):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(size=12)).pack(anchor="w")
        entry = ctk.CTkEntry(frame)
        entry.pack(fill="x")
        if default: entry.insert(0, default)
        return entry

    def create_labeled_option(self, parent, text, row, col, values, default):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=col, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(size=12)).pack(anchor="w")
        option = ctk.CTkOptionMenu(frame, values=values)
        option.set(default)
        option.pack(fill="x")
        return option

    # --- Logic ---
    def load_captcha(self):
        """
        Carga o refresca el captcha del MX60E.
        """
        ip = self.ip_entry.get().strip()
        
        if not ip:
            self.status_label.configure(text="Enter IP first", text_color="red")
            return
        
        self.device_ip = ip

        post_url = f"http://{ip}/xml?method=gw.config.language&id=900"

        try:
            # POST vacío, obligatorio para generar captcha
            post_resp = self.session.post(post_url, timeout=5)
            if post_resp.status_code != 200:
                self.status_label.configure(text=f"POST Failed ({post_resp.status_code})", text_color="red")
                return

            # Generamos un parámetro random para forzar la actualización de la imagen
            tmp_val = str(random.random())
            captcha_url = f"http://{ip}/vcode.bmp?tmp={tmp_val}"

            img_resp = self.session.get(captcha_url, timeout=5)
            if img_resp.status_code == 200:
                image = Image.open(io.BytesIO(img_resp.content))
                image = image.resize((100, 40), Image.Resampling.LANCZOS)

                self.photo = ImageTk.PhotoImage(image)
                self.captcha_label.configure(image=self.photo, text="")
                self.status_label.configure(text="Captcha Loaded", text_color="green")
                self.login_btn.configure(state="normal")
                # Guardamos el tmp_val para login
                self.tmp_val = tmp_val
            else:
                self.status_label.configure(text=f"Captcha fetch failed ({img_resp.status_code})", text_color="red")

        except Exception as e:
            self.status_label.configure(text=f"Error: {e}", text_color="red")



    def login(self):
        """
        Realiza login en el MX60E utilizando contraseña y captcha.
        """
        password = self.pass_entry.get().strip()
        captcha = self.captcha_entry.get().strip()
        
        if not password or not captcha:
            self.status_label.configure(text="Missing Credentials", text_color="red")
            return

        url = f"http://{self.device_ip}/xml"
        
        payload = {
            "method": "gw.account.login",
            "id51": password,
            "id900": captcha,
            "tmp": getattr(self, "tmp_val", str(random.random()))
        }

        try:
            response = self.session.post(url, data=payload, timeout=5)
            if response.status_code == 200 and "error" not in response.text.lower():
                self.is_logged_in = True
                self.status_label.configure(
                    text="LOGGED IN", 
                    text_color="#2CC985", 
                    font=ctk.CTkFont(weight="bold")
                )
                self.login_btn.configure(
                    text="LOGOUT", 
                    command=self.logout, 
                    fg_color="#E53935"
                )
                self.ip_entry.configure(state="disabled")
            else:
                self.status_label.configure(text="Login Failed", text_color="red")
        except Exception as e:
            self.status_label.configure(text=f"Error: {e}", text_color="red")



    def logout(self):
        self.is_logged_in = False
        self.session = requests.Session() # Clear session
        self.status_label.configure(text="Disconnected", text_color="gray")
        self.login_btn.configure(text="LOGIN", command=self.login, fg_color="#2CC985")
        self.ip_entry.configure(state="normal")
        self.captcha_label.configure(image=None, text="[Captcha]")

    def check_login(self):
        if not self.is_logged_in:
            messagebox.showwarning("Not Logged In", "Please login via the sidebar first.")
            return False
        return True

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.file_path.set(filename)

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def start_bulk_update(self):
        if not self.check_login(): return
        if self.is_running: return
        
        file = self.file_path.get().strip()
        if not file:
            messagebox.showerror("Error", "Select a file first.")
            return

        self.is_running = True
        self.start_btn.configure(state="disabled", text="Running...")
        self.log_box.configure(state="normal")
        self.log_box.delete("0.0", "end")
        self.log_box.configure(state="disabled")
        
        thread = threading.Thread(target=self.run_bulk_logic)
        thread.daemon = True
        thread.start()

    def run_bulk_logic(self):
        try:
            df = pd.read_excel(self.file_path.get())
            url = f"http://{self.device_ip}/xml"
            
            c_ext = self.col_ext.get()
            c_user = self.col_user.get()
            c_pass = self.col_pass.get()
            
            tls_val = "1" if self.chk_tls.get() else "0"
            srtp_val = "1" if self.chk_srtp.get() else "0"
            volt_val = self.entry_volt.get().strip()

            total = len(df)
            success = 0
            
            for index, row in df.iterrows():
                line_id = index + 1
                payload = {
                    "method": "gw.config.set",
                    "line_id": str(line_id),
                    "id401": str(row[c_user]),
                    "id455": str(row[c_user]),
                    "id432": str(row[c_pass]),
                    "id912": tls_val,
                    "id913": srtp_val,
                    "id433": 1,
                    "tmp": str(random.random())
                }
                if volt_val: payload["id755"] = volt_val

                try:
                    self.log(f"Line {line_id}: Updating...")
                    resp = self.session.post(url, data=payload, timeout=5)
                    if resp.status_code == 200:
                        self.log(f"  -> Success")
                        success += 1
                    else:
                        self.log(f"  -> Failed: {resp.status_code}")
                except Exception as e:
                    self.log(f"  -> Error: {e}")
                time.sleep(0.1)
            
            messagebox.showinfo("Done", f"Processed {total} lines.\nSuccess: {success}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.is_running = False
            self.start_btn.configure(state="normal", text="START BULK UPDATE")

    def apply_network_settings(self):
        if not self.check_login(): return
        
        payload = {"method": "gw.config.set", "tmp": str(random.random())}
        proto_val = self.proto_map[self.sip_proto.get()]
        srtp_val = self.srtp_map[self.sip_srtp_mode.get()]

        fields = {
            "id9": self.net_ip.get(),
            "id10": self.net_mask.get(),
            "id3": self.net_gw.get(),
            "id7": self.net_dns.get(),
            "id477": self.sip_proxy.get(),
            "id480": self.sip_sub.get(),
            "id911": self.sip_tls.get(),
            "id113": proto_val,
            "id914": srtp_val
        }
        
        for k, v in fields.items():
            if v and v.strip():
                payload[k] = v.strip()
        
        self.send_single_request(payload, "Network Settings Applied")

    def change_password(self):
        if not self.check_login(): return
        old = self.old_pass.get()
        new = self.new_pass.get()
        conf = self.conf_pass.get()
        
        if new != conf:
            messagebox.showerror("Error", "New passwords do not match.")
            return
            
        payload = {
            "method": "gw.account.change",
            "id51": new,
            "tmp": str(random.random())
        }
        self.send_single_request(payload, "Password Changed")

    def reboot_device(self):
        if not self.check_login(): return
        if messagebox.askyesno("Confirm", "Are you sure you want to reboot the device?"):
            payload = {"method": "gw.system.reboot", "tmp": str(random.random())}
            self.send_single_request(payload, "Reboot Command Sent")

    def send_single_request(self, payload, success_msg):
        try:
            url = f"http://{self.device_ip}/xml"
            resp = self.session.post(url, data=payload, timeout=5)
            if resp.status_code == 200:
                messagebox.showinfo("Success", success_msg)
            else:
                messagebox.showerror("Failed", f"Status: {resp.status_code}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = ModernMX60EUpdater()
    app.mainloop()
