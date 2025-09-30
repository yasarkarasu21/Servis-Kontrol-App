import tkinter as tk
from tkinter import simpledialog, messagebox
import threading
import time
import json
import os
import win32serviceutil
import ctypes
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import socket

# -------------------------------
# Sunucunun IP adresi
# -------------------------------
def get_ipv4():
    hostname = socket.gethostname()
    ip = socket.gethostbyname(hostname)
    return ip

# -------------------------------
# Yönetici kontrolü
# -------------------------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit()

# -------------------------------
# Mail ayarları
# -------------------------------
SENDER_EMAIL = "mail adresiniz"
SENDER_PASSWORD = "mail adresi şifresi"
SMTP_SERVER = "smtp.yandex.com"
SMTP_PORT = 465

# -------------------------------
# Uygulama şifresi
# -------------------------------
APP_PASSWORD = "159263"

# -------------------------------
# JSON dosyaları
# -------------------------------
SERVICES_FILE = "services.json"
USERS_FILE = "users.json"

# -------------------------------
# Veriler
# -------------------------------
services = []
service_status = {}  # {service_name: True/False}
users = []

# -------------------------------
# Otomatik başlatma durumu
# -------------------------------
auto_start_enabled = False  # Varsayılan: devre dışı

def toggle_auto_start():
    global auto_start_enabled
    auto_start_enabled = not auto_start_enabled
    if auto_start_enabled:
        auto_start_btn.config(text="Otomatik Başlatma: Devrede", bg="green")
    else:
        auto_start_btn.config(text="Otomatik Başlatma: Devre Dışı", bg="red")

# -------------------------------
# JSON yükleme/kaydetme
# -------------------------------
def load_services():
    global services
    if os.path.exists(SERVICES_FILE):
        try:
            with open(SERVICES_FILE, "r") as f:
                services = json.load(f)
        except json.JSONDecodeError:
            services = []
    for s in services:
        service_status[s["name"]] = True

def save_services():
    with open(SERVICES_FILE, "w") as f:
        json.dump(services, f, indent=4)

def load_users():
    global users
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                users = json.load(f)
        except json.JSONDecodeError:
            users = []

def save_users():
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=4)

# -------------------------------
# Mail gönderme
# -------------------------------
def send_alert(server_ip, service_name, auto_start_enabled, auto_start_result):
    if not users:
        print("📭 Mail alıcısı yok.")
        return
    try:
        msg = MIMEMultipart("alternative")
        msg["From"] = SENDER_EMAIL
        msg["To"] = ", ".join(users)
        msg["Subject"] = f"🚨 Servis Durduruldu: {service_name}"
        msg["X-Priority"] = "1"
        msg["Importance"] = "High"

        status_text = "Otomatik başlatma devre dışı." if not auto_start_enabled else f"Otomatik başlatma denendi. Sonuç: {auto_start_result}"
        html = f"""
        <html>
        <head>
          <style>
            body {{ font-family: Arial, sans-serif; background-color: #f9f9f9; color: #333; }}
            .container {{ max-width:600px; margin:20px auto; padding:20px; background:#fff; border:1px solid #ddd; border-radius:8px; }}
            .header {{ background:#c0392b; color:#fff; padding:12px; text-align:center; font-size:18px; font-weight:bold; border-radius:6px 6px 0 0; }}
            .content {{ padding:15px; font-size:15px; line-height:1.6; }}
            .highlight {{ background:#f2d7d5; padding:8px; border-radius:5px; font-weight:bold; }}
            .footer {{ margin-top:20px; font-size:12px; color:#888; text-align:center; }}
          </style>
        </head>
        <body>
          <div class="container">
            <div class="header">🚨 SERVİS UYARISI</div>
            <div class="content">
              <p class="highlight">{server_ip} sunucusunda <b>{service_name}</b> servisi durdu.</p>
              <p>{status_text}</p>
            </div>
            <div class="footer">
              Bu e-posta otomatik olarak gönderilmiştir. Lütfen cevaplamayın.
            </div>
          </div>
        </body>
        </html>
        """
        mime_html = MIMEText(html, "html")
        msg.attach(mime_html)

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, users, msg.as_string())

        print(f"📧 Mail gönderildi → {users}")
    except Exception as e:
        print(f"Mail gönderilemedi: {e}")

# -------------------------------
# Windows servis kontrol
# -------------------------------
def check_windows_service(service_name):
    try:
        status = win32serviceutil.QueryServiceStatus(service_name)[1]
        return status == 4  # 4 = Running
    except Exception as e:
        print(f"{service_name} kontrol hatası:", e)
        return False

def start_service(service_name):
    try:
        win32serviceutil.StartService(service_name)
        print(f"{service_name} başlatıldı ✅")
        service_status[service_name] = True
        update_service_list()
    except Exception as e:
        print(f"{service_name} başlatılamadı:", e)
        service_status[service_name] = False

def stop_service(service_name):
    try:
        win32serviceutil.StopService(service_name)
        print(f"{service_name} durduruldu ❌")
        service_status[service_name] = False
        update_service_list()
    except Exception as e:
        print(f"{service_name} durdurulamadı:", e)

# -------------------------------
# Servis izleme
# -------------------------------
def monitor_services():
    while True:
        for service in services:
            try:
                running = check_windows_service(service["service_name"])
                prev_status = service_status.get(service["name"], True)

                if running:
                    service_status[service["name"]] = True
                else:
                    server_ip = get_ipv4()
                    # Eğer önceki durum True ise veya otomatik başlatma devreye alındı ve servis duruyor
                    if prev_status or (auto_start_enabled and not running):
                        if auto_start_enabled:
                            try:
                                win32serviceutil.StartService(service["service_name"])
                                service_status[service["name"]] = True
                                result = "Otomatik başlatma başarılı oldu"
                            except Exception as e:
                                result = f"Otomatik başlatma başarısız oldu ({e})"
                        else:
                            result = "Otomatik başlatma devre dışı"

                        send_alert(server_ip, service["name"], auto_start_enabled, result)

                    service_status[service["name"]] = False
            except Exception as e:
                print(f"{service['name']} kontrol hatası:", e)
                service_status[service["name"]] = False

        update_service_list()
        time.sleep(5)

# -------------------------------
# Arayüz fonksiyonları
# -------------------------------
def disable_event():
    pw = simpledialog.askstring("Kapatma Şifresi", "Şifreyi girin:", show='*')
    if pw == APP_PASSWORD:
        root.destroy()
    else:
        messagebox.showerror("Hata", "Yanlış şifre!")

def add_service():
    name = simpledialog.askstring("Servis Adı", "Servis adı girin:")
    service_name = simpledialog.askstring("Windows Servis Adı", "Örn: Audiosrv, Spooler")
    interval = simpledialog.askinteger("Kontrol Saniyesi", "Bu servisi kaç saniyede bir kontrol edelim?", minvalue=60)
    if name and service_name and interval:
        services.append({"name": name, "service_name": service_name, "interval": interval})
        service_status[name] = True
        save_services()
        update_service_list()

def remove_service():
    try:
        selected = listbox.curselection()
        if selected:
            index = selected[0]
            service = services.pop(index)
            service_status.pop(service["name"], None)
            save_services()
            update_service_list()
            messagebox.showinfo("Başarılı", f"{service['name']} servisi kaldırıldı.")
        else:
            messagebox.showwarning("Uyarı", "Lütfen listeden bir servis seçin.")
    except Exception as e:
        messagebox.showerror("Hata", f"Servis kaldırılamadı:\n{str(e)}")

def add_user():
    pw = simpledialog.askstring("Şifre Gerekli", "Kullanıcı eklemek için şifreyi girin:", show="*")
    if pw != APP_PASSWORD:
        messagebox.showerror("Hata", "Yanlış şifre!")
        return
    email = simpledialog.askstring("Mail Adresi", "Mail adresi girin:")
    if email and email not in users:
        users.append(email)
        save_users()
        update_user_list()

def remove_user():
    pw = simpledialog.askstring("Şifre Gerekli", "Kullanıcı silmek için şifreyi girin:", show="*")
    if pw != APP_PASSWORD:
        messagebox.showerror("Hata", "Yanlış şifre!")
        return
    selected = user_listbox.curselection()
    if selected:
        index = selected[0]
        users.pop(index)
        save_users()
        update_user_list()

def start_selected_service():
    selected = listbox.curselection()
    if selected:
        index = selected[0]
        service = services[index]
        threading.Thread(target=start_service, args=(service["service_name"],), daemon=True).start()

def stop_selected_service():
    selected = listbox.curselection()
    if selected:
        index = selected[0]
        service = services[index]
        threading.Thread(target=stop_service, args=(service["service_name"],), daemon=True).start()

def refresh_lists():
    update_service_list()
    update_user_list()

# -------------------------------
# Liste güncelleme
# -------------------------------
def update_service_list():
    listbox.delete(0, tk.END)
    for service in services:
        status = service_status.get(service["name"], True)
        text = f"{service['name']} ({service.get('service_name')}) - {service.get('interval',60)}sn"
        listbox.insert(tk.END, text)
        listbox.itemconfig(tk.END, {'fg': 'green' if status else 'red'})

def update_user_list():
    user_listbox.delete(0, tk.END)
    for u in users:
        user_listbox.insert(tk.END, u)

# -------------------------------
# Tkinter Arayüzü
# -------------------------------
root = tk.Tk()
root.title("Windows Servis İzleme Paneli")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)
listbox = tk.Listbox(frame, width=70, height=10, selectmode=tk.SINGLE)
listbox.pack()

btn_frame = tk.Frame(root)
btn_frame.pack(pady=5)
tk.Button(btn_frame, text="Servis Ekle", command=add_service).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Servis Sil", command=remove_service).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Başlat", command=start_selected_service).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Durdur", command=stop_selected_service).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Yenile", command=refresh_lists).pack(side=tk.LEFT, padx=5)

auto_start_btn = tk.Button(
    root,
    text="Otomatik Başlatma: Devre Dışı" if not auto_start_enabled else "Otomatik Başlatma: Devrede",
    bg="red" if not auto_start_enabled else "green",
    fg="white",
    command=toggle_auto_start
)
auto_start_btn.pack(pady=5)

user_frame = tk.Frame(root)
user_frame.pack(padx=10, pady=10)
user_listbox = tk.Listbox(user_frame, width=50, height=5)
user_listbox.pack()
user_btn_frame = tk.Frame(root)
user_btn_frame.pack(pady=5)
tk.Button(user_btn_frame, text="Kullanıcı Ekle", command=add_user).pack(side=tk.LEFT, padx=5)
tk.Button(user_btn_frame, text="Kullanıcı Sil", command=remove_user).pack(side=tk.LEFT, padx=5)

footer = tk.Label(root, text="Kutal Dijital", font=("Arial", 10, "bold"))
footer.pack(pady=10)

# JSON yükleme
load_services()
load_users()
update_service_list()
update_user_list()

# İzleme thread
threading.Thread(target=monitor_services, daemon=True).start()

root.geometry("400x450")
root.resizable(False, False)
root.protocol("WM_DELETE_WINDOW", disable_event)
root.mainloop()
