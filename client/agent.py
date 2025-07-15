import requests
import json
import os
import time
import subprocess
import logging
import sys

# Logging setup: log ทุกอย่างลง agent.log
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    handlers=[
        logging.FileHandler("agent.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
# redirect print, error ไป log
class LoggerWriter:
    def __init__(self, level):
        self.level = level
    def write(self, message):
        message = message.rstrip()
        if message:
            self.level(message)
    def flush(self):
        pass
sys.stdout = LoggerWriter(logging.info)
sys.stderr = LoggerWriter(logging.error)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'agent_config.json')
import platform

def ensure_agent_config():
    # โหลด config ถ้ามี
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r') as f:
            config = json.load(f)
        # ถ้ามี agent_api_key และ machine_name แล้ว ให้ใช้เลย
        if config.get('agent_api_key') and config.get('machine_name'):
            return config
    else:
        config = {}
    # ถ้ายังไม่มี agent_api_key หรือ machine_name ให้ register ใหม่
    api_url = config.get('api_url') or 'http://127.0.0.1:8000'  # fallback ถ้าไม่ได้ตั้ง
    set_id = config.get('set_id')
    if not set_id:
        raise Exception('agent_config.json ต้องมี set_id')
    machine_name = platform.node()
    resp = requests.post(f'{api_url}/machine/register', json={
        'set_id': set_id,
        'machine_name': machine_name
    })
    if resp.status_code != 200:
        raise Exception(f'Register agent failed: {resp.text}')
    data = resp.json()
    config['api_url'] = api_url
    config['set_id'] = set_id
    config['machine_name'] = machine_name
    config['agent_api_key'] = data['agent_api_key']
    with open(CONFIG_PATH, 'w') as f:
        json.dump(config, f, indent=2)
    return config

config = ensure_agent_config()
API_URL = config['api_url']
MACHINE_NAME = config['machine_name']
AGENT_API_KEY = config['agent_api_key']
MACHINE_ID = None

# ------------------- Agent Utility Functions -------------------
def get_machine_id():
    url = f"{API_URL}/machine/config?name={MACHINE_NAME}"
    headers = {"X-AGENT-KEY": AGENT_API_KEY}
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return resp.json()['machine_id']
    print('Get machine_id failed:', resp.text)
    return None

def poll_pending_commands(machine_id):
    url = f"{API_URL}/machine/command/pending?machine_id={machine_id}"
    headers = {"X-AGENT-KEY": AGENT_API_KEY}
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return resp.json()
    return []

def report_command_result(command_id, status, result=None, machine_id=None):
    # ต้องส่ง machine_id เป็น query param ด้วย
    if machine_id is None:
        print("Warning: report_command_result called without machine_id!")
        return
    url = f"{API_URL}/machine/command/result?machine_id={machine_id}"
    headers = {"X-AGENT-KEY": AGENT_API_KEY}
    data = {"command_id": command_id, "status": status, "result": result}
    requests.post(url, json=data, headers=headers)

def report_remote(machine_id, anydesk_id, rustdesk_id):
    # ต้องส่ง machine_id เป็น query param ด้วย
    url = f"{API_URL}/machine/report_remote?machine_id={machine_id}"
    headers = {"X-AGENT-KEY": AGENT_API_KEY}
    data = {"machine_id": machine_id, "anydesk_id": anydesk_id, "rustdesk_id": rustdesk_id}
    requests.post(url, json=data, headers=headers)

# ------------------- Remote Tool Automation -------------------
def install_anydesk():
    import shutil
    import os
    import subprocess
    anydesk_installed_path = "C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe"
    # ถ้ามี AnyDesk อยู่แล้ว ไม่ต้องติดตั้งใหม่
    if os.path.exists(anydesk_installed_path):
        print("AnyDesk already installed, skip install.")
        return anydesk_installed_path
    url = "https://download.anydesk.com/AnyDesk.exe"
    exe_path = "AnyDesk.exe"
    # ลบไฟล์เดิมถ้ามี
    if os.path.exists(exe_path):
        try:
            os.remove(exe_path)
            print(f"Removed old {exe_path}")
        except Exception as e:
            raise Exception(f"Cannot remove old {exe_path}: {e}")
    # ดาวน์โหลดใหม่แบบ stream
    try:
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            with open(exe_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
    except Exception as e:
        raise Exception(f"Download failed: {e}")
    # ตรวจสอบขนาดไฟล์
    try:
        size = os.path.getsize(exe_path)
        print("Downloaded AnyDesk size:", size)
        if size < 2000000:
            raise Exception("AnyDesk.exe download corrupted or incomplete!")
    except Exception as e:
        raise Exception(f"File check failed: {e}")
    # ติดตั้ง AnyDesk
    try:
        subprocess.run([exe_path, "--install", "C:\\Program Files (x86)\\AnyDesk", "--start-with-win", "--silent"], check=True)
    except subprocess.CalledProcessError as e:
        raise Exception(f"AnyDesk install error: {e}")
    except Exception as e:
        raise Exception(f"Unknown error during AnyDesk install: {e}")
    return anydesk_installed_path

def set_anydesk_password(password="123456"):
    import winreg
    import time
    import subprocess
    import os
    # พยายามรัน AnyDesk.exe แบบ background (ถ้ายังไม่รัน)
    anydesk_path = os.path.join("C:\\Program Files (x86)\\AnyDesk", "AnyDesk.exe")
    try:
        subprocess.Popen([anydesk_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("Started AnyDesk.exe to create registry...")
    except Exception as e:
        print("Warning: Unable to start AnyDesk.exe:", e)
    reg_paths = [
        r"SOFTWARE\\AnyDesk",
        r"SOFTWARE\\WOW6432Node\\AnyDesk"
    ]
    key = None
    for _ in range(10):  # ลองเช็ค 10 รอบ (รอ AnyDesk สร้าง registry)
        for reg_path in reg_paths:
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    reg_path,
                    0,
                    winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
                )
                print(f"Found registry: {reg_path}")
                break
            except FileNotFoundError:
                continue
        if key:
            break
        print("Waiting for AnyDesk registry to be created...")
        time.sleep(2)
    if not key:
        print("Set AnyDesk password failed: Registry key not found. Please open AnyDesk once to initialize.")
        return False
    try:
        winreg.SetValueEx(key, "Password", 0, winreg.REG_SZ, password)
        winreg.CloseKey(key)
        print("Set AnyDesk password success.")
        return True
    except Exception as e:
        print("Set AnyDesk password failed:", e)
        return False

def get_anydesk_id():
    import os
    import time
    conf_path = r"C:\ProgramData\AnyDesk\service.conf"
    # รอไฟล์ service.conf ปรากฏ (สูงสุด 10 รอบ)
    for _ in range(10):
        if os.path.exists(conf_path):
            break
        print("Waiting for AnyDesk service.conf to appear...")
        time.sleep(2)
    if not os.path.exists(conf_path):
        print("Read AnyDesk ID failed: service.conf not found")
        return None
    # อ่านไฟล์หา ad.telemetry.last_cid
    try:
        with open(conf_path, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip().startswith("ad.telemetry.last_cid"):
                    return line.strip().split("=")[1].strip()
        print("ad.telemetry.last_cid not found in service.conf")
        return None
    except Exception as e:
        print("Read AnyDesk ID failed:", e)
        return None


# ------------------- Main Agent Loop -------------------
def main():
    global MACHINE_ID
    while True:
        MACHINE_ID = get_machine_id()
        if not MACHINE_ID:
            print("Retry get machine_id in 10s...")
            time.sleep(10)
            continue
        print(f"Agent started for machine_id={MACHINE_ID}")
        break
    # Main loop
    while True:
        try:
            cmds = poll_pending_commands(MACHINE_ID)
            for cmd in cmds:
                print(f"Executing command: {cmd['command']}")
                status = "done"
                result = None
                if cmd['command'] == "shutdown":
                    os.system("shutdown /s /t 5")
                elif cmd['command'] == "reboot":
                    os.system("shutdown /r /t 5")
                elif cmd['command'] == "reset":
                    install_anydesk()
                    set_anydesk_password("123456")
                    time.sleep(5)
                    anydesk_id = get_anydesk_id()
                    report_remote(MACHINE_ID, anydesk_id, None)
                elif cmd['command'] == "reinstall":

                    def write_unattend_xml():
                        unattend_xml = '''<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <InputLocale>en-US</InputLocale>
      <SystemLocale>en-US</SystemLocale>
      <UILanguage>en-US</UILanguage>
      <UILanguageFallback>en-US</UILanguageFallback>
      <UserLocale>en-US</UserLocale>
    </component>
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <OOBE>
        <HideEULAPage>true</HideEULAPage>
        <NetworkLocation>Work</NetworkLocation>
        <ProtectYourPC>3</ProtectYourPC>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
        <SkipMachineOOBE>true</SkipMachineOOBE>
        <SkipUserOOBE>true</SkipUserOOBE>
      </OOBE>
      <UserAccounts>
        <LocalAccounts>
          <LocalAccount wcm:action="add">
            <Name>FinoDDC</Name>
            <Group>Administrators</Group>
            <Password>
              <Value>123456</Value>
              <PlainText>true</PlainText>
            </Password>
          </LocalAccount>
        </LocalAccounts>
      </UserAccounts>
      <AutoLogon>
        <Username>FinoDDC</Username>
        <Enabled>true</Enabled>
        <!-- LogonCount removed for infinite auto login -->
      </AutoLogon>
      <TimeZone>SE Asia Standard Time</TimeZone>
      <RegisteredOwner>DDC</RegisteredOwner>
      <RegisteredOrganization>PC Rental</RegisteredOrganization>
      <ComputerName>*</ComputerName>
      <FirstLogonCommands>
        <SynchronousCommand wcm:action="add">
          <Order>1</Order>
          <Description>Install Python</Description>
          <CommandLine>powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe -OutFile C:\\python_installer.exe; Start-Process -FilePath C:\\python_installer.exe -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1' -Wait"</CommandLine>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>2</Order>
          <Description>Install AnyDesk</Description>
          <CommandLine>powershell -Command "Invoke-WebRequest -Uri https://download.anydesk.com/AnyDesk.exe -OutFile C:\\AnyDesk.exe; Start-Process -FilePath C:\\AnyDesk.exe -ArgumentList '/install /silent' -Wait"</CommandLine>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>3</Order>
          <Description>Set AnyDesk Password</Description>
          <CommandLine>powershell -ExecutionPolicy Bypass -File C:\\set_anydesk_pw.ps1</CommandLine>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>4</Order>
          <Description>Start Agent</Description>
          <CommandLine>powershell -Command "Start-Process python -ArgumentList 'C:\\project\\ddcweb\\client\\agent.py'"</CommandLine>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>5</Order>
          <Description>Report AnyDesk Info</Description>
          <CommandLine>powershell -ExecutionPolicy Bypass -File C:\\report_anydesk.ps1</CommandLine>
        </SynchronousCommand>
      </FirstLogonCommands>
    </component>
    <component name="Microsoft-Windows-LUA-Settings" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <EnableLUA>false</EnableLUA>
    </component>
  </settings>
</unattend>'''
                        path = r"C:\\Windows\\System32\\Sysprep\\unattend.xml"
                        print(f"[reinstall] Writing unattend.xml to {path}")
                        with open(path, "w", encoding="utf-8") as f:
                            f.write(unattend_xml)
                        # สร้างสคริปต์ powershell สำหรับตั้งรหัส anydesk และ report id+pw
                        print("[reinstall] Writing set_anydesk_pw.ps1 and report_anydesk.ps1")
                        with open(r"C:\\set_anydesk_pw.ps1", "w", encoding="utf-8") as f:
                            f.write("""
$pw = ConvertTo-SecureString '123456' -AsPlainText -Force
Set-ItemProperty -Path 'HKLM:\SOFTWARE\AnyDesk' -Name 'ad_password' -Value ([System.Text.Encoding]::UTF8.GetBytes('123456'))
""")
                        with open(r"C:\\report_anydesk.ps1", "w", encoding="utf-8") as f:
                            f.write("""
$anydesk_id = Get-Content 'C:\\ProgramData\\AnyDesk\\service.conf' | Select-String -Pattern 'ad_id' | ForEach-Object { $_.Line.Split('=')[1].Trim() }
$pw = '123456'
Invoke-RestMethod -Uri 'http://localhost:8000/machine/report_remote' -Method POST -Body (@{machine_name='FinoDDC'; anydesk_id=$anydesk_id; anydesk_password=$pw} | ConvertTo-Json) -ContentType 'application/json'
""")
                        return path
                    try:
                        unattend_path = write_unattend_xml()
                        sysprep_cmd = f"C:\\Windows\\System32\\Sysprep\\sysprep.exe /oobe /generalize /reboot /unattend:{unattend_path}"
                        print(f"[reinstall] Running: {sysprep_cmd}")
                        subprocess.run(sysprep_cmd, shell=True, check=True)
                        status = "done"
                        result = "Sysprep executed"
                    except Exception as e:
                        status = "failed"
                        result = f"Sysprep error: {e}"
                else:
                    status = "failed"
                    result = "Unknown command"
                report_command_result(cmd['id'], status, result, machine_id=MACHINE_ID)
            time.sleep(10)
        except Exception as e:
            print("Agent error:", e)
            time.sleep(10)

if __name__ == "__main__":
    main()
