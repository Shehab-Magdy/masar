import json
import os
import hashlib
import hmac
import base64
from cryptography.fernet import Fernet
import uuid
import platform
import subprocess

class LicenseManager:
    def __init__(self, license_file='license.json'):
        self.license_file = license_file
        # Obfuscated keys
        self.fernet_key = base64.b64decode('NjhyX0tpLTMxVmFjTjB5bmZTTHFYbW1rUXpUVlNUYjlZcFN4dEdOVGxSdz0=')
        self.hmac_key = base64.b64decode('HrIJ4BDWQbKGt40uWL6RuBawhB25z7Jm2DhPdmnTdY8=')  # 32 bytes base64

    def get_hardware_id(self) -> str:
        """Generate hardware fingerprint"""
        components = []
        
        # MAC Address
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) for elements in range(0,8*6,8)][::-1])
        components.append(f"MAC:{mac}")
        
        # OS UUID (if available)
        try:
            if platform.system() == 'Windows':
                import winreg
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography")
                machine_guid, _ = winreg.QueryValueEx(key, "MachineGuid")
                components.append(f"GUID:{machine_guid}")
            elif platform.system() == 'Linux':
                with open('/etc/machine-id', 'r') as f:
                    components.append(f"MachineID:{f.read().strip()}")
            elif platform.system() == 'Darwin':  # macOS
                result = subprocess.run(['system_profiler', 'SPHardwareDataType'], capture_output=True, text=True)
                for line in result.stdout.split('\n'):
                    if 'Hardware UUID' in line:
                        guid = line.split(':')[1].strip()
                        components.append(f"GUID:{guid}")
                        break
        except:
            pass
        
        # CPU Info
        cpu = platform.processor()
        if cpu:
            components.append(f"CPU:{cpu}")
        
        # Concatenate and hash
        data = '|'.join(components)
        return hashlib.sha256(data.encode()).hexdigest()

    def encrypt_A(self, data: str) -> str:
        """Encrypt with Fernet (Method A)"""
        f = Fernet(self.fernet_key)
        return f.encrypt(data.encode()).decode()

    def decrypt_A(self, data: str) -> str:
        """Decrypt with Fernet (Method A)"""
        f = Fernet(self.fernet_key)
        return f.decrypt(data.encode()).decode()

    def encrypt_B(self, data: str) -> str:
        """Generate activation code with HMAC (Method B)"""
        h = hmac.new(self.hmac_key, data.encode(), hashlib.sha256)
        sig = h.hexdigest()
        return f"{data}|{sig}"

    def decrypt_B(self, data: str) -> str:
        """Verify activation code with HMAC (Method B)"""
        parts = data.split('|', 1)
        if len(parts) != 2:
            raise ValueError("Invalid activation code")
        hid, sig = parts
        h = hmac.new(self.hmac_key, hid.encode(), hashlib.sha256)
        if hmac.compare_digest(h.hexdigest(), sig):
            return hid
        raise ValueError("Invalid signature")

    def load_license(self) -> dict:
        """Load license from file"""
        if not os.path.exists(self.license_file):
            return {}
        try:
            with open(self.license_file, 'r') as f:
                return json.load(f)
        except:
            return {}

    def save_license(self, data: dict) -> None:
        """Save license to file"""
        with open(self.license_file, 'w') as f:
            json.dump(data, f, indent=2)

    def validate_license(self) -> bool:
        """Validate existing license"""
        license_data = self.load_license()
        if 'HID_A' not in license_data:
            return False
        try:
            decrypted_hid = self.decrypt_A(license_data['HID_A'])
            current_hid = self.get_hardware_id()
            return decrypted_hid == current_hid
        except:
            return False

    def activate_license(self, activation_code: str) -> bool:
        """Activate license with activation code"""
        try:
            decrypted_hid = self.decrypt_B(activation_code)
            current_hid = self.get_hardware_id()
            if decrypted_hid == current_hid:
                license_data = {
                    'HID_A': self.encrypt_A(current_hid),
                    'HID_B': activation_code
                }
                self.save_license(license_data)
                return True
            return False
        except:
            return False