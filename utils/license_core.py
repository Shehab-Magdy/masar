import json
import os
import hashlib
import hmac
import base64
import uuid
import platform
import subprocess

# Crockford Base32: excludes I, L, O, U to avoid visual ambiguity with 1/1/0/V.
_B32_ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"
_CODE_BYTES = 10  # 80 bits -> 16 base32 chars -> grouped as XXXX-XXXX-XXXX-XXXX

class LicenseManager:
    def __init__(self, license_file='license.json'):
        self.license_file = license_file
        # Obfuscated key
        self.hmac_key = base64.b64decode('HrIJ4BDWQbKGt40uWL6RuBawhB25z7Jm2DhPdmnTdY8=')  # 32 bytes base64

    def _collect_fingerprint(self) -> str:
        """Collect raw hardware fingerprint components"""
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

        return '|'.join(components)

    @staticmethod
    def _encode_short_code(raw: bytes) -> str:
        """Encode raw bytes as a short, human-readable grouped code (e.g. XK3F-9QPL-7ZRT-M2CD)"""
        total_bits = len(raw) * 8
        num_chars = -(-total_bits // 5)  # ceil(total_bits / 5)
        value = int.from_bytes(raw, 'big')
        chars = []
        for i in range(num_chars - 1, -1, -1):
            chars.append(_B32_ALPHABET[(value >> (i * 5)) & 0x1f])
        code = ''.join(chars)
        return '-'.join(code[i:i + 4] for i in range(0, len(code), 4))

    @staticmethod
    def _normalize(code: str) -> str:
        """Normalize a user/admin-entered code for comparison (strip separators, uppercase)"""
        return ''.join(ch for ch in code.strip().upper() if ch.isalnum())

    def get_hardware_id(self) -> str:
        """Short, human-readable hardware code, e.g. XK3F-9QPL-7ZRT-M2CD"""
        data = self._collect_fingerprint()
        digest = hashlib.sha256(data.encode()).digest()[:_CODE_BYTES]
        return self._encode_short_code(digest)

    def generate_activation_code(self, hardware_id: str) -> str:
        """Admin-side: sign a hardware code into a short activation code"""
        normalized = self._normalize(hardware_id)
        sig = hmac.new(self.hmac_key, normalized.encode(), hashlib.sha256).digest()[:_CODE_BYTES]
        return self._encode_short_code(sig)

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
        stored_code = license_data.get('activation_code')
        if not stored_code:
            return False
        try:
            current_hid = self.get_hardware_id()
            expected_code = self.generate_activation_code(current_hid)
            return hmac.compare_digest(self._normalize(expected_code), self._normalize(stored_code))
        except:
            return False

    def activate_license(self, activation_code: str) -> bool:
        """Activate license with activation code"""
        try:
            current_hid = self.get_hardware_id()
            expected_code = self.generate_activation_code(current_hid)
            if hmac.compare_digest(self._normalize(expected_code), self._normalize(activation_code)):
                license_data = {
                    'hardware_id': current_hid,
                    'activation_code': expected_code
                }
                self.save_license(license_data)
                return True
            return False
        except:
            return False
