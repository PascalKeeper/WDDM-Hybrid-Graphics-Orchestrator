"""
Copyright Joseph Peransi, two thousand twenty-six. Be excellent to each other.

System: Intel UHD 630 (iGPU) + NVIDIA GTX 1070 (dGPU)
Objective: automated WDDM Hybrid Graphics orchestration.
Features: Auto-Elevation, Context Menu Integration, Registry Injection, Power Plan Enforcement.
"""

import os
import sys
import json
import subprocess
from pathlib import Path

# --- OS COMPATIBILITY GATE ---
# This script relies on the Windows Registry (winreg) and Windows API (ctypes.windll).
# It will crash immediately on Linux/macOS without this check.
if sys.platform != 'win32':
    print(f"[CRITICAL] This script is strictly for Windows 10/11 WDDM environments.")
    print(f"           Detected OS: {sys.platform}")
    print("           Please transfer this file to your Windows machine to execute it.")
    sys.exit(1)

import ctypes
import winreg

class HybridOrchestrator:
    def __init__(self):
        self.reg_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
        self.context_menu_path = r"exefile\shell\RunOnGTX1070"
        self.nvidia_pci_id = None
        self.intel_pci_id = None

    def elevate(self):
        """Re-launches the script as Administrator if not already."""
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("[*] Requesting elevation...")
            try:
                # Re-run the script with Admin privileges
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, " ".join(sys.argv), None, 1
                )
            except Exception as e:
                print(f"[ERROR] Elevation failed: {e}")
            sys.exit(0)

    def detect_hardware(self):
        print("[*] Scanning WDDM Topology via CIM...")
        try:
            # Get hardware info as JSON for robust parsing
            cmd = "Get-CimInstance Win32_VideoController | Select-Object Name, DeviceID, PNPDeviceID | ConvertTo-Json"
            result = subprocess.check_output(["powershell", "-Command", cmd], text=True)
            
            # Handle single vs multiple GPU return format from PowerShell JSON
            data = json.loads(result)
            if isinstance(data, dict): data = [data]

            for gpu in data:
                name = gpu.get('Name', 'Unknown')
                pnp_id = gpu.get('PNPDeviceID', '')
                
                if "Intel" in name or "UHD" in name:
                    self.intel_pci_id = pnp_id
                    print(f"[+] Integrated: {name}")
                elif "NVIDIA" in name or "GTX" in name:
                    self.nvidia_pci_id = pnp_id
                    print(f"[+] Discrete:   {name}")

            if not self.nvidia_pci_id:
                print("[WARNING] Discrete GPU not detected via CIM. Proceeding assuming GTX 1070 exists.")

        except Exception as e:
            print(f"[ERROR] Hardware scan failed: {e}")

    def optimize_power_plan(self):
        """Enforces Ultimate Performance and disables PCI Link State Power Management."""
        print("[*] Optimizing Power Subsystem...")
        try:
            # Ultimate Performance GUID
            guid = "e9a42b02-d5df-448d-aa00-03f14749eb61"
            subprocess.run(["powercfg", "-duplicatescheme", guid], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.run(["powercfg", "-setactive", guid], check=True)
            
            # Disable PCI Express Link State Power Management (Plugged In)
            # 0 = Off (Maximum Performance)
            subprocess.run(["powercfg", "/setacvalueindex", "SCHEME_CURRENT", "SUB_PCIEXPRESS", "ASPM", "0"])
            subprocess.run(["powercfg", "/setactive", "SCHEME_CURRENT"])
            print("[+] Power Plan: Ultimate Performance (PCIe ASPM Disabled)")
        except:
            print("[!] Failed to set power plan. Ensure you have 'Ultimate Performance' enabled in Windows.")

    def set_registry_preference(self, app_path, force_gpu=True):
        """Writes the DX preference to HKCU."""
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.reg_path)
            # 0=Let Windows Decide, 1=Power Saving, 2=High Performance
            val = "GpuPreference=2;" if force_gpu else "GpuPreference=1;"
            winreg.SetValueEx(key, str(app_path), 0, winreg.REG_SZ, val)
            winreg.CloseKey(key)
            print(f"[+] Registry: {os.path.basename(app_path)} -> High Performance")
        except Exception as e:
            print(f"[ERROR] Registry write failed: {e}")

    def inject_context_menu(self):
        """Adds a right-click context menu to .exe files."""
        print("[*] Injecting Context Menu extension...")
        try:
            # Create key HKCR\exefile\shell\RunOnGTX1070
            key_path = r"Software\Classes\exefile\shell\RegisterHighPerfGPU"
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValue(key, "", winreg.REG_SZ, "Register for GTX 1070 Performance")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "imageres.dll,-1010") # GPU icon style
            
            # The command to run: calls this script with the target file
            command_key = winreg.CreateKey(key, "command")
            script_path = str(Path(sys.argv[0]).resolve())
            
            # If running as python script
            if script_path.endswith(".py"):
                cmd_str = f'"{sys.executable}" "{script_path}" "%1"'
            else:
                # If compiled to exe
                cmd_str = f'"{script_path}" "%1"'
                
            winreg.SetValue(command_key, "", winreg.REG_SZ, cmd_str)
            print("[+] Context Menu added! Right-click any .exe to register it.")
        except Exception as e:
            print(f"[ERROR] Context menu injection failed: {e}")

    def set_legacy_environment(self):
        """Sets environment variables for current session (helps older apps)."""
        os.environ["SHIM_MCCOMPAT"] = "0x800000001" # Hybrid Override
        os.environ["CUDA_VISIBLE_DEVICES"] = "0"    # Usually dGPU is index 0 in CUDA
        print("[+] Active Session Environment Variables injected.")

    def run(self):
        self.elevate()
        print("--- WDDM Hybrid Graphics Orchestrator v2.0 ---")
        
        # If arguments provided, we are in "Context Menu Mode" (Registering specific app)
        if len(sys.argv) > 1:
            target_app = sys.argv[1]
            print(f"[*] Processing request for: {target_app}")
            self.set_registry_preference(target_app)
            # Keep window open briefly to show success
            subprocess.run(["timeout", "/t", "3"], shell=True)
            sys.exit(0)

        # Interactive / Setup Mode
        self.detect_hardware()
        self.optimize_power_plan()
        self.inject_context_menu()
        self.set_legacy_environment()
        
        # Register standard pro apps if they exist
        defaults = [
            r"C:\Program Files\Blender Foundation\Blender\blender.exe",
            r"C:\Windows\System32\cmd.exe", # Useful for launching CLI tools on dGPU
        ]
        for app in defaults:
            if os.path.exists(app):
                self.set_registry_preference(app)

        print("\n[SUCCESS] System Optimized.")
        print("1. Right-click any .exe and select 'Register for GTX 1070 Performance'")
        print("2. Power plan locked to Ultimate Performance.")
        input("Press Enter to exit...")

if __name__ == "__main__":
    orchestrator = HybridOrchestrator()
    orchestrator.run()
