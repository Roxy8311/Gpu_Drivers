import importlib.util
import subprocess
import sys


def module_installed(module_name):
    spec = importlib.util.find_spec(module_name)
    return spec is not None

def install_module(module_name):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', module_name])


print("Updating pip...")
subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
print("pip has been updated.")

required_modules = ['requests', 'bs4', 'wmi', 'pywin32']


for module in required_modules:
    if not module_installed(module):
        print(f"Installing {module}...")
        install_module(module)
        print(f"{module} has been installed.")


import requests
from bs4 import BeautifulSoup
import re
import wmi

def detect_graphics_cards():
    try:
        c = wmi.WMI()
        gpus = [(gpu.Caption.strip(), gpu.DriverVersion.strip()) for gpu in c.Win32_VideoController() if 'NVIDIA' in gpu.Caption]
        return gpus
    except Exception as e:
        print(f"Error detecting graphics card: {e}")
        return []

def is_notebook():
    try:
        c = wmi.WMI()
        for enclosure in c.Win32_SystemEnclosure():
            chassis_types = enclosure.ChassisTypes
            if chassis_types:
                print(f"Detected Chassis Types: {chassis_types}")
                for chassis_type in chassis_types:
                    if chassis_type in [8, 9, 10, 14]:
                        return True
        return False
    except Exception as e:
        print(f"Error detecting system enclosure: {e}")
        return False

def get_latest_driver_url(graphics_card_name):
    try:
        url = "https://www.nvidia.com/Download/processFind.aspx"
        params = {
            'psid': '123',
            'pfid': '939',
            'osid': '57',
            'lid': '1',
            'whql': '',
            'lang': 'fr-eu',
            'ctk': '0',
            'qnfslb': '00',
            'dtcid': '1'
        }

        response = requests.get(url, params=params)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        driver_link = None
        for link in soup.find_all('a', href=True):
            if "driverResult" in link['href']:
                driver_link = link['href']
                break

        if driver_link:
            return driver_link
        else:
            print("Could not find the latest driver link.")
            return None
    except requests.exceptions.RequestException as e:
        print(f"Error fetching NVIDIA driver page: {e}")
        return None
    except Exception as e:
        print(f"Error parsing HTML or finding driver link: {e}")
        return None

def adjust_url_for_non_notebook(url):
    match = re.search(r'(\d+)/en-us$', url)
    if match:
        number = int(match.group(1))
        adjusted_url = url.replace(f"{number}/en-us", f"{number - 1}/en-us")
        return adjusted_url
    else:
        print("URL format is unexpected. No adjustment made.")
        return url

def main():
    gpus = detect_graphics_cards()
    print('-' * 100)
    if gpus:
        for gpu_name, driver_version in gpus:
            print(f"Detected NVIDIA graphics card: {gpu_name} (Driver Version: {driver_version})")
            driver_link = get_latest_driver_url(gpu_name)
            if driver_link:
                if not is_notebook():
                    driver_link = adjust_url_for_non_notebook(driver_link)
                print(f"Latest driver download link for {gpu_name}: {driver_link}")
            else:
                print(f"Failed to retrieve the latest driver link for {gpu_name}.")
    else:
        print("No NVIDIA graphics cards detected.")
    print('-' * 100)

    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
