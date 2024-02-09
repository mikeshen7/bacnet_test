# Must install Visual Studio Build Tools 2022.  Select option for C++ development
# To activate virtual environment, run from powershell:
# ./.venv/Scripts/activate

# Git command: run from powershell
# git add .
# git commit -m "message"
# git push origin main


import BAC0
import time
from BAC0.core.devices.local.models import (
    analog_input,
    analog_output,
    analog_value,
    binary_input,
    binary_output,
    binary_value,
    character_string,
    date_value,
    datetime_value,
    humidity_input,
    humidity_value,
    make_state_text,
    multistate_input,
    multistate_output,
    multistate_value,
    temperature_input,
    temperature_value,
)
from BAC0.core.devices.local.object import ObjectFactory

### HELPER FUNCTIONS ###
def create_bacnet_device(device_instance, ip, port=47808, description=""):
    return BAC0.lite(ip=ip, deviceId=device_instance, port=port, description=description)

def create_bacnet_device_fdr(device_instance, ip, bbmdIP):
    return BAC0.lite(ip=ip, deviceId=device_instance, bbmdAddress=bbmdIP, bbmdTTL=900)

def add_points(device):
    analog_input(instance=0, description = "Space Temp", presentValue=71.0).add_objects_to_application(device)
    analog_input(instance=1, description = "DAT_input", presentValue=71.0).add_objects_to_application(device)
    analog_output(instance=0, description = "damper_cmd", presentValue=71.0).add_objects_to_application(device)
    analog_value(instance=15, description="DAT", presentValue=72.0).add_objects_to_application(device)
    binary_input(instance=0, description = "Fan Status", presentValue=False).add_objects_to_application(device)
    binary_output(instance=0, description = "fan_cmd", presentValue=False).add_objects_to_application(device)
    binary_value(instance=0, description = "Heat_cool mode", presentValue=False).add_objects_to_application(device)
    return

def start_device(device):
    device.this_application
    
    while True:
        time.sleep(0.01)
    return

def main():
    # Create BACnet device
    dev100001 = create_bacnet_device(100001, "192.168.105.39/24", 47808)

    # Add points
    add_points(dev100001)

    # Start BACnet device
    start_device(dev100001)

    return


if __name__ == '__main__':
    main()
