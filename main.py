import BAC0
# Must install Visual Studio Build Tools 2022.  Select option for C++ development
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



def add_points(qty_per_type, device):
    # Start from fresh
    ObjectFactory.clear_objects()
    basic_qty = qty_per_type - 1
    # Analog Inputs
    # Default... percent
    for _ in range(basic_qty):
        _new_objects = analog_input(presentValue=99.9)
        _new_objects = multistate_value(presentValue=1)

    # Supplemental with more details, for demonstration
    _new_objects = analog_input(
        name="ZN-T",
        properties={"units": "degreesCelsius"},
        description="Zone Temperature",
        presentValue=21,
    )

    states = make_state_text(["Normal", "Alarm", "Super Emergency"])
    _new_objects = multistate_value(
        description="An Alarm Value",
        properties={"stateText": states},
        name="BIG-ALARM",
        is_commandable=True,
    )

    # All others using default implementation
    for _ in range(qty_per_type):
        _new_objects = analog_output(presentValue=89.9)
        _new_objects = analog_value(presentValue=79.9)
        _new_objects = binary_input()
        _new_objects = binary_output()
        _new_objects = binary_value()
        _new_objects = multistate_input()
        _new_objects = multistate_output()
        _new_objects = date_value()
        _new_objects = datetime_value()
        _new_objects = character_string(presentValue="test", is_commandable=True)

    _new_objects.add_objects_to_application(device)


# Create BACnet device: local BACnet device
my_ip = "192.168.105.80/24"
device100001 = BAC0.lite(ip=my_ip, deviceId=100001)

# Create BACnet device: Register as Foreign Device
# my_ip = "192.168.105.80/24"
# bbmdIP = '192.168.105.68:47808'
# bbmdTTL = 900
# bacnet = BAC0.lite(ip=my_ip, deviceId=100001, bbmdAddress=bbmdIP, bbmdTTL=bbmdTTL)

# Add points
analog_input(instance=0, description = "Space Temp", presentValue=71.0).add_objects_to_application(device100001)
analog_output(instance=0, description = "damper_cmd", presentValue=71.0).add_objects_to_application(device100001)
analog_value(instance=15, description="DAT", presentValue=72.0, ).add_objects_to_application(device100001)
binary_input(instance=0, description = "Fan Status", presentValue=False).add_objects_to_application(device100001)
binary_output(instance=0, description = "fan_cmd", presentValue=False).add_objects_to_application(device100001)
binary_value(instance=0, description = "Heat_cool mode", presentValue=False).add_objects_to_application(device100001)

device100001.this_application

while True:
    time.sleep(0.01)
