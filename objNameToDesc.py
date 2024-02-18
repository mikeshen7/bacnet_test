import BAC0
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
import re
import time
import configparser
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import numpy as np
import os
import re
import warnings
import logging
import math
import datetime
from dotenv import load_dotenv

### Logging Settings ###
# Silence (use CRITICAL so not much messages will be sent)
BAC0.log_level("silence")

### Ignore openpyxl data validation warning
warnings.simplefilter(action="ignore", category=UserWarning)

# Configure logging to write to a file named 'log.txt'
logging.basicConfig(
    filename="log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Customer logging for just BACnet write operations
bacnet_logger = logging.getLogger("custom_logger")
bacnet_logger.setLevel(logging.INFO)
file_handler = logging.FileHandler("bacnet log.txt")
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
file_handler.setFormatter(formatter)
bacnet_logger.addHandler(file_handler)


### CLASSES ###
class Device:
    def __init__(self, address, device_instance, ip_address, net, mac):
        self.address = address
        self.deviceInstance = device_instance
        self.ipAddress = ip_address
        self.net = net
        self.mac = mac

    def __repr__(self):
        return f"Device(deviceInstance={self.deviceInstance}, ipAddress={self.ipAddress}, net={self.net}, mac={self.mac})"


### FUNCTIONS ###
def bacnet_initialize():
    """
    Parameters: None
    Takes in BACnet configuration parameters from settings.ini
    Returns: BACnet device

    REV History:
    2024-02-08 (mikes): initial
    """
    # Takes in BACnet configuration parameters from settings.ini
    # Creates and returns a BACnet device
    config = configparser.ConfigParser()
    config.read("settings.ini")
    ipAddress = config.get("bacnet", "ipAddress")
    udpPort = config.get("bacnet", "udpPort")
    bacnet = BAC0.lite(ip=ipAddress, port=udpPort)
    return bacnet


def range_to_list(range_string):
    """
    Parameters:
    - range_string: string with ranges broken up by semi-colons.
    Example: 1000-1010;2000-3000;4000;5000

    Return: a list

    REV History:
    2024-02-18 (mikes): initial
    """

    result = []

    for item in range_string.split(";"):
        if "-" in item:
            start, end = map(int, item.split("-"))
            result.extend(range(start, end + 1))
        elif item:  # Add this condition to handle semicolon at the end
            result.append(int(item))
    return result


def build_device_manager(bacnet, DI_list):
    """
    Parameters: bacnet device, list of Device Instances
    Calls device_scan() once for entire range, then scans for missing devices
    Return: Pandas df with address information for each Device Instance

    REV History:
    2024-02-08 (mikes): initial
    """

    # Scan for entire range from max to min
    device_manager = device_scan(bacnet, min(DI_list), max(DI_list))
    # device_manager = device_scan(bacnet, min(DI_list), 1200)

    # Scan for missing items
    for device in DI_list:
        if device not in np.array(device_manager["deviceInstance"].values, dtype=np.int64):
            print(f"Scanning for missing device {device}")
            device_info = device_scan(bacnet, device, device)
            if device_info is not None:  # Check if device_info is not empty
                device_manager = pd.concat([device_manager, device_info], ignore_index=True)

    # Remove duplicate rows based on 'deviceInstance' column
    device_manager = device_manager.drop_duplicates(subset="deviceInstance", keep="first").reset_index(drop=True)

    # Remove deviees that aren't in DI_list
    device_manager = device_manager[device_manager["deviceInstance"].isin(DI_list)]

    return device_manager


def device_scan(bacnet, start_instance, end_instance):
    """
    Parameters: bacnet device, list of Device Instances
    Conducts a scan for a set range of device instances.  Will repeat scans based on settings.ini
    Return: Pandas df with address information for each Device Instance

    REV History:
    2024-02-08 (mikes): initial
    """

    config = configparser.ConfigParser()
    config.read("settings.ini")
    scanTimeout = config.getint("bacnet", "scanTimeout")

    device_manager = []
    passes = 1

    while passes <= scanTimeout:
        # If single instance scan and device is found, break
        if start_instance == end_instance and any(device.deviceInstance == start_instance for device in device_manager):
            break

        print(f"Scan for devices {start_instance} to {end_instance}, Pass # {passes}")

        # Scan for device
        bacnet.discover(limits=(start_instance, end_instance))
        device_dict = bacnet.discoveredDevices

        for key, value in device_dict.items():
            address = key[0]
            deviceInstance = key[1]
            ipAddress = ""
            net = ""
            mac = ""

            if re.match(r"^(\d{1,3}\.){3}\d{1,3}$", address):
                ipAddress = address
            else:
                address_parts = address.split(":")
                if len(address_parts) == 2:
                    net = address_parts[0]
                    mac = address_parts[1]

            device = Device(address, deviceInstance, ipAddress, net, mac)
            if all(device.deviceInstance != d.deviceInstance for d in device_manager):
                device_manager.append(device)
                print("** Found device " + str(device.deviceInstance))
                device_found = True

        passes += 1

    device_manager = sorted(device_manager, key=lambda x: x.deviceInstance)

    # Create a dictionary to store device information
    data = {
        "address": [device.address for device in device_manager],
        "deviceInstance": [device.deviceInstance for device in device_manager],
        "IP": [device.ipAddress for device in device_manager],
        "Network": [device.net for device in device_manager],
        "MAC": [device.mac for device in device_manager],
    }

    # Create a Pandas DataFrame from the device information
    df = pd.DataFrame(data)

    return df


def read_point(bacnet, device_manager, device_instance, object_type, object_instance, property, index=None):
    """
    Parameters: lots
    Performs BACnet read
    Return: BACnet value.  If error, returns "NR"

    Example:
    value = read_point(bacnet, device_manager, 1001, "binaryOutput", "0", "priorityArray", 14)

    REV History:
    2024-02-08 (mikes): initial
    """

    # Variables
    address = None
    value = None

    # Get bacnet address from device_manager
    if device_instance in device_manager["deviceInstance"].values:
        # Get the address for device_instance 1000
        address = device_manager.loc[device_manager["deviceInstance"] == device_instance, "address"].iloc[0]
    else:
        return "NR"

    # Read BACnet point
    try:
        # Read DDC file name
        if object_type == "program":
            value = bacnet.read(f"{address} program 0 {property}")

        # Read from device
        elif object_type == "device":
            value = bacnet.read(f"{address} device {device_instance} {property}")

        # Read point
        else:
            value = bacnet.read(f"{address} {object_type} {object_instance} {property}")

    except Exception as e:
        logging.error(
            f"read_point error.  error: {e} device: {device_instance} object_type: {object_type} object_instance: {object_instance} property: {property} index: {index}"
        )
        value = "NR"
        return value

    # Unpack priority array
    if property == "priorityArray":
        array = serialize_priority_array(value.dict_contents(), object_type)
        if str(index) in array.keys():
            value = array[str(index)]
        else:
            value = "NR"

    return value


def write_point(bacnet, device_manager, device_instance, object_type, object_instance, property, value, index=None):
    """
    Parameters: lots
    Performs BACnet write
    Logs previous value prior to writing
    Return: None

    Example:
    write_point(bacnet, device_manager, 1001, "binaryOutput", "20", "priorityArray", "inactive", 3)

    REV History:
    2024-02-08 (mikes): initial
    """

    # Variables
    address = None
    value = re.sub(r"\s+", "_", str(value))

    # Get bacnet address from device_manager
    if device_instance in device_manager["deviceInstance"].values:
        # Get the address for device_instance 1000
        address = device_manager.loc[device_manager["deviceInstance"] == device_instance, "address"].iloc[0]
    else:
        return

    # Don't allow writing to program
    if object_type == "program":
        return

    # Read BACnet point and log value
    read_value = read_point(bacnet, device_manager, device_instance, object_type, object_instance, property, index)

    if index is None:
        bacnet_logger.info(f"Writing to {device_instance}:{object_type}{object_instance} {property}.  Original value: {read_value}")
    else:
        bacnet_logger.info(
            f"Writing to {device_instance}:{object_type}{object_instance} {property} priority {index}.  Original value: {read_value}"
        )

    # Write BACnet point
    try:
        # Write to device
        if object_type == "device":
            bacnet.write(f"{address} device {device_instance} {property} {value}")

        # Write to priority array
        elif property == "priorityArray":
            bacnet.write(f"{address} {object_type} {object_instance} presentValue {value} - {index}")

        # Write to point
        else:
            bacnet.write(f"{address} {object_type} {object_instance} {property} {value}")

    except Exception as e:
        logging.error(
            f"write_point error.  error: {e} device: {device_instance} object_type: {object_type} object_instance: {object_instance} property: {property} index: {index}"
        )
        bacnet_logger.error(
            f"write_point error.  error: {e} device: {device_instance} object_type: {object_type} object_instance: {object_instance} property: {property} index: {index}"
        )

        return None
    return


def objName_to_description(bacnet, DI_range, av_range, bv_range, mv_range, ai_range, bi_range, mi_range, ao_range, bo_range, mo_range):
    """
    Parameters:
    - xx_range: range of instances

    Main call to read BACnet object name and write to description
    Return: none

    REV History:
    2024-02-18 (mikes): initial
    """

    # Check for invalid DI Range
    if (DI_range == "") or (DI_range is None):
        return

    # Convert ranges to lists
    DI_list = range_to_list(DI_range)
    av_list = range_to_list(av_range)
    bv_list = range_to_list(bv_range)
    mv_list = range_to_list(mv_range)
    ai_list = range_to_list(ai_range)
    bi_list = range_to_list(bi_range)
    mi_list = range_to_list(mi_range)
    ao_list = range_to_list(ao_range)
    bo_list = range_to_list(bo_range)
    mo_list = range_to_list(mo_range)

    # Build device manager
    device_manager = build_device_manager(bacnet, DI_list)

    print(device_manager)

    # Iterate through each DI.
    for device_instance in DI_list:
        # Check to see if device is in device_manager.  Skip if not in there
        if device_instance in device_manager["deviceInstance"].values:
            print(f"Reading from {device_instance}...")

            # Iterate through AV list
            for instance in av_list:
                # Read the av object name
                value = read_point(bacnet, device_manager, device_instance, "analogValue", instance, "objectName")

                # Write to the av description
                if value == "NR":
                    print(f"Point not read: {device_instance}:AV{instance} {value}")
                else:
                    print(f"Copying {device_instance}:AV{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "analogValue", instance, "description", value)

            # Iterate through BV list
            for instance in bv_list:
                # Read the av object name
                value = read_point(bacnet, device_manager, device_instance, "binaryValue", instance, "objectName")

                # Write to the av description
                if value == "NR":
                    print(f"Point not read: {device_instance}:BV{instance} {value}")
                else:
                    print(f"Copying {device_instance}:BV{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "binaryValue", instance, "description", value)

            # Iterate through MV list
            for instance in mv_list:
                # Read the mv object name
                value = read_point(bacnet, device_manager, device_instance, "multiStateValue", instance, "objectName")

                # Write to the mv description
                if value == "NR":
                    print(f"Point not read: {device_instance}:MV{instance} {value}")
                else:
                    print(f"Copying {device_instance}:MV{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "multiStateValue", instance, "description", value)

            # Iterate through AI list
            for instance in ai_list:
                # Read the ai object name
                value = read_point(bacnet, device_manager, device_instance, "analogInput", instance, "objectName")

                # Write to the ai description
                if value == "NR":
                    print(f"Point not read: {device_instance}:AI{instance} {value}")
                else:
                    print(f"Copying {device_instance}:AI{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "analogInput", instance, "description", value)

            # Iterate through BI list
            for instance in bi_list:
                # Read the bi object name
                value = read_point(bacnet, device_manager, device_instance, "binaryInput", instance, "objectName")

                # Write to the bi description
                if value == "NR":
                    print(f"Point not read: {device_instance}:BI{instance} {value}")
                else:
                    print(f"Copying {device_instance}:BI{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "binaryInput", instance, "description", value)

            # Iterate through MI list
            for instance in mi_list:
                # Read the mi object name
                value = read_point(bacnet, device_manager, device_instance, "multiStateInput", instance, "objectName")

                # Write to the mi description
                if value == "NR":
                    print(f"Point not read: {device_instance}:MI{instance} {value}")
                else:
                    print(f"Copying {device_instance}:MI{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "multiStateInput", instance, "description", value)

            # Iterate through AO list
            for instance in ao_list:
                # Read the ao object name
                value = read_point(bacnet, device_manager, device_instance, "analogOutput", instance, "objectName")

                # Write to the ao description
                if value == "NR":
                    print(f"Point not read: {device_instance}:AO{instance} {value}")
                else:
                    print(f"Copying {device_instance}:AO{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "analogOutput", instance, "description", value)

            # Iterate through BO list
            for instance in bo_list:
                # Read the bo object name
                value = read_point(bacnet, device_manager, device_instance, "binaryOutput", instance, "objectName")

                # Write to the bo description
                if value == "NR":
                    print(f"Point not read: {device_instance}:BO{instance} {value}")
                else:
                    print(f"Copying {device_instance}:BO{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "binaryOutput", instance, "description", value)

            # Iterate through MO list
            for instance in mo_list:
                # Read the mo object name
                value = read_point(bacnet, device_manager, device_instance, "multiStateOutput", instance, "objectName")

                # Write to the mo description
                if value == "NR":
                    print(f"Point not read: {device_instance}:MO{instance} {value}")
                else:
                    print(f"Copying {device_instance}:MO{instance} {value}")
                    write_point(bacnet, device_manager, device_instance, "multiStateOutput", instance, "description", value)


        else:
            print(f"{device_instance} not found.  Skipping")
            bacnet_logger.error(f"Read error.  Device: {device_instance} not found.  Skipped.")
    return


def main():
    bacnet = bacnet_initialize()
    DI_range = "1001"
    av_range = "2-3;492384"
    bv_range = "0-2;7"
    mv_range = "9000;9001"

    ai_range = "0-1"
    bi_range = "0-1"
    mi_range = "9000;9001"

    ao_range = "6-7"
    bo_range = "0-5"
    mo_range = "0-5"

    objName_to_description(bacnet, DI_range, av_range, bv_range, mv_range, ai_range, bi_range, mi_range, ao_range, bo_range, mo_range)


if __name__ == "__main__":
    main()
