# bacnet_test

# Must install Visual Studio Build Tools 2022.  Select option for C++ development
# To activate virtual environment, run from powershell:
# ./.venv/Scripts/activate

# Git command: run from powershell
# git add .
# git commit -m "message"
# git push origin main


<!-- 

    Found devices
    print(bacnet.devices)
    [('Device 1000', 'Alerton', '192.168.105.7', 1000), 
    ('Device 1603', 'Alerton', '192.168.105.63', 1603), 
    ('Dev 0001601', 'Alerton', '192.168.105.8', 1601), 
    ('DEV-1001', 'ALERTON TECHNOLOGIES', '1000:1', 1001), 
    ('DEV-1101', 'ALERTON TECHNOLOGIES', '1001:2', 1101), 
    ('Device 1100', 'Alerton', '1007:0x7f000001b4c2', 1100), 
    ('Device 1100', 'Alerton', '1007:0x7f000001b4c3', 1200)]

    discovered devices
    print(bacnet.discoveredDevices)
    defaultdict(<class 'int'>, 
    {('192.168.105.7', 1000): 1, 
    ('192.168.105.63', 1603): 1, 
    ('192.168.105.8', 1601): 1, 
    ('1000:1', 1001): 1, 
    ('1001:2', 1101): 1, 
    ('1007:0x7f000001b4c2', 1100): 1, 
    ('1007:0x7f000001b4c3', 1200): 1})

    Known network numbers
    print(bacnet.known_network_numbers)
    {999, 1000, 1001, 1007, 1009}

    routing table
    print(bacnet.routing_table)
    {'192.168.105.7': Source Network: 999 | Address: 192.168.105.7 | Destination Networks: {1000: 0, 1001: 0, 1007: 0, 1009: 0} | Path: (999, 1009)}
 -->


