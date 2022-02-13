# Column supports both column index or Column string name


# Column 'A' is always reserved for Hostname
# keys is a list of keys that can be found in the gather_{function}, it can vary depending on the textfsm file
# name of work sheet should be gather_function minus the gather_, i.g. for gather_cdp, name of worksheet should be Cdp,
# First letter is always capitalized, if sheet does not exist then one will be generated.
# Only when the Work Sheet is  generated will the column titles be added.
OUTPUT_KEY_MAP = {
    'gather_inventory': [
        {'column': 'B', 'column_name': 'Part ID', 'keys': ['pid']},
        {'column': 'C', 'column_name': 'Device', 'keys': ['name']},
        {'column': 'D', 'column_name': 'Serial', 'keys': ['sn', 'serial']},
        {'column': 'E', 'width': 40,  'column_name': 'Description', 'keys': ['description']},
    ],
    'gather_cdp': [
        {'column': 2, 'column_name': 'Local Port', 'keys': ['local_port']},
        {'column': 3, 'column_name': 'Remote Host', 'keys': ['destination_host']},
        {'column': 4, 'column_name': 'Remote Port', 'keys': ['remote_port']},
        {'column': 5, 'column_name': 'Interface IP', 'keys': ['interface_ip']},
        {'column': 6, 'column_name': 'Management IP', 'keys': ['management_ip']},
        {'column': 7, 'column_name': 'Platform', 'keys': ['platform']},
        {'column': 8, 'column_name': 'Software', 'keys': ['software_version']},
        {'column': 9, 'column_name': 'Capabilities', 'keys': ['capabilities']},
    ],
    'gather_lldp': [
        {'column': 2, 'column_name': 'Local Port', 'keys': ['local_interface']},
        {'column': 3, 'column_name': 'Remote Host', 'keys': ['neighbor']},
        {'column': 4, 'column_name': 'Remote Port', 'keys': ['neighbor_interface']},
        {'column': 5, 'column_name': 'Management IP', 'keys': ['mgmt_address']},
        {'column': 6, 'column_name': 'Chassis ID', 'keys': ['chassis_id']},
        {'column': 7, 'column_name': 'Capabilities', 'keys': ['capabilities']},
        {'column': 8, 'column_name': 'System Description', 'keys': ['system_description']},
    ],
    'gather_mac': [
        {'column': 'B', 'column_name': 'Destination Address', 'keys': ['destination_address', 'mac']},
        {'column': 'C', 'column_name': 'Type', 'keys': ['type']},
        {'column': 'D', 'column_name': 'VLAN', 'keys': ['vlan']},
        {'column': 'E', 'column_name': 'Destination Port', 'keys': ['destination_port', 'ports']},
    ],
    'gather_arp': [
        {'column': 'B', 'column_name': 'VRF', 'keys': ['vrf']},
        {'column': 'C', 'column_name': 'IP Address', 'keys': ['address']},
        {'column': 'D', 'column_name': 'Age', 'keys': ['age']},
        {'column': 'E', 'column_name': 'Hardware/MAC', 'keys': ['mac']},
        {'column': 'F', 'column_name': 'Type', 'keys': ['type']},
        {'column': 'G', 'column_name': 'Interface', 'keys': ['interface']},
    ],
    'gather_interfaces': [
        {'column': 'B', 'column_name': 'Interface', 'keys': ['interface']},
        {'column': 'C', 'column_name': 'Description', 'keys': ['description']},
        {'column': 'D', 'column_name': 'Type', 'keys': ['hardware_type']},
        {'column': 'E', 'column_name': 'VRF', 'keys': ['vrf']},
        {'column': 'F', 'column_name': 'Link', 'keys': ['link_status']},
        {'column': 'G', 'column_name': 'Protocol', 'keys': ['protocol_status', 'admin_state']},
        {'column': 'H', 'column_name': 'L2/L3', 'keys': ['l2_l3']},
        {'column': 'I', 'column_name': 'Trunk/Access', 'keys': ['trunk_access']},
        {'column': 'J', 'column_name': 'Access VLAN', 'keys': ['vlan']},
        {'column': 'K', 'column_name': 'Trunk Allowed', 'keys': ['allowed']},
        {'column': 'L', 'column_name': 'Trunk Forwarding', 'keys': ['not_pruned']},
        {'column': 'M', 'column_name': 'Native VLAN', 'keys': ['native']},
        {'column': 'N', 'column_name': 'MAC Add', 'keys': ['address']},
        {'column': 'O', 'column_name': 'IP Add', 'keys': ['ip_address']},
        {'column': 'P', 'column_name': 'MTU', 'keys': ['mtu']},
        {'column': 'Q', 'column_name': 'Duplex', 'keys': ['duplex']},
        {'column': 'R', 'column_name': 'Speed', 'keys': ['speed']},
        {'column': 'S', 'column_name': 'Bandwidth', 'keys': ['bandwidth']},
        {'column': 'T', 'column_name': 'Delay', 'keys': ['delay']},
        {'column': 'U', 'column_name': 'Encapsulation', 'keys': ['encapsulation']},
        {'column': 'V', 'column_name': 'Last Input', 'keys': ['last_input']},
        {'column': 'W', 'column_name': 'Last Output', 'keys': ['last_output']},
        {'column': 'X', 'column_name': 'Queue Strategy', 'keys': ['queue_strategy']},
        {'column': 'Y', 'column_name': 'Input Rate', 'keys': ['input_rate']},
        {'column': 'Z', 'column_name': 'Output Rate', 'keys': ['output_rate']},
        {'column': 'AA', 'column_name': 'Input Packets', 'keys': ['input_packets']},
        {'column': 'AB', 'column_name': 'Output Packets', 'keys': ['output_packets']},
        {'column': 'AC', 'column_name': 'Input Errors', 'keys': ['input_errors']},
        {'column': 'AD', 'column_name': 'Output Errors', 'keys': ['output_errors']},
    ],
    'gather_route': [
        {'column': 'B', 'column_name': 'VRF', 'keys': ['vrf']},
        {'column': 'C', 'column_name': 'Protocol', 'keys': ['protocol']},
        {'column': 'D', 'column_name': 'Route', 'keys': ['network']},
        {'column': 'E', 'column_name': 'Subnet', 'keys': ['mask']},
        {'column': 'F', 'column_name': 'CIDR', 'keys': ['cidr']},
        {'column': 'G', 'column_name': 'Next Hop IP', 'keys': ['nexthop_ip']},
        {'column': 'H', 'column_name': 'Next Hop IF', 'keys': ['nexthop_if']},
        {'column': 'I', 'column_name': 'Distance', 'keys': ['distance']},
        {'column': 'J', 'column_name': 'Metric', 'keys': ['metric']},
        {'column': 'K', 'column_name': 'Uptime', 'keys': ['uptime']},
    ],
    'gather_bgp': [
        {'column': 'B', 'column_name': 'Status', 'keys': ['status']},
        {'column': 'C', 'column_name': 'Path Selection', 'keys': ['path_selection']},
        {'column': 'D', 'column_name': 'Route Source', 'keys': ['route_source']},
        {'column': 'E', 'column_name': 'Network', 'keys': ['network']},
        {'column': 'F', 'column_name': 'Next Hop', 'keys': ['next_hop']},
        {'column': 'G', 'column_name': 'Metric', 'keys': ['metric']},
        {'column': 'H', 'column_name': 'Local Preference', 'keys': ['local_pref']},
        {'column': 'I', 'column_name': 'Weight', 'keys': ['weight']},
        {'column': 'J', 'column_name': 'AS Path', 'keys': ['as_path']},
        {'column': 'K', 'column_name': 'Origin', 'keys': ['origin']},
    ],
    'gather_ap': [
        {'column': 'B', 'column_name': 'AP Name', 'keys': ['ap_name']},
        {'column': 'C', 'column_name': 'AP Model', 'keys': ['ap_model']},
        {'column': 'D', 'column_name': 'MAC', 'keys': ['mac']},
        {'column': 'E', 'column_name': 'Serial Number', 'keys': ['serial_number']},
        {'column': 'F', 'column_name': 'Software Version', 'keys': ['version']},
        {'column': 'G', 'column_name': 'Image', 'keys': ['image']},
        {'column': 'H', 'column_name': 'IP Address', 'keys': ['ip']},
        {'column': 'I', 'column_name': 'Netmask', 'keys': ['netmask']},
        {'column': 'J', 'column_name': 'Gateway', 'keys': ['gateway']},
        {'column': 'K', 'column_name': 'Clients', 'keys': ['clients']},
        {'column': 'L', 'column_name': 'AP Group', 'keys': ['ap_group']},
        {'column': 'M', 'column_name': 'Mode', 'keys': ['mode']},
        {'column': 'N', 'column_name': 'Primary Controller IP', 'keys': ['primary_switch_ip']},
        {'column': 'O', 'column_name': 'Primary Controller Name', 'keys': ['primary_switch_name']},
        {'column': 'P', 'column_name': 'Secondary Controller IP', 'keys': ['secondary_switch_ip']},
        {'column': 'Q', 'column_name': 'Secondary Controller Name', 'keys': ['secondary_switch_name']},
        {'column': 'R', 'column_name': 'Tertiary Controller IP', 'keys': ['tertiary_switch_ip']},
        {'column': 'S', 'column_name': 'Tertiary Controller Name', 'keys': ['tertiary_switch_name']},
        {'column': 'T', 'column_name': 'Location', 'keys': ['location']},
        {'column': 'U', 'column_name': 'Country', 'keys': ['country']},
        {'column': 'V', 'column_name': 'Uptime', 'keys': ['uptime']},
        {'column': 'W', 'column_name': 'Join Date Time', 'keys': ['join_date_time']},
        {'column': 'X', 'column_name': 'Join Taken Time', 'keys': ['join_taken_time']},
    ],
    'gather_ip_mroute': [
        {'column': 'B', 'column_name': 'Multicast Source IP', 'keys': ['no_data', 'multicast_source_ip']},
        {'column': 'C', 'column_name': 'Multicast Group IP', 'keys': ['multicast_group_ip']},
        {'column': 'D', 'column_name': 'Up Time', 'keys': ['up_time']},
        {'column': 'E', 'column_name': 'Expiration Time', 'keys': ['expiration_time']},
        {'column': 'F', 'column_name': 'Rendezvous Point', 'keys': ['rendezvous_point']},
        {'column': 'G', 'column_name': 'Flags', 'keys': ['flags']},
        {'column': 'H', 'column_name': 'Incoming Interface', 'keys': ['incoming_interface']},
        {'column': 'I', 'column_name': 'Reverse Path Neighbouring Ip', 'keys': ['reverse_path_forwarding_neighbour_ip']},
        {'column': 'J', 'column_name': 'Registering', 'keys': ['registering']},
        {'column': 'K', 'column_name': 'Outgoing Interface', 'keys': ['outgoing_interface']},
        {'column': 'L', 'column_name': 'Forward Mode', 'keys': ['forward_mode']},
        {'column': 'M', 'column_name': 'Outgoing Multicast Uptime', 'keys': ['outgoing_multicast_uptime']},
        {'column': 'N', 'column_name': 'Outgoing Multicast Expiration Time',
         'keys': ['outgoing_multicast_expiration_time']},
    ]
}
# This will be the output map for the Main page.
# This data will be filled in from the class variable device_info
# The below mapper will only work with values that are string or list
DEVICE_INFO_MAP = [
    {'column': 'H', 'keys': ['start_time']},
    {'column': 'J', 'keys': ['hostname']},
    {'column': 'K', 'keys': ['hardware', 'pid', 'platform']},
    {'column': 'L', 'keys': ['serial', 'sn']},
    {'column': 'O', 'keys': ['uptime', 'system_up_time']},
    {'column': 'P', 'keys': ['os', 'version', 'product_version']},
    {'column': 'Q', 'keys': ['boot_image', 'running_image']},
    {'column': 'AR', 'keys': ['cpu_5_sec']},
    {'column': 'AS', 'keys': ['cpu_1_min']},
    {'column': 'AT', 'keys': ['cpu_5_min']},
    {'column': 'AU', 'keys': ['cpu_15_min']},
    {'column': 'AV', 'keys': ['sfp_count']},
    {'column': 'AW', 'keys': ['elapsed_time']},
    {'column': 'I', 'keys': ['device_type']},
]

INTERACE_COUNT_MAP = {
    "Ethernet": {"count": 20, "active": 21},
    "FastEthernet": {"count": 22, "active": 23},
    "GigabitEthernet": {"count": 24, "active": 25},
    "TenGigabitEthernet": {"count": 26, "active": 27},
    "TwentyFiveGigEthernet": {"count": 28, "active": 29},
    "FortyGigEthernet": {"count": 30, "active": 31},
    "HundredGigEthernet": {"count": 32, "active": 33},
    "Serial": {"count": 18, "active": 19},
    "Subinterfaces": {"count": 34, "active": 35},
    "Tunnel": {"count": 36, "active": 37},
    "Port-channel": {"count": 38, "active": 39},
    "Loopback": {"count": 40, "active": 41},
    "Vlan": {"count": 42, "active": 43}
}
