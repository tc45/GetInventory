# Import openpyxl module
import openpyxl
import os
from netmiko import ConnectHandler
from netmiko import SSHDetect
import json
import socket
import re
import logging
from datetime import datetime

DEBUG = True

# GLOBAL VARIABLES
# xls_input_file = "D:\\Data\\My Documents\\Projects\\ParseIT\\ParseIT - Default.xlsx"
xls_input_file = "D:\\Data\\My Documents\\Projects\\ParseIT\\ParseIT - LL.xlsx"
device_row_start = 0
current_row = 0
device_list = []
device_type = ""
wb_obj = None
sheet_obj = None
conn = ""
username, password, secret, file_output = "", "", "", ""
xls_main_row_username, xls_main_row_password, xls_row_error_current, file_name = "", "", 0, ""
xls_col_main_hostname, xls_col_main_protocol, xls_col_main_port, xls_col_main_type, xls_col_main_ios, \
xls_col_main_uptime = "", "", "", "", "", ""
xls_col_main_parse, xls_col_main_connerror, command_list, current_hostname, json_output = "", "", "", "", ""
xls_col_main_output_dir, xls_col_main_command_output, xls_col_main_json_output = "", "", ""
xls_col_main_username, xls_col_main_password, xls_col_main_collection_time, xls_col_main_model = "", "", "", ""
xls_col_main_serial, xls_col_main_flash, xls_col_main_memory, xls_col_main_active = "", "", "", ""
xls_col_serial_if, xls_col_eth_if, xls_col_fe_if, xls_col_ge_if = "", "", "", ""
xls_col_te_if, xls_col_tfge_if, xls_col_fge_if, xls_col_hunge_if = "", "", "", ""
xls_col_serial_if_active, xls_col_eth_if_active, xls_col_fe_if_active, xls_col_ge_if_active = "", "", "", ""
xls_col_te_if_active, xls_col_tfge_if_active, xls_col_fge_if_active, xls_col_hunge_if_active = "", "", "", ""
xls_col_subif, xls_col_subif_active, xls_col_vlan_if, xls_col_vlan_if_active = "", "", "", ""
xls_col_tunnel_if, xls_col_tunnel_if_active, xls_col_port_chl_if, xls_col_port_chl_if_active, \
    xls_col_loop_if, xls_col_loop_if_active = "", "", "", "", "", ""
xls_col_sfp_count, xls_col_cpu_one, xls_col_cpu_five = "", "", ""
xls_col_routes_hostname, xls_col_routes_protocol, xls_col_routes_metric, xls_col_routes_route = "", "", "", ""
xls_col_routes_subnet, xls_col_routes_cidr, xls_col_routes_nexthopip, xls_col_routes_nexthopif = "", "", "", ""
xls_col_routes_distance, xls_col_routes_uptime = "", ""
xls_col_cdp_hostname, xls_col_cdp_local_port, xls_col_cdp_remote_port = "", "", ""
xls_col_cdp_remote_host, xls_col_cdp_mgmt_ip, xls_col_cdp_software, xls_col_cdp_platform = "", "", "", ""
xls_col_cdp_if_ip, xls_col_cdp_capabilities = "", ""
xls_col_if_hostname, xls_col_if_interface, xls_col_if_link_status, xls_col_if_protocol_status = "", "", "", ""
xls_col_if_l2_l3, xls_col_if_trunk_access, xls_col_if_access_vlan = "", "", ""
xls_col_if_trunk_allowed, xls_col_if_trunk_forwarding = "", ""
xls_col_if_mac_address, xls_col_if_ip_address, xls_col_if_desc, xls_col_if_mtu, xls_col_if_duplex = "", "", "", "", ""
xls_col_if_speed, xls_col_if_bw, xls_col_if_delay, xls_col_if_encapsulation, xls_col_if_last_in = "", "", "", "", ""
xls_col_if_last_out, xls_col_if_queue, xls_col_if_in_rate, xls_col_if_out_rate, xls_col_if_in_pkts = "", "", "", "", ""
xls_col_if_out_pkts, xls_col_if_in_err, xls_col_if_out_err, xls_col_if_short_if = "", "", "", ""
xls_col_if_trunk_native = ""
xls_col_mac_dest_add, xls_col_mac_type, xls_col_mac_vlan, xls_col_mac_dest_port = "", "", "", ""
xls_col_log_date, xls_col_log_time, xls_col_log_timezone, xls_col_log_facility, \
    xls_col_log_severity, xls_col_log_mnemonic, xls_col_log_message = "", "", "", "", "", "", ""
xls_col_arp_ip, xls_col_arp_age, xls_col_arp_mac, xls_col_arp_type, xls_col_arp_if = "", "", "", "", ""
xls_col_arp_vrf, xls_col_if_vrf, xls_col_routes_vrf, xls_col_if_type = "", "", "", ""
xls_col_error_device, xls_col_error_time, xls_col_error_message = "", "", ""

os.environ["NET_TEXTFSM"] = str("ntc-templates\\templates")


def main():

    # if DEBUG is True:
    #    logging.basicConfig(filename="NETMIKO_LOG.txt", level=logging.DEBUG)
    #    logger = logging.getLogger("netmiko")

    # open XLS file
    open_xls()
    # Get hostnames from XLS file
    # Must run get_devices FIRST before getting other info.  This step indexes beginning row and adds device
    #    hostname/IP addresses into a list for further use.
    get_devices()
    # Get settings from XLS file
    get_settings()
    # Get column headers for data on Main tab of XLS page (index column numbers)
    get_column_headers()
    # Get commands from XLS spreadsheet on commands tab
    get_commands()
    # Connect to device list
    connect_devices()
    print("/n/n/nBatch job completed.")

# Reference sheet lookup
# sheet_obj.cell(row=current_row, column=xls_col_connerror).value = str(e)


def probe_port(device, port):
    if isOpen(device, 22):
        return "ssh"
    elif isOpen(device, 23):
        return "telnet"
    else:
        return "Unknown"


def open_xls():
    global wb_obj, sheet_obj
    wb_obj = openpyxl.load_workbook(xls_input_file)
    sheet_obj = wb_obj['Main']


def save_xls():
    global wb_obj
    if DEBUG is True:
        print("Saving XLS workbook now - " + file_output + file_name)
    wb_obj.save(file_output + file_name)


def get_settings():
    global username, password, secret, file_output, file_name

    for i in range(1, device_row_start - 1):
        cell_value = rw_cell(i, 1)
        if cell_value == "Username":
            username = rw_cell(i, 2)
            if DEBUG is True:
                print("Username set to " + username)
        elif cell_value == "Password":
            password = rw_cell(i, 2)
            if DEBUG is True:
                print("Password has been set to *******")
        elif cell_value == "Secret":
            secret = rw_cell(i, 2)
            if DEBUG is True:
                print("Secret has been set to ********")
        elif cell_value == "Output Directory":
            file_output = rw_cell(i, 2)
            if DEBUG is True:
                print("File output directory set to " + file_output)
        elif cell_value == "Output Name":
            file_name = rw_cell(i, 2)
            if DEBUG is True:
                print("File output name set to " + file_name)

    if right(file_output, 1) != "\\":
        if DEBUG is True:
            print("File path didn't end in backslash.")
        file_output = file_output + "\\"

    if right(file_name, 3) != "xls" or right(file_name, 4) != "xlsx":
        if DEBUG is True:
            print("File path didn't end in XLS.")
        file_name = file_name + ".xlsx"

    if DEBUG is True:
        print("Output file will be stored as " + file_output + file_name)


def get_column_headers():
    global xls_col_main_protocol, xls_col_main_port, xls_col_main_type, xls_col_main_hostname, xls_col_main_ios, xls_col_main_uptime, \
        xls_col_main_connerror, xls_col_main_output_dir, xls_col_main_command_output, xls_col_main_json_output, \
        xls_col_routes_cidr, xls_col_routes_distance, xls_col_routes_hostname, xls_col_routes_metric, \
        xls_col_routes_nexthopif, xls_col_routes_nexthopip, xls_col_routes_protocol, xls_col_routes_route, \
        xls_col_routes_subnet, xls_col_routes_uptime, xls_col_cdp_hostname, xls_col_cdp_local_port, \
        xls_col_cdp_remote_port, xls_col_cdp_remote_host, xls_col_cdp_mgmt_ip, xls_col_cdp_software, \
        xls_col_cdp_platform, xls_col_cdp_if_ip, xls_col_cdp_capabilities, \
        xls_col_if_hostname, xls_col_if_interface, xls_col_if_link_status, \
        xls_col_if_protocol_status, xls_col_if_l2_l3, xls_col_if_trunk_access, xls_col_if_access_vlan, \
        xls_col_if_trunk_allowed, xls_col_if_trunk_forwarding, xls_col_if_mac_address, xls_col_if_ip_address, \
        xls_col_if_desc, xls_col_if_mtu, xls_col_if_duplex, xls_col_if_speed, xls_col_if_bw, xls_col_if_delay, \
        xls_col_if_encapsulation, xls_col_if_last_in, xls_col_if_last_out, xls_col_if_queue, xls_col_if_in_rate, \
        xls_col_if_out_rate, xls_col_if_in_pkts, xls_col_if_out_pkts, xls_col_if_in_err, xls_col_if_out_err, \
        xls_col_if_short_if, xls_col_if_trunk_native, xls_col_serial_if, xls_col_eth_if, xls_col_fe_if, xls_col_ge_if, \
        xls_col_te_if, xls_col_tfge_if, xls_col_fge_if, xls_col_hunge_if, xls_col_serial_if_active, \
        xls_col_eth_if_active, xls_col_fe_if_active, xls_col_ge_if_active, xls_col_te_if_active,  \
        xls_col_tfge_if_active, xls_col_fge_if_active, xls_col_hunge_if_active, xls_col_sfp_count, xls_col_cpu_one, \
        xls_col_cpu_five, xls_col_main_serial, xls_col_main_flash, xls_col_main_memory, xls_col_main_active, xls_col_main_username, \
        xls_col_main_password, xls_col_main_collection_time, xls_col_main_model, xls_col_mac_dest_add, xls_col_mac_type, \
        xls_col_mac_vlan, xls_col_mac_dest_port, xls_col_log_date, xls_col_log_time, xls_col_log_timezone, \
        xls_col_log_facility, xls_col_log_severity, xls_col_log_mnemonic, xls_col_arp_ip, xls_col_arp_age, \
        xls_col_arp_mac, xls_col_arp_type, xls_col_arp_if, xls_col_log_message, xls_col_subif, xls_col_subif_active, \
        xls_col_main_parse, xls_col_arp_vrf, xls_col_if_vrf, xls_col_routes_vrf, xls_col_if_type, xls_col_error_device, \
        xls_col_error_message, xls_col_error_time, xls_col_tunnel_if, xls_col_tunnel_if_active, xls_col_port_chl_if, \
        xls_col_port_chl_if_active, xls_col_loop_if, xls_col_loop_if_active, xls_col_vlan_if, xls_col_vlan_if_active

    sheet = wb_obj["Main"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet_obj.cell(row=device_row_start - 1, column=i).value
        if cell_value is not None:
            if cell_value == "IP/DNS Host":
                xls_col_main_hostname = i
            elif left(cell_value, 12) == "Parse Method":
                xls_col_main_parse = i
            elif cell_value == "Protocol":
                xls_col_main_protocol = i
            elif cell_value == "Port Override":
                xls_col_main_port = i
            elif cell_value == "Connection Error":
                xls_col_main_connerror = i
            elif cell_value == "Device Type":
                xls_col_main_type = i
            elif cell_value == "Hostname":
                xls_col_main_hostname = i
            elif cell_value == "IOS Version":
                xls_col_main_ios = i
            elif cell_value == "Uptime":
                xls_col_main_uptime = i
            elif cell_value == "Output Directory":
                xls_col_main_output_dir = i
            elif cell_value == "Command Output":
                xls_col_main_command_output = i
            elif cell_value == "JSON Output":
                xls_col_main_json_output = i
            elif cell_value == "Active":
                xls_col_main_active = i
            elif cell_value == "Username":
                xls_col_main_username = i
            elif cell_value == "Password":
                xls_col_main_password = i
            elif cell_value == "Collection Date/Time":
                xls_col_main_collection_time = i
            elif cell_value == "Model":
                xls_col_main_model = i
            elif cell_value == "Serial Number":
                xls_col_main_serial = i
            elif cell_value == "Memory":
                xls_col_main_memory = i
            elif cell_value == "Flash":
                xls_col_main_flash = i
            elif cell_value == "Serial IF":
                xls_col_serial_if = i
            elif cell_value == "Serial IF - Active":
                xls_col_serial_if_active = i
            elif cell_value == "Ethernet IF":
                xls_col_eth_if = i
            elif cell_value == "Ethernet IF - Active":
                xls_col_eth_if_active = i
            elif cell_value == "FastEthernet IF":
                xls_col_fe_if = i
            elif cell_value == "FastEthernet IF - Active":
                xls_col_fe_if_active = i
            elif cell_value == "GigEth IF":
                xls_col_ge_if = i
            elif cell_value == "GigEth IF - Active":
                xls_col_ge_if_active = i
            elif cell_value == "TenGig IF":
                xls_col_te_if = i
            elif cell_value == "TenGig IF - Active":
                xls_col_te_if_active = i
            elif cell_value == "TwentyFiveGig IF":
                xls_col_tfge_if = i
            elif cell_value == "TwentyFiveGig IF - Active":
                xls_col_tfge_if_active = i
            elif cell_value == "FortyGig IF":
                xls_col_fge_if = i
            elif cell_value == "FortyGig IF - Active":
                xls_col_fge_if_active = i
            elif cell_value == "HundredGig IF":
                xls_col_hunge_if = i
            elif cell_value == "HundredGig IF - Active":
                xls_col_hunge_if_active = i
            elif cell_value == "Subinterfaces":
                xls_col_subif = i
            elif cell_value == "Subinterfaces - Active":
                xls_col_subif_active = i
            elif cell_value == "Tunnel IF":
                xls_col_tunnel_if = i
            elif cell_value == "Tunnel IF - Active":
                xls_col_tunnel_if_active = i
            elif cell_value == "Port-Channel IF":
                xls_col_port_chl_if = i
            elif cell_value == "Port-Channel IF - Active":
                xls_col_port_chl_if_active = i
            elif cell_value == "Loopback IF":
                xls_col_loop_if = i
            elif cell_value == "Loopback IF - Active":
                xls_col_loop_if_active = i
            elif cell_value == "VLAN IF":
                xls_col_vlan_if = i
            elif cell_value == "VLAN IF - Active":
                xls_col_vlan_if_active = i
            elif cell_value == "One Min CPU":
                xls_col_cpu_one = i
            elif cell_value == "Five Min CPU":
                xls_col_cpu_five = i
            elif cell_value == "SFP Count":
                xls_col_sfp_count = i

    sheet = wb_obj["Routes"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Hostname":
                xls_col_routes_hostname = i
            elif cell_value == "VRF":
                xls_col_routes_vrf = i
            elif cell_value == "Protocol":
                xls_col_routes_protocol = i
            elif cell_value == "Metric":
                xls_col_routes_metric = i
            elif cell_value == "Route":
                xls_col_routes_route = i
            elif cell_value == "Subnet":
                xls_col_routes_subnet = i
            elif cell_value == "CIDR":
                xls_col_routes_cidr = i
            elif cell_value == "Next Hop IP":
                xls_col_routes_nexthopip = i
            elif cell_value == "Next Hop IF":
                xls_col_routes_nexthopif = i
            elif cell_value == "Distance":
                xls_col_routes_distance = i
            elif cell_value == "Metric":
                xls_col_routes_metric = i
            elif cell_value == "Uptime":
                xls_col_routes_uptime = i


    sheet = wb_obj["CDP"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Hostname":
                xls_col_cdp_hostname = i
            elif cell_value == "Local Port":
                xls_col_cdp_local_port = i
            elif cell_value == "Remote Host":
                xls_col_cdp_remote_host = i
            elif cell_value == "Remote Port":
                xls_col_cdp_remote_port = i
            elif cell_value == "Interface IP":
                xls_col_cdp_if_ip = i
            elif cell_value == "MGMT IP":
                xls_col_cdp_mgmt_ip = i
            elif cell_value == "Platform":
                xls_col_cdp_platform = i
            elif cell_value == "Software":
                xls_col_cdp_software = i
            elif cell_value == "Capabilities":
                xls_col_cdp_capabilities = i

    sheet = wb_obj["Interfaces"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Hostname":
                xls_col_if_hostname = i
            elif cell_value == "Interface":
                xls_col_if_interface = i
            elif cell_value == "Short IF":
                xls_col_if_short_if = i
            elif cell_value == "Description":
                xls_col_if_desc = i
            elif cell_value == "Type":
                xls_col_if_type = i
            elif cell_value == "VRF":
                xls_col_if_vrf = i
            elif cell_value == "Link":
                xls_col_if_link_status = i
            elif cell_value == "Protocol":
                xls_col_if_protocol_status = i
            elif cell_value == "L2/L3":
                xls_col_if_l2_l3 = i
            elif cell_value == "Trunk/Access":
                xls_col_if_trunk_access = i
            elif cell_value == "Access VLAN":
                xls_col_if_access_vlan = i
            elif cell_value == "Trunk Allowed":
                xls_col_if_trunk_allowed = i
            elif cell_value == "Trunk Forwarding":
                xls_col_if_trunk_forwarding = i
            elif cell_value == "Native VLAN":
                xls_col_if_trunk_native = i
            elif cell_value == "MAC Add":
                xls_col_if_mac_address = i
            elif cell_value == "IP Add":
                xls_col_if_ip_address = i
            elif cell_value == "MTU":
                xls_col_if_mtu = i
            elif cell_value == "Duplex":
                xls_col_if_duplex = i
            elif cell_value == "Speed":
                xls_col_if_speed = i
            elif cell_value == "BW":
                xls_col_if_bw = i
            elif cell_value == "Delay":
                xls_col_if_delay = i
            elif cell_value == "Encap":
                xls_col_if_encapsulation = i
            elif cell_value == "Last Input":
                xls_col_if_last_in = i
            elif cell_value == "Last Output":
                xls_col_if_last_out = i
            elif cell_value == "Queue Strategy":
                xls_col_if_queue = i
            elif cell_value == "Input Rate":
                xls_col_if_in_rate = i
            elif cell_value == "Output Rate":
                xls_col_if_out_rate = i
            elif cell_value == "Input Packets":
                xls_col_if_in_pkts = i
            elif cell_value == "Output Packets":
                xls_col_if_out_pkts = i
            elif cell_value == "Input Errors":
                xls_col_if_in_err = i
            elif cell_value == "Output Errors":
                xls_col_if_out_err = i

    sheet = wb_obj["ARP"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "IP Address":
                xls_col_arp_ip = i
            elif cell_value == "VRF":
                xls_col_arp_vrf = i
            elif cell_value == "Age":
                xls_col_arp_age = i
            elif cell_value == "Hardware/MAC":
                xls_col_arp_mac = i
            elif cell_value == "Type":
                xls_col_arp_type = i
            elif cell_value == "Interface":
                xls_col_arp_if = i

    sheet = wb_obj["MAC Tables"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Destination Address":
                xls_col_mac_dest_add = i
            elif cell_value == "Type":
                xls_col_mac_type = i
            elif cell_value == "VLAN":
                xls_col_mac_vlan = i
            elif cell_value == "Destination Port":
                xls_col_mac_dest_port = i

    sheet = wb_obj["Logging"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Date":
                xls_col_log_date = i
            elif cell_value == "Time":
                xls_col_log_time = i
            elif cell_value == "Timezone":
                xls_col_log_timezone = i
            elif cell_value == "Facility":
                xls_col_log_facility = i
            elif cell_value == "Severity":
                xls_col_log_severity = i
            elif cell_value == "Mnemonic":
                xls_col_log_mnemonic = i
            elif cell_value == "Message":
                xls_col_log_message = i

    sheet = wb_obj["Errors"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Hostname":
                xls_col_error_device = i
            elif cell_value == "Time":
                xls_col_error_time = i
            elif cell_value == "Error":
                xls_col_error_message = i


def get_commands():
    global command_list
    sheet = wb_obj["Commands"]
    max_row = sheet.max_row
    for i in range(1, max_row + 1):
        command = sheet.cell(row=i, column=1).value
        if command != "":
            command_list = command_list + command
            if i < max_row:
                command_list = command_list + ","

    command_list = command_list.split(",")

    if DEBUG is True:
        print("The following commands were found on the 'Commands' tab")
        for x in range(len(command_list)):
            print(str(x + 1) + " - " + command_list[x])


def get_devices():
    global device_row_start, device_list, current_row
    xls_rows_total = sheet_obj.max_row

    for i in range(1, xls_rows_total + 1):
        cell_value = rw_cell(i, 1)
        if device_row_start > 0:
            if cell_value is not None:
                device_list.append(cell_value)
        if cell_value == "IP/DNS Host":
            device_row_start = i + 1
            current_row = device_row_start
    if DEBUG is True:
        print("Total rows in this sheet is " + str(xls_rows_total))
        print("Devices found in spreadsheet:")
        for i in range(0, len(device_list)):
            print(str(i + 1) + ": " + device_list[i])
        print('\n')


def set_protocol(device):

    port = rw_cell(current_row, xls_col_main_port)
    protocol = rw_cell(current_row, xls_col_main_protocol)

    if port is None:
        if DEBUG is True:
            print("No Protocol found for device " + device + ".")
        if protocol == "telnet":
            if DEBUG is True:
                print("Setting protocol to telnet.")
            port = 23
        else:
            if DEBUG is True:
                print("Setting protocol to ssh.")
            port = 22

    if protocol is None:
        protocol = "ssh"

    if DEBUG is True:
        print("Protocol is currently set to: " + str(protocol))

    return [port, protocol]


def connect_devices():
    global conn, current_row, current_hostname, json_output, file_output
    json_file, commands_file, commands = "", "", ""
    device = {}

    for i in device_list:
        conn_type = rw_cell(current_row, xls_col_main_parse, False, "", "Main")
        if conn_type == "autodetect":
            conn_type = guess_os(i, username, password, secret)

        # if conn_type is original not set, or comes back as none, set it to cisco_ios
        if conn_type is None:
            conn_type = "cisco_ios"

        rw_cell(current_row, xls_col_main_parse, True, conn_type, "Main")

        # Reset JSON and command string for each device
        json_output = ""
        commands = ""

        # Check if device has protocol and ports associated.  If not assume SSH port 22.
        ports = set_protocol(i)

        conn_port = ports[0]
        conn_protocol = ports[1]

        print('\n' + '\n' + '\n' + "Connecting to device " + i + " on port " + str(conn_port) +
            " using protocol " + str(conn_protocol) + "." + '\n')

        device = {
            'device_type': conn_type + "_" + conn_protocol,
            'ip': i,
            'username': username,
            'password': password,
            'secret': secret,
            'port': conn_port
        }

        if DEBUG is True:
            device = {
                'device_type': conn_type + "_" + conn_protocol,
                'ip': i,
                'username': username,
                'password': password,
                'secret': secret,
                'port': conn_port,
                'verbose': True
            }

        try:
            conn = ConnectHandler(**device)
            conn.enable()
            command = "term len 0"

            output = conn.send_command(command)
            if DEBUG is True:
                print("Set terminal length to 0 for this session.")
        except Exception as e:
            rw_cell(current_row, xls_col_main_connerror, True, str(e))
            write_error(i, str(e))
        else:
            try:
                # Run all JSON related output here.
                show_version(i)
            except Exception as e:
                print("ERROR - Show version failed due to error: " + str(e))
                write_error(current_hostname, "ERROR - Show version failed due to error: " + str(e))
            try:
                show_interfaces(i, conn_type)
            except Exception as e:
                print("ERROR - Show interfaces failed due to error: " + str(e))
                write_error(current_hostname, "ERROR - Show interfaces failed due to error: " + str(e))
            try:
                show_ip_route(i)
            except Exception as e:
                print("ERROR - Show ip route failed due to error: " + str(e))
                write_error(current_hostname, "ERROR - Show ip route failed due to error: " + str(e))
            try:
                show_cdp_neighbor(i, conn_type)
            except Exception as e:
                print("ERROR - Show cdp neighbor failed due to error: " + str(e))
                write_error(current_hostname, "ERROR - Show cdp neighbor failed due to error: " + str(e))
            try:
                # Run commands to send to text file
                commands = show_commands(i)
            except Exception as e:
                print("ERROR - Show multiple commands failed due to error: " + str(e))
                write_error(current_hostname, "ERROR - Show multiple commands failed due to error: " + str(e))
            # Grab additional JSON data
            show_inventory(i)
            show_ip_arp(i, conn_type)
            show_mac_address_table(i, conn_type)
            show_logging(i)
            # show_proc_memory(i)
            # show_proc_cpu(i)
            try:
                # Write commands returned from function to text file.
                commands_file = current_hostname + "-commands.txt"
                write_file(file_output + commands_file, commands, False)
                # Write JSON File for each device
                json_file = current_hostname + "-JSON-commands.txt"
                write_file(file_output + json_file, json_output, False)
            except Exception as e:
                write_error(current_hostname, "ERROR - Writing commands to file failed: " + str(e))

            # Write unique device data to spreadsheet
            rw_cell(current_row, xls_col_main_protocol, True, conn_protocol)
            rw_cell(current_row, xls_col_main_port, True, conn_port)
            rw_cell(current_row, xls_col_main_output_dir, True, file_output)
            rw_cell(current_row, xls_col_main_command_output, True, commands_file)
            rw_cell(current_row, xls_col_main_json_output, True, json_file)

            conn.disconnect()

        current_row = current_row + 1
        # Save XLS file after device completed
        save_xls()


def show_commands(device):
    global command_list
    command_output = "------------------------------------------------------------" + "\n" + \
                     "------------------------------------------------------------" + "\n" + \
                     "                 Connected to " + device + "\n" + \
                     "------------------------------------------------------------" + "\n" + \
                     "------------------------------------------------------------" + "\n" + \
                     "The following commands will be executed" + "\n"

    for x in range(len(command_list)):
        command_output = command_output + command_list[x] + "\n"

    for x in range(len(command_list)):

        output = conn.send_command(command_list[x])
        command_output = command_output + wrap_command(command_list[x], output)

    if DEBUG is True:
        print(command_output)

    return command_output


def write_file(filename, file_data, append=False):
    file = ""

    if append is True:
        if os.path.exists(filename):
            append = True
        else:
            append = False

    if append is False:
        file = open(filename, "w+")
    elif append is True:
        file = open(filename, "a+")

    if DEBUG is True:
        print("Writing file  " + filename + "' to disk.")

    file.write(file_data)
    file.close()


def show_version(current_device):
    global device_type, current_hostname, json_output
    file_data = ""
    if DEBUG is True:
        print("Starting show version for device " + current_device)
    command = "show version"

    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        # print(json.dumps(output, indent=2))
        print(string_output)

    rw_cell(current_row, xls_col_main_collection_time, True, get_current_time("dt"), "Main")

    current_hostname = output[0]['hostname']
    rw_cell(current_row, xls_col_main_hostname, True, output[0]['hostname'])
    rw_cell(current_row, xls_col_main_ios, True, output[0]['running_image'])

    if output[0]['serial']:
        serial_len = len(output[0]['serial'])
        if serial_len > 1:
            rw_cell(current_row, xls_col_main_serial, True, output[0]['serial'])
        elif serial_len == 1:
            rw_cell(current_row, xls_col_main_serial, True, output[0]['serial'][0])
    if output[0]['hardware']:
        hw_len = len(output[0]['hardware'])
        if hw_len > 1:
            rw_cell(current_row, xls_col_main_model, True, output[0]['hardware'])
        elif hw_len == 1:
            rw_cell(current_row, xls_col_main_model, True, output[0]['hardware'][0])
    # rw_cell(current_row, xls_col_last_reload, True, output[0]['reload_reason'])
    rw_cell(current_row, xls_col_main_uptime, True, format_uptime(output[0]['uptime']))

    if DEBUG is True:
        print("///// ENDING show version for device " + current_device + "/////")

    file_data = wrap_command(command, string_output)
    json_output = json_output + file_data

    try:
        conn.send_command('show interface switchport', use_textfsm=True)
        device_type = "Switch"
        rw_cell(current_row, xls_col_main_type, True, device_type)
    except Exception as e:
        device_type = "Router"
        rw_cell(current_row, xls_col_main_type, True, device_type)


def show_interfaces_old(current_device):
    global json_output
    if DEBUG is True:
        print("Starting show interfaces for device " + current_device)

    command = "show ip interface brief"
    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)

    if DEBUG is True:
        for interface in output:
            if interface['status'] == "administratively down":
                print(f"{interface['intf']} is ADMIN DOWN!")

    if DEBUG is True:
        print ("///// ENDING show interfaces for device " + current_device + "/////")

    file_data = wrap_command(command, string_output)
    json_output = json_output + file_data


def show_inventory(current_device):
    global json_output
    if DEBUG is True:
        print("Starting show inventory for device " + current_device)

    command = "show inventory"
    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)

    if DEBUG is True:
        print ("///// ENDING show inventory for device " + current_device + "/////")

    json_output = json_output + wrap_command(command, string_output)


def guess_os(device, str_username, str_password, str_secret):
    if str_secret == "":
        str_secret = password
    remote_device = {'device_type': 'autodetect',
                     'host': device,
                     'username': str_username,
                     'password': str_password,
                     'secret': str_secret}
    try:
        guesser = SSHDetect(**remote_device)
        best_match = guesser.autodetect()
    except:
        return None
    else:
        print("Device Guesser: " + best_match)
        return best_match


def show_ip_arp(current_device, conn_type):
    global json_output

    if current_device == "172.30.254.111":
        print("Found device 172.30.254.111")
    sheet = wb_obj['ARP']
    max_row = sheet.max_row + 1

    if DEBUG is True:
        print("Starting show ip arp for device " + current_device)

    command = "show ip arp"
    try:
        output = conn.send_command(command, use_textfsm=True)
    except Exception as e:
        if DEBUG is True:
            print("show ip arp could not be parsed" + str(e))
        write_error(current_hostname, "show ip arp could not be parsed" + str(e))
    else:
        string_output = json.dumps(output, indent=2)

        if DEBUG is True:
            print(string_output)

        # Write Routing data to spreadsheet 'ARP' tab
        if isinstance(output, list):
            for arp in output:
                rw_cell(max_row, 1, True, current_hostname, "ARP")
                rw_cell(max_row, xls_col_arp_ip, True, arp['address'], "ARP")
                rw_cell(max_row, xls_col_arp_age, True, arp['age'], "ARP")
                rw_cell(max_row, xls_col_arp_mac, True, arp['mac'], "ARP")
                if conn_type == "cisco_ios":
                    rw_cell(max_row, xls_col_arp_type, True, arp['type'], "ARP")
                elif conn_type == "cisco_nxos":
                    rw_cell(max_row, xls_col_arp_type, True, "ARPA", "ARP")
                rw_cell(max_row, xls_col_arp_if, True, arp['interface'], "ARP")
                max_row = max_row + 1
        else:
            rw_cell(max_row, 1, True, current_hostname, "ARP")
            rw_cell(max_row, xls_col_arp_ip, True, "No ARP Data Found", "ARP")

        if DEBUG is True:
            print ("///// ENDING show ip arp for device " + current_device + "/////")

        json_output = json_output + wrap_command(command, string_output)


def show_mac_address_table(current_device, current_device_type):
    global json_output

    sheet = wb_obj['MAC Tables']
    max_row = sheet.max_row + 1

    if DEBUG is True:
        print("Starting show mac address-table for device " + current_device)

    command = "show mac address-table"
    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)

    # Write Routing data to spreadsheet 'MAC Tables' tab
    if isinstance(output, list):
        for mac in output:
            if current_device_type == "cisco_ios":
                rw_cell(max_row, 1, True, current_hostname, "MAC Tables")
                rw_cell(max_row, xls_col_mac_dest_add, True, str(mac['destination_address']), "MAC Tables")
                rw_cell(max_row, xls_col_mac_type, True, mac['type'], "MAC Tables")
                rw_cell(max_row, xls_col_mac_vlan, True, str(mac['vlan']), "MAC Tables")
                rw_cell(max_row, xls_col_mac_dest_port, True, mac['destination_port'], "MAC Tables")
            elif current_device_type == "cisco_nxos":
                rw_cell(max_row, 1, True, current_hostname, "MAC Tables")
                rw_cell(max_row, xls_col_mac_dest_add, True, str(mac['mac']), "MAC Tables")
                rw_cell(max_row, xls_col_mac_type, True, mac['type'], "MAC Tables")
                rw_cell(max_row, xls_col_mac_vlan, True, str(mac['vlan']), "MAC Tables")
                rw_cell(max_row, xls_col_mac_dest_port, True, mac['ports'], "MAC Tables")
            max_row = max_row + 1
    else:
        if DEBUG is True:
            print("No MAC Address Table results found for device " + current_device + ".")
        write_error(current_hostname, "No MAC Address Table results found for device " + current_device + ".")

    if DEBUG is True:
        print ("///// ENDING show mac address-table for device " + current_device + "/////")

    json_output = json_output + wrap_command(command, string_output)


def show_logging(current_device):
    global json_output

    sheet = wb_obj['Logging']
    max_row = sheet.max_row + 1

    if DEBUG is True:
        print("Starting show logging for device " + current_device)

    command = "show logging"
    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)

    # Write Routing data to spreadsheet 'Logging' tab
    if isinstance(output, list):
        for log in output:
            for message in log["message"]:
                rw_cell(max_row, 1, True, current_hostname, "Logging")
                rw_cell(max_row, xls_col_log_date , True, log['month'] + " - " + log['day'], "Logging")
                rw_cell(max_row, xls_col_log_time, True, log['time'], "Logging")
                rw_cell(max_row, xls_col_log_timezone, True, log['timezone'], "Logging")
                rw_cell(max_row, xls_col_log_facility, True, log['facility'], "Logging")
                rw_cell(max_row, xls_col_log_severity, True, log['severity'], "Logging")
                rw_cell(max_row, xls_col_log_mnemonic, True, log['mnemonic'], "Logging")
                rw_cell(max_row, xls_col_log_message, True, message, "Logging")
                max_row = max_row + 1
    else:
        if DEBUG is True:
            print("No logs found or info could not be parsed for device " + current_device + ".")

    if DEBUG is True:
        print ("///// ENDING show logging for device " + current_device + "/////")

    json_output = json_output + wrap_command(command, string_output)


def show_proc_memory(current_device):
    global json_output
    if DEBUG is True:
        print("Starting show processes memory for device " + current_device)

    command = "show processes memory"
    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)

    if DEBUG is True:
        print ("///// ENDING show processes memory for device " + current_device + "/////")

    json_output = json_output + wrap_command(command, string_output)


def show_proc_cpu(current_device):
    global json_output
    if DEBUG is True:
        print("Starting show processes cpu for device " + current_device)

    command = "show processes cpu"
    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)

    if DEBUG is True:
        print ("///// ENDING show processes cpu for device " + current_device + "/////")

    json_output = json_output + wrap_command(command, string_output)


def show_ip_route(current_device):
    global json_output, xls_col_routes_cidr, xls_col_routes_distance, xls_col_routes_hostname, xls_col_routes_metric, \
        xls_col_routes_nexthopif, xls_col_routes_nexthopip, xls_col_routes_protocol, xls_col_routes_route, \
        xls_col_routes_subnet, xls_col_routes_uptime

    sheet = wb_obj['Routes']
    max_row = sheet.max_row + 1
    command = "show ip route"

    if DEBUG is True:
        print("Starting gathering JSON data for '" + command + "' on " + current_device + ".")

    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)
        print ("///// ENDING gathering JSON data for '" + command + "' on " + current_device + "./////")

    json_output = json_output + wrap_command(command, string_output)

    # Write Routing data to spreadsheet 'Route' tab
    if is_json(output):
        for route in output:
            rw_cell(max_row, xls_col_routes_hostname, True, current_hostname, "Routes")
            rw_cell(max_row, xls_col_routes_protocol, True, route['protocol'], "Routes")
            rw_cell(max_row, xls_col_routes_route, True, route['network'], "Routes")
            rw_cell(max_row, xls_col_routes_subnet, True, route['mask'], "Routes")
            rw_cell(max_row, xls_col_routes_cidr, True, route['network'] + "\\" + route['mask'], "Routes")
            rw_cell(max_row, xls_col_routes_nexthopip, True, route['nexthop_ip'], "Routes")
            rw_cell(max_row, xls_col_routes_nexthopif, True, route['nexthop_if'], "Routes")
            rw_cell(max_row, xls_col_routes_distance, True, route['distance'], "Routes")
            rw_cell(max_row, xls_col_routes_metric, True, route['metric'], "Routes")
            rw_cell(max_row, xls_col_routes_uptime, True, route['uptime'], "Routes")
            max_row = max_row + 1
            #print("New max_row on tab Routes is " + str(max_row))
    else:
        default_gateway = re.search("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", output)
        rw_cell(max_row, xls_col_routes_hostname, True, current_hostname, "Routes")
        rw_cell(max_row, xls_col_routes_protocol, True, "Layer 2 only", "Routes")

        if default_gateway != "":
            rw_cell(max_row, xls_col_routes_nexthopip, True, default_gateway[0], "Routes")
            print("New max_row on tab Routes is " + str(max_row))


def show_cdp_neighbor(current_device, current_device_type):
    global json_output

    sheet = wb_obj['CDP']
    max_row = sheet.max_row + 1
    command = "show cdp neighbor detail"

    if DEBUG is True:
        print("Starting gathering JSON data for '" + command + "' on " + current_device + ".")

    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    if DEBUG is True:
        print(string_output)
        print ("///// ENDING gathering JSON data for '" + command + "' on " + current_device + "./////")

    json_output = json_output + wrap_command(command, string_output)

    # Write Routing data to spreadsheet 'Route' tab
    if is_json(output):
        for cdp in output:
            rw_cell(max_row, xls_col_cdp_hostname, True, current_hostname, "CDP")
            rw_cell(max_row, xls_col_cdp_local_port, True, cdp['local_port'], "CDP")
            rw_cell(max_row, xls_col_cdp_remote_port, True, cdp['remote_port'], "CDP")
            rw_cell(max_row, xls_col_cdp_if_ip, True, cdp['interface_ip'], "CDP")
            rw_cell(max_row, xls_col_cdp_capabilities, True, cdp['capabilities'], "CDP")
            if current_device_type =="cisco_ios":
                rw_cell(max_row, xls_col_cdp_mgmt_ip, True, cdp['management_ip'], "CDP")
                rw_cell(max_row, xls_col_cdp_remote_host, True, cdp['destination_host'], "CDP")
                rw_cell(max_row, xls_col_cdp_software, True, cdp['software_version'], "CDP")
            if current_device_type == "cisco_nxos":
                rw_cell(max_row, xls_col_cdp_mgmt_ip, True, cdp['mgmt_ip'], "CDP")
                rw_cell(max_row, xls_col_cdp_remote_host, True, cdp['dest_host'], "CDP")
                rw_cell(max_row, xls_col_cdp_software, True, cdp['version'], "CDP")
            rw_cell(max_row, xls_col_cdp_platform, True, cdp['platform'], "CDP")

            max_row = max_row + 1
    else:
        rw_cell(max_row, xls_col_cdp_hostname, True, current_hostname, "CDP")
        rw_cell(max_row, xls_col_cdp_local_port, True, "No CDP Data", "CDP")


def show_interfaces(current_device, current_device_type):
    global json_output

        # xls_col_if_hostname, xls_col_if_interface, xls_col_if_link_status, \
        # xls_col_if_protocol_status, xls_col_if_mac_address, xls_col_if_ip_address, xls_col_if_desc, \
        # xls_col_if_mtu, xls_col_if_duplex, xls_col_if_speed, xls_col_if_bw, xls_col_if_delay, \
        # xls_col_if_encapsulation, xls_col_if_last_in, xls_col_if_last_out, xls_col_if_queue, \
        # xls_col_if_in_rate, xls_col_if_out_rate, xls_col_if_in_pkts, xls_col_if_out_pkts, xls_col_if_in_err, \
        # xls_col_if_out_err, xls_col_if_access_vlan, xls_col_if_trunk_allowed, xls_col_if_trunk_forwarding, \
        # xls_col_if_l2_l3, xls_col_if_trunk_access, xls_col_if_short_if, xls_col_if_trunk_native,  \
        # xls_col_serial_if, xls_col_eth_if, xls_col_fe_if, xls_col_ge_if, \
        # xls_col_te_if, xls_col_tfge_if, xls_col_fge_if, xls_col_hunge_if, xls_col_serial_if_active, \
        # xls_col_eth_if_active, xls_col_fe_if_active, xls_col_ge_if_active, xls_col_te_if_active, \
        # xls_col_tfge_if_active, xls_col_fge_if_active, xls_col_hunge_if_active, xls_col_sfp_count, xls_col_cpu_one, \
        # xls_col_cpu_five, xls_col_serial, xls_col_flash, xls_col_memory, xls_col_active, xls_col_username, \
        # xls_col_password, xls_col_collection_time, xls_col_model, xls_col_subif, xls_col_subif_active, \
        # xls_col_tunnel_if, xls_col_tunnel_if_active, xls_col_port_chl_if, xls_col_port_chl_if_active, \
        # xls_col_loop_if, xls_col_loop_if_active

    sheet = wb_obj['Interfaces']
    max_row = sheet.max_row + 1
    command = "show interface"
    command2 = "show interface status"

    if DEBUG is True:
        print("Starting gathering JSON data for '" + command + "' on " + current_device + ".")
        print("Starting gathering JSON data for '" + command2 + "' on " + current_device + ".")

    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    output2 = conn.send_command(command2, use_textfsm=True)
    string_output2 = json.dumps(output2, indent=2)

    if DEBUG is True:
        print(string_output)
        print("///// ENDING gathering JSON data for '" + command + "' on " + current_device + "./////")
        print(string_output2)
        print("///// ENDING gathering JSON data for '" + command2 + "' on " + current_device + "./////")

    json_output = json_output + wrap_command(command, string_output)
    json_output = json_output + wrap_command(command2, string_output2)

    switchport_data_found = False
    if isinstance(output2, list):
        switchport_data_found = True

    # Write Interface data to spreadsheet 'Interfaces' tab
    if isinstance(output, list):
        for i in output:
            short_if_name = get_short_if_name(i['interface'])
            rw_cell(max_row, xls_col_if_hostname, True, current_hostname, "Interfaces")
            rw_cell(max_row, xls_col_if_short_if, True, short_if_name, "Interfaces")
            rw_cell(max_row, xls_col_if_interface, True, i['interface'], "Interfaces")
            rw_cell(max_row, xls_col_if_link_status, True, i['link_status'], "Interfaces")
            if current_device_type == "cisco_ios":
                rw_cell(max_row, xls_col_if_protocol_status, True, i['protocol_status'], "Interfaces")
            if current_device_type == "cisco_nxos":
                rw_cell(max_row, xls_col_if_protocol_status, True, i['admin_state'], "Interfaces")
            if i['ip_address'] != "":
                rw_cell(max_row, xls_col_if_l2_l3, True, "Layer 3", "Interfaces")
            rw_cell(max_row, xls_col_if_mac_address, True, i['address'], "Interfaces")
            rw_cell(max_row, xls_col_if_ip_address, True, i['ip_address'], "Interfaces")
            rw_cell(max_row, xls_col_if_desc, True, i['description'], "Interfaces")
            rw_cell(max_row, xls_col_if_mtu, True, i['mtu'], "Interfaces")
            rw_cell(max_row, xls_col_if_duplex, True, i['duplex'], "Interfaces")
            rw_cell(max_row, xls_col_if_speed, True, i['speed'], "Interfaces")
            rw_cell(max_row, xls_col_if_bw, True, i['bandwidth'], "Interfaces")
            rw_cell(max_row, xls_col_if_delay, True, i['delay'], "Interfaces")
            rw_cell(max_row, xls_col_if_encapsulation, True, i['encapsulation'], "Interfaces")
            rw_cell(max_row, xls_col_if_in_pkts, True, i['input_packets'], "Interfaces")
            rw_cell(max_row, xls_col_if_out_pkts, True, i['output_packets'], "Interfaces")
            rw_cell(max_row, xls_col_if_in_err, True, i['input_errors'], "Interfaces")
            rw_cell(max_row, xls_col_if_out_err, True, i['output_errors'], "Interfaces")
            if current_device_type == "cisco_ios":
                rw_cell(max_row, xls_col_if_last_in, True, i['last_input'], "Interfaces")
                rw_cell(max_row, xls_col_if_last_out, True, i['last_output'], "Interfaces")
                rw_cell(max_row, xls_col_if_queue, True, i['queue_strategy'], "Interfaces")
                rw_cell(max_row, xls_col_if_in_rate, True, i['input_rate'], "Interfaces")
                rw_cell(max_row, xls_col_if_out_rate, True, i['output_rate'], "Interfaces")
            if switchport_data_found is True:
                for x in output2:
                    if short_if_name == x['port']:
                        if x['vlan'].isnumeric():
                            rw_cell(max_row, xls_col_if_access_vlan, True, x['vlan'], "Interfaces")
                            rw_cell(max_row, xls_col_if_trunk_access, True, "Access", "Interfaces")
                            rw_cell(max_row, xls_col_if_l2_l3, True, "Layer 2", "Interfaces")
                        elif x['vlan'] == "trunk":
                            try:
                                trunk_info = get_trunk_info(short_if_name)
                            except:
                                write_error("Error getting trunk info from device" + current_hostname)

                            rw_cell(max_row, xls_col_if_trunk_access, True, "Trunk", "Interfaces")
                            rw_cell(max_row, xls_col_if_trunk_native, True, trunk_info[0], "Interfaces")
                            rw_cell(max_row, xls_col_if_trunk_allowed, True, trunk_info[1],
                                    "Interfaces")
                            rw_cell(max_row, xls_col_if_trunk_forwarding, True, trunk_info[2],
                                    "Interfaces")
                            rw_cell(max_row, xls_col_if_l2_l3, True, "Layer 2", "Interfaces")
                        elif x['vlan'] == "routed":
                            rw_cell(max_row, xls_col_if_trunk_access, True, "Routed", "Interfaces")
            max_row = max_row + 1

    # Count interfaces by type and number of active interfaces
    interface_counts = count_if_details(output)
    for i in interface_counts:
        if i['type'] == "Ethernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_eth_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_eth_if_active, True, i['active'], "Main")
        elif i['type'] == "FastEthernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_fe_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_fe_if_active, True, i['active'], "Main")
        elif i['type'] == "GigabitEthernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_ge_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_ge_if_active, True, i['active'], "Main")
        elif i['type'] == "TenGigEthernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_te_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_te_if_active, True, i['active'], "Main")
        elif i['type'] == "TwentyFiveGigEthernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_tfge_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_tfge_if_active, True, i['active'], "Main")
        elif i['type'] == "FortyGigEthernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_fge_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_fge_if_active, True, i['active'], "Main")
        elif i['type'] == "HundredGigEthernet":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_hunge_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_hunge_if_active, True, i['active'], "Main")
        elif i['type'] == "Serial":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_serial_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_serial_if_active, True, i['active'], "Main")
        elif i['type'] == "Subinterfaces":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_subif, True, i['count'], "Main")
                rw_cell(current_row, xls_col_subif_active, True, i['active'], "Main")
        elif i['type'] == "Tunnel":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_tunnel_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_tunnel_if_active, True, i['active'], "Main")
        elif i['type'] == "Port-channel":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_port_chl_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_port_chl_if_active, True, i['active'], "Main")
        elif i['type'] == "Loopback":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_loop_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_loop_if_active, True, i['active'], "Main")
        elif i['type'] == "VLAN":
            if i['count'] > 0:
                rw_cell(current_row, xls_col_vlan_if, True, i['count'], "Main")
                rw_cell(current_row, xls_col_vlan_if_active, True, i['active'], "Main")


def count_if_details(if_dictionary):
    eth_if, eth_if_active, fast_if, fast_if_active, gig_if, gig_if_active, serial_if = 0, 0, 0, 0, 0, 0, 0
    serial_if_active, ten_if, ten_if_active, tfg_if, tfg_if_active, forty_if = 0, 0, 0, 0, 0, 0
    forty_if_active, hundred_if, hundred_if_active, subinterface, subinterface_active = 0, 0, 0, 0, 0
    tunnel_if, tunnel_if_active, port_channel_if, port_channel_if_active = 0, 0, 0, 0
    loopback_if, loopback_if_active , vlan_if, vlan_if_active = 0, 0, 0, 0


    for i in if_dictionary:
        split_if = i['interface'].split(".")
        if len(split_if) == 1:
            if left(i['interface'], 3).lower() == "eth":
                eth_if = eth_if + 1
                if i['link_status'] == "up":
                    eth_if_active = eth_if_active + 1
            elif left(i['interface'], 3).lower() == "fas":
                fast_if = fast_if + 1
                if i['link_status'] == "up":
                    fast_if_active = fast_if_active + 1
            elif left(i['interface'], 3).lower() == "gig":
                gig_if = gig_if + 1
                if i['link_status'] == "up":
                    gig_if_active = gig_if_active + 1
            elif left(i['interface'], 3).lower() == "ten":
                ten_if = ten_if + 1
                if i['link_status'] == "up":
                    ten_if_active = ten_if_active + 1
            elif left(i['interface'], 3).lower() == "twe":
                tfg_if = tfg_if + 1
                if i['link_status'] == "up":
                    tfg_if_active = tfg_if_active + 1
            elif left(i['interface'], 3).lower() == "for":
                forty_if = forty_if + 1
                if i['link_status'] == "up":
                    forty_if_active = forty_if_active + 1
            elif left(i['interface'], 3).lower() == "hun":
                hundred_if = hundred_if + 1
                if i['link_status'] == "up":
                    hundred_if_active = hundred_if_active + 1
            elif left(i['interface'], 6).lower() == "serial":
                serial_if = serial_if + 1
                if i['link_status'] == "up":
                    serial_if_active = serial_if_active + 1
            elif left(i['interface'], 3).lower() == "tun":
                tunnel_if = tunnel_if + 1
                if i['link_status'] == "up":
                    tunnel_if_active = tunnel_if_active + 1
            elif left(i['interface'], 5).lower() == "port-":
                port_channel_if = port_channel_if + 1
                if i['link_status'] == "up":
                    port_channel_if_active = port_channel_if_active + 1
            elif left(i['interface'], 3).lower() == "loo":
                loopback_if = loopback_if + 1
                if i['link_status'] == "up":
                    loopback_if_active = loopback_if_active + 1
            elif left(i['interface'], 3).lower() == "vla":
                vlan_if = vlan_if + 1
                if i['link_status'] == "up":
                    vlan_if_active = vlan_if_active + 1
        elif len(split_if) == 2:
            subinterface = subinterface + 1
            if i['link_status'] == "up":
                subinterface_active = subinterface_active + 1

    str_return = [{"type": "Ethernet", "count": eth_if, "active": eth_if_active},
                  {"type": "FastEthernet", "count": fast_if, "active": fast_if_active},
                  {"type": "GigabitEthernet", "count": gig_if, "active": gig_if_active},
                  {"type": "TenGigEthernet", "count": ten_if, "active": ten_if_active},
                  {"type": "TwentyFiveGigEthernet", "count": tfg_if, "active": tfg_if_active},
                  {"type": "FortyGigEthernet", "count": forty_if, "active": forty_if_active},
                  {"type": "HundredGigEthernet", "count": hundred_if, "active": hundred_if_active},
                  {"type": "Serial", "count": serial_if, "active": serial_if_active},
                  {"type": "Subinterfaces", "count": subinterface, "active": subinterface_active},
                  {"type": "Tunnel", "count": tunnel_if, "active": tunnel_if_active},
                  {"type": "Port-channel", "count": port_channel_if, "active": port_channel_if_active},
                  {"type": "Loopback", "count": loopback_if, "active": loopback_if_active},
                  {"type": "VLAN", "count": vlan_if, "active": vlan_if_active}
                  ]

    return str_return


def get_trunk_info(if_name):
    trunk_all_info = conn.send_command("show int trunk").split('\n')
    native_vlan, vlans_allowed, vlans_forwarding = "", "", ""
    x = 0
    for line in trunk_all_info:
        if x == 0:
            if left(line, len(if_name)) == if_name:
                number = re.compile(r"trunking\s+(\d+.*)$")
                native_vlan = number.search(line).group(1)
                x = x + 1
        elif x == 1:
            if left(line, len(if_name)) == if_name:
                number = re.compile(r"\s+(\d.*)$")
                vlans_allowed = number.search(line).group(1)
                x = x + 1
        elif x == 2:
            if line.find("not pruned") != -1:
                x = x + 1
        elif x == 3:
            if left(line, len(if_name)) == if_name:
                if right(line, 1).isnumeric():
                    number = re.compile(r"\s+(\d.*)$")
                elif right(line, 1).isnumeric() is not True:
                    number = re.compile(r"\s+(\w.*)$")
                vlans_forwarding = number.search(line).group(1)
                x = x + 1

    trunk_info = [native_vlan, vlans_allowed, vlans_forwarding]
    return trunk_info


def wrap_command(command, command_data):

    command_output = "------------------------------------------------------------\n" + \
            "*******       " + command + "        *******" + "\n" + \
            "------------------------------------------------------------\n" + \
            "------------------------------------------------------------\n" + \
            "------------------------------------------------------------\n" + \
             command_data + "\n" + \
            "------------------------------------------------------------\n" + \
            "------------------------------------------------------------\n"

    if DEBUG is True:
        print("Wrapping command text for text file for command '" + command + "'.")

    return command_output


def format_uptime(uptime):
    str_year, str_weeks, str_days, str_hours, str_minutes = 0, 0, 0, 0, 0
    str_input = uptime.split(",")
    for i in str_input:
        i = i.strip()
        str_split = i.split(" ")
        if left(str_split[1], 4) == "year":
            str_year = str_split[0]
        if left(str_split[1], 4) == "week":
            str_weeks = str_split[0]
        if left(str_split[1], 3) == "day":
            str_days = str_split[0]
        if left(str_split[1], 4) == "hour":
            str_hours = str_split[0]
        if left(str_split[1], 4) == "minu":
            str_minutes = str_split[0]

    return (str(str_year) + "y " +
            str(str_weeks) + "w " +
            str(str_days) + "d " +
            str(str_hours) + "h " +
            str(str_minutes) + "m "
            )


def write_error(device_name, error_msg):
    global xls_row_error_current

    sheet = wb_obj["Errors"]

    if xls_row_error_current == 0:
        xls_row_error_current = sheet.max_row + 1

    rw_cell(xls_row_error_current, xls_col_error_device, True, device_name, "Errors")
    rw_cell(xls_row_error_current, xls_col_error_time, True, get_current_time("t"), "Errors")
    rw_cell(xls_row_error_current, xls_col_error_message, True, error_msg, "Errors")


def get_current_time(str_option="dt"):
    now = datetime.now()  # Get current date and time
    str_option = str_option.lower()

    time = now.strftime("%H:%M:%S")
    date = now.strftime("%m/%d/%Y")
    if str_option=="dt":
        return date + ", " + time
    elif str_option=="d":
        return date
    elif str_option=="t":
        return time
    else:
        return "Invalid selection.  Choose d for date, t for time, or dt for date + time."


def get_short_if_name(interface):

    number = re.compile(r"(\d.*)$")
    name = re.compile("([a-zA-Z]+)")

    number = number.search(interface).group(1)
    name = name.search(interface).group(1)

    short_name = left(name, 2)

    if int(left(number, 1)) >= 0 or number is None:
        return short_name + str(number)
    else:
        return short_name


def isOpen(ip, port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect((ip, int(port)))
        s.shutdown(2)
        return True
    except:
        return False


def left(s, amount):
    return s[:amount]


def right(s, amount):
    return s[-amount:]


def mid(s, offset, amount):
    return s[offset:offset+amount]


def rw_cell(row, column, write=False, value="", sheet="Main"):
    global wb_obj
    sheet = wb_obj[sheet]

    if write is False:
        value = sheet.cell(row=row, column=column).value
        return value
    elif write is True:
        sheet.cell(row=row, column=column).value = value


def is_json(myjson):
    try:
        json.loads(json.dumps(myjson))
    except ValueError as e:
        return False
    return True


main()
