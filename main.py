# Import openpyxl module
import openpyxl
import os
from netmiko import ConnectHandler
import json
import socket
import re

DEBUG = True

# GLOBAL VARIABLES
xls_input_file = "D:\\Data\\My Documents\\Projects\\ParseIT\\ParseIT - Default.xlsx"
device_row_start = 0
current_row = 0
device_list = []
device_type = ""
wb_obj = None
sheet_obj = None
conn = ""
username, password, secret, file_output = "", "", "", ""
xls_row_username, xls_row_password, file_name = "", "", ""
xls_col_hostname, xls_col_protocol, xls_col_port, xls_col_type, xls_col_ios, xls_col_uptime = "", "", "", "", "", ""
xls_col_connerror, command_list, current_hostname, json_output = "", "", "", ""
xls_col_output_dir, xls_col_command_output, xls_col_json_output = "", "", ""
xls_col_routes_hostname, xls_col_routes_protocol, xls_col_routes_metric, xls_col_routes_route = "", "", "", ""
xls_col_routes_subnet, xls_col_routes_cidr, xls_col_routes_nexthopip, xls_col_routes_nexthopif = "", "", "", ""
xls_col_routes_distance, xls_col_routes_uptime = "", ""
xls_col_cdp_hostname, xls_col_cdp_local_port, xls_col_cdp_remote_port = "", "", ""
xls_col_cdp_remote_host, xls_col_cdp_mgmt_ip, xls_col_cdp_software, xls_col_cdp_platform = "", "", "", ""
xls_col_if_hostname, xls_col_if_interface, xls_col_if_link_status, xls_col_if_protocol_status = "", "", "", ""
xls_col_if_l2_l3, xls_col_if_trunk_access, xls_col_if_access_vlan = "", "", ""
xls_col_if_trunk_allowed, xls_col_if_trunk_forwarding = "", ""
xls_col_if_mac_address, xls_col_if_ip_address, xls_col_if_desc, xls_col_if_mtu, xls_col_if_duplex = "", "", "", "", ""
xls_col_if_speed, xls_col_if_bw, xls_col_if_delay, xls_col_if_encapsulation, xls_col_if_last_in = "", "", "", "", ""
xls_col_if_last_out, xls_col_if_queue, xls_col_if_in_rate, xls_col_if_out_rate, xls_col_if_in_pkts = "", "", "", "", ""
xls_col_if_out_pkts, xls_col_if_in_err, xls_col_if_out_err, xls_col_if_short_if = "", "", "", ""
xls_col_if_trunk_native = ""
os.environ["NET_TEXTFSM"] = str("C:\\Users\\tony\\PycharmProjects\\SecureCRT\\untitled\\ntc-templates\\templates")


def main():
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
    # Save XLS file after all connections made
    save_xls()

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
    global xls_col_protocol, xls_col_port, xls_col_type, xls_col_hostname, xls_col_ios, xls_col_uptime, \
        xls_col_connerror, xls_col_output_dir, xls_col_command_output, xls_col_json_output, \
        xls_col_routes_cidr, xls_col_routes_distance, xls_col_routes_hostname, xls_col_routes_metric, \
        xls_col_routes_nexthopif, xls_col_routes_nexthopip, xls_col_routes_protocol, xls_col_routes_route, \
        xls_col_routes_subnet, xls_col_routes_uptime, xls_col_cdp_hostname, xls_col_cdp_local_port, \
        xls_col_cdp_remote_port, xls_col_cdp_remote_host, xls_col_cdp_mgmt_ip, xls_col_cdp_software, \
        xls_col_cdp_platform, xls_col_if_hostname, xls_col_if_interface, xls_col_if_link_status, \
        xls_col_if_protocol_status, xls_col_if_l2_l3, xls_col_if_trunk_access, xls_col_if_access_vlan, \
        xls_col_if_trunk_allowed, xls_col_if_trunk_forwarding, xls_col_if_mac_address, xls_col_if_ip_address, \
        xls_col_if_desc, xls_col_if_mtu, xls_col_if_duplex, xls_col_if_speed, xls_col_if_bw, xls_col_if_delay, \
        xls_col_if_encapsulation, xls_col_if_last_in, xls_col_if_last_out, xls_col_if_queue, xls_col_if_in_rate, \
        xls_col_if_out_rate, xls_col_if_in_pkts, xls_col_if_out_pkts, xls_col_if_in_err, xls_col_if_out_err, \
        xls_col_if_short_if, xls_col_if_trunk_native

    for i in range(1, sheet_obj.max_column + 1):
        cell_value = sheet_obj.cell(row=device_row_start - 1, column=i).value
        if cell_value != "":
            if cell_value == "Hostname":
                xls_col_hostname = i
            elif cell_value == "Protocol":
                xls_col_protocol = i
            elif cell_value == "Port Override":
                xls_col_port = i
            elif cell_value == "Connection Error":
                xls_col_connerror = i
            elif cell_value == "Device Type":
                xls_col_type = i
            elif cell_value == "Hostname":
                xls_col_hostname = i
            elif cell_value == "IOS Version":
                xls_col_ios = i
            elif cell_value == "Uptime":
                xls_col_uptime = i
            elif cell_value == "Output Directory":
                xls_col_output_dir = i
            elif cell_value == "Command Output":
                xls_col_command_output = i
            elif cell_value == "JSON Output":
                xls_col_json_output = i

    sheet = wb_obj["Routes"]
    max_column = sheet.max_column

    for i in range(1, max_column + 1):
        cell_value = sheet.cell(row=1, column=i).value
        if cell_value != "":
            if cell_value == "Hostname":
                xls_col_routes_hostname = i
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
            elif cell_value == "MGMT IP":
                xls_col_cdp_mgmt_ip = i
            elif cell_value == "Software":
                xls_col_cdp_software = i
            elif cell_value == "Platform":
                xls_col_cdp_platform = i


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
            device_list.append(cell_value)
        if cell_value == "Hostname":
            device_row_start = i + 1
            current_row = device_row_start
    if DEBUG is True:
        print("Total rows in this sheet is " + str(xls_rows_total))
        print("Devices found in spreadsheet:")
        for i in range(1, len(device_list)):
            print(str(i) + " - " + device_list[i])
        print('\n')


def set_protocol(device):

    port = rw_cell(current_row, xls_col_port)
    protocol = rw_cell(current_row, xls_col_protocol)

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

    for i in device_list:
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
            'device_type': 'cisco_ios_' + conn_protocol,
            'ip': i,
            'username': username,
            'password': password,
            'secret': secret,
            'port': conn_port
        }

        try:
            conn = ConnectHandler(**device)
            conn.enable()
            try:
                # Run all JSON related output here.
                show_version(i)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            try:
                show_interfaces(i)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            try:
                show_ip_route(i)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            try:
                show_cdp_neighbor(i)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            try:
                # Run commands to send to text file
                commands = show_commands(i)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            try:
                # Write commands returned from function to text file.
                commands_file = current_hostname + "-commands.txt"
                write_file(file_output + commands_file, commands, False)
                # Write JSON File for each device
                json_file = current_hostname + "-JSON-commands.txt"
                write_file(file_output + json_file, json_output, False)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            try:
                # Write unique device data to spreadsheet
                rw_cell(current_row, xls_col_protocol, True, conn_protocol)
                rw_cell(current_row, xls_col_port, True, conn_port)
                rw_cell(current_row, xls_col_output_dir, True, file_output)
                rw_cell(current_row, xls_col_command_output, True, commands_file)
                rw_cell(current_row, xls_col_json_output, True, json_file)
            except Exception as e:
                rw_cell(current_row, xls_col_connerror + 5, True, str(e))
            # / END Run Commands
            conn.disconnect()
        except Exception as e:
            rw_cell(current_row, xls_col_connerror, True, str(e))

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

    current_hostname = output[0]['hostname']
    rw_cell(current_row, xls_col_hostname, True, output[0]['hostname'])
    rw_cell(current_row, xls_col_ios, True, output[0]['running_image'])

    if DEBUG is True:
        print("///// ENDING show version for device " + current_device + "/////")

    file_data = wrap_command(command, string_output)
    json_output = json_output + file_data

    try:
        conn.send_command('show interface switchport', use_textfsm=True)
        device_type = "Switch"
        rw_cell(current_row, xls_col_type, True, device_type)
    except Exception as e:
        device_type = "Router"
        rw_cell(current_row, xls_col_type, True, device_type)


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

    file_data = wrap_command(command, string_output)
    json_output = json_output + file_data

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


def show_cdp_neighbor(current_device):
    global json_output, xls_col_cdp_hostname, xls_col_cdp_local_port, xls_col_cdp_remote_port, \
        xls_col_cdp_remote_host, xls_col_cdp_mgmt_ip, xls_col_cdp_software, xls_col_cdp_platform

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

    file_data = wrap_command(command, string_output)
    json_output = json_output + file_data

    # Write Routing data to spreadsheet 'Route' tab
    if is_json(output):
        for cdp in output:
            rw_cell(max_row, xls_col_cdp_hostname, True, current_hostname, "CDP")
            rw_cell(max_row, xls_col_cdp_local_port, True, cdp['local_port'], "CDP")
            rw_cell(max_row, xls_col_cdp_remote_port, True, cdp['remote_port'], "CDP")
            rw_cell(max_row, xls_col_cdp_mgmt_ip, True, cdp['management_ip'], "CDP")
            rw_cell(max_row, xls_col_cdp_remote_host, True, cdp['destination_host'], "CDP")
            rw_cell(max_row, xls_col_cdp_platform, True, cdp['platform'], "CDP")
            rw_cell(max_row, xls_col_cdp_software, True, cdp['software_version'], "CDP")
            max_row = max_row + 1
    else:
        rw_cell(max_row, xls_col_cdp_hostname, True, current_hostname, "CDP")
        rw_cell(max_row, xls_col_cdp_local_port, True, "No CDP Data", "CDP")


def show_interfaces(current_device):
    global json_output, xls_col_if_hostname, xls_col_if_interface, xls_col_if_link_status, \
        xls_col_if_protocol_status, xls_col_if_mac_address, xls_col_if_ip_address, xls_col_if_desc, \
        xls_col_if_mtu, xls_col_if_duplex, xls_col_if_speed, xls_col_if_bw, xls_col_if_delay, \
        xls_col_if_encapsulation, xls_col_if_last_in, xls_col_if_last_out, xls_col_if_queue, \
        xls_col_if_in_rate, xls_col_if_out_rate, xls_col_if_in_pkts, xls_col_if_out_pkts, xls_col_if_in_err, \
        xls_col_if_out_err, xls_col_if_access_vlan, xls_col_if_trunk_allowed, xls_col_if_trunk_forwarding, \
        xls_col_if_l2_l3, xls_col_if_trunk_access, xls_col_if_short_if, xls_col_if_trunk_native

    sheet = wb_obj['Interfaces']
    max_row = sheet.max_row + 1
    command = "show interfaces"
    command2 = "show interface status"

    if DEBUG is True:
        print("Starting gathering JSON data for '" + command + "' on " + current_device + ".")
        print("Starting gathering JSON data for '" + command2 + "' on " + current_device + ".")

    output = conn.send_command(command, use_textfsm=True)
    string_output = json.dumps(output, indent=2)

    output2 = conn.send_command(command2, use_textfsm=True)
    string_output = json.dumps(output2, indent=2)

    if DEBUG is True:
        print(string_output)
        print("///// ENDING gathering JSON data for '" + command + "' on " + current_device + "./////")
        print(string_output)
        print("///// ENDING gathering JSON data for '" + command2 + "' on " + current_device + "./////")

    file_data = wrap_command(command, string_output)
    json_output = json_output + file_data

    file_data2 = wrap_command(command2, string_output)
    json_output2 = json_output + file_data

    switchport_data_found = False
    if isinstance(output2, list):
        switchport_data_found = True

    # Write Interface data to spreadsheet 'Interfaces' tab
    if is_json(output):
        for i in output:
            short_if_name = get_short_if_name(i['interface'])
            rw_cell(max_row, xls_col_if_hostname, True, current_hostname, "Interfaces")
            rw_cell(max_row, xls_col_if_short_if, True, short_if_name, "Interfaces")
            rw_cell(max_row, xls_col_if_interface, True, i['interface'], "Interfaces")
            rw_cell(max_row, xls_col_if_link_status, True, i['link_status'], "Interfaces")
            rw_cell(max_row, xls_col_if_protocol_status, True, i['protocol_status'], "Interfaces")
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
            rw_cell(max_row, xls_col_if_last_in, True, i['last_input'], "Interfaces")
            rw_cell(max_row, xls_col_if_last_out, True, i['last_output'], "Interfaces")
            rw_cell(max_row, xls_col_if_queue, True, i['queue_strategy'], "Interfaces")
            rw_cell(max_row, xls_col_if_in_rate, True, i['input_rate'], "Interfaces")
            rw_cell(max_row, xls_col_if_out_rate, True, i['output_rate'], "Interfaces")
            rw_cell(max_row, xls_col_if_in_pkts, True, i['input_packets'], "Interfaces")
            rw_cell(max_row, xls_col_if_out_pkts, True, i['output_packets'], "Interfaces")
            rw_cell(max_row, xls_col_if_in_err, True, i['input_errors'], "Interfaces")
            rw_cell(max_row, xls_col_if_out_err, True, i['output_errors'], "Interfaces")
            if switchport_data_found is True:
                for x in output2:
                    if short_if_name == x['port']:
                        if x['vlan'].isnumeric():
                            rw_cell(max_row, xls_col_if_access_vlan, True, x['vlan'], "Interfaces")
                            rw_cell(max_row, xls_col_if_trunk_access, True, "Access", "Interfaces")
                            rw_cell(max_row, xls_col_if_l2_l3, True, "Layer 2", "Interfaces")
                        elif x['vlan'] == "trunk":
                            trunk_info = get_trunk_info(short_if_name)

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


def get_trunk_info(if_name):
    trunk_all_info = conn.send_command("show int trunk").split('\n')
    native_vlan, vlans_allowed, vlans_forwarding = "", "", ""
    x = 0
    for line in trunk_all_info:
        if x == 0:
            if left(line, len(if_name)) == if_name:
                number = re.compile(r"(\d.*)$")
                native_vlan = number.search(line).group(1)
                x = x + 1
        elif x == 1:
            if left(line, len(if_name)) == if_name:
                number = re.compile(r"(\d.*)$")
                vlans_allowed = number.search(line).group(1)
                if line.find("Vlans in spanning tree forwarding state and not pruned") != -1:
                    x = x + 1
        elif x == 2:
            if left(line, len(if_name)) == if_name:
                number = re.compile(r"(\d.*)$")
                vlans_forwarding = number.search(line).group(1)
                x = x + 1

        trunk_info = [native_vlan, vlans_allowed, vlans_forwarding]
        return trunk_info


def wrap_command(command, command_data):

    command_output = "------------------------------------------------------------" + "\n" + \
            "*******       " + command + "        *******" + "\n" + \
            "------------------------------------------------------------" + "\n" + \
            "------------------------------------------------------------" + "\n" + \
            "------------------------------------------------------------" + "\n" + \
             command_data + "\n" + \
            "------------------------------------------------------------" + "\n" + \
            "------------------------------------------------------------" + "\n"

    if DEBUG is True:
        print("Wrapping command text for text file for command '" + command + "'.")

    return command_output


def get_short_if_name(interface):

    number = re.compile(r"(\d.*)$")
    #number = re.compile(r"(\d+|\d+\.\d|\d+/\d+)$")
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


def rw_cell(var_row, var_column, write=False, value="", sheet="Main"):
    global wb_obj
    sheet = wb_obj[sheet]

    if write is False:
        value = sheet.cell(row=var_row, column=var_column).value
        return value
    elif write is True:
        sheet.cell(row=var_row, column=var_column).value = value


def is_json(myjson):
    try:
        json.loads(json.dumps(myjson))
    except ValueError as e:
        return False
    return True


main()
