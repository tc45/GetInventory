import os
import sys
import netcapt
import json
from . import helper_functions as hf
from . import xls
from . import networkdevice
from . import threadable_functions as thread_f
from unipath import Path
from . xlsx_output_key_map import OUTPUT_KEY_MAP, DEVICE_INFO_MAP, INTERACE_COUNT_MAP
from datetime import datetime


class GetInventoryProject:
    # Start Row of Main column to begin reading Network Devices
    row_start = 8
    default_device_settings = {
        'conn_timeout': 30,
        'keepalive': 10,
        'global_delay_factor': 5,
        'auto_connect': False
    }
    output_key_map = OUTPUT_KEY_MAP
    device_info_map = DEVICE_INFO_MAP
    interface_count_map = INTERACE_COUNT_MAP
    # list of gather functions to not add to the work book
    gathers_to_txt_file = ['gather_config', 'gather_commands']

    dflt_col_width = 25

    def __init__(self):
        self.start_time = datetime.now()
        self.end_time = None

        cli_args = hf.cli_args()

        # only values that are defaulted by cli_args function
        self.input_file = Path(cli_args['input_file'])
        cli_args.pop('input_file')
        self.verbose = cli_args['verbose']
        cli_args.pop('verbose')
        self.output_raw_cli = cli_args['output_raw_cli']
        cli_args.pop('output_raw_cli')
        self.max_threads = cli_args['max_threads']
        cli_args.pop('max_threads')
        self.data_to_json = cli_args['data_to_json']
        cli_args.pop('data_to_json')

        # Open up the Work Book
        self.work_book = xls.open_xls(self.input_file)

        # Default Values Need to be updated here,
        # Following values are not replaceable via CLI

        # Following values will be replaced first by Excel Information, then by CLI
        self.output_file = 'GetInventory - Output.xlsx'
        self.output_path = str()
        self.global_username = str()
        self.global_password = str()
        self.global_secret = str()

        # Update the Arguments from CLI and from XLSX File.
        self.update_cli_xlsx_attributes(cli_args)

        # This should be in the same order as the settings page
        self.gather_f_bools = {
            'gather_arp': False,
            'gather_mac': False,
            'gather_interfaces': False,
            'gather_cdp': False,
            'gather_lldp': False,
            'gather_route': False,
            'gather_bgp': False,
            'gather_inventory': False,
            'gather_commands': False,
            'gather_ip_mroute': False,
            'gather_ap': False,
            'gather_config': False,
        }

        # List of netcapt Network Devices, will only be in list of Status is set to 'Yes'
        self.network_devices = list()

        # Need to update path to a Path object
        self.output_path = Path(self.output_path)

        # Get the List of Commands
        self.other_commands = xls.cell_iter_to_list(self.work_book["Commands"]["A"], True)

        self.load_the_project()

    def load_the_project(self):
        """
        Run all the Methods to get the project loaded
        """
        self.print_start()
        self.build_out_directories()
        self.update_gather_functions()
        self.update_network_devices()

    def print_start(self):
        print(60 * "_")
        print(hf.center_message('Start Time:' + str(self.start_time), 60))
        print(hf.center_message('Input File:' + str(self.input_file), 60))
        print(60 * "_")

    def print_end(self):
        self.end_time = datetime.now()
        print(60 * "_")
        print(hf.center_message('Elapsed Time: %s' % (self.end_time - self.start_time), 60))
        print(hf.center_message('End Time: %s' % self.end_time, 60))
        print(60 * "_")

    def build_out_directories(self):
        self.output_path.mkdir(parents=True)
        # subdirectories = ['raw_cli_logs', 'json', 'gather_commands']
        if self.gather_f_bools['gather_commands']:
            self.output_path.child('gather_commands').mkdir(parents=True)
        if self.data_to_json:
            self.output_path.child('json').mkdir(parents=True)
        if self.output_raw_cli:
            self.output_path.child('raw_cli_logs').mkdir(parents=True)

    def start(self):
        """
        Contains all the functionality for the project.
        """
        hf.start_multithreading(self.network_devices, thread_f.run_gathers, self.max_threads)

    def end(self):
        self.save_everything()
        self.open_xls_file_in_app()
        self.print_end()

    def wr_net_dev_to_wb(self):
        for net_dev in self.network_devices:
            self._net_dev_to_main(net_dev)
            self._intf_count_to_main(net_dev)
            self._net_dev_gathers_to_wb(net_dev)

    def _net_dev_gathers_to_wb(self, net_dev: networkdevice.NetworkDevice):
        for gather_name, data in net_dev.gather_data.items():
            sheet_name = ' '.join(gather_name.split('_')[1:]).upper()
            # Only Continue for items have an index built out.
            if gather_name in self.output_key_map:
                # Only build if we want it in the Work Book
                self._build_ws_obj(sheet_name, self.output_key_map[gather_name])
                # Add all the Data
                try:
                    self._gather_to_ws(sheet_name, gather_name, data, net_dev.hostname)
                except Exception as e:
                    net_dev.add_exception_error(e)

    def _gather_to_ws(self, sheet_name, gather_name, data, hostname):
        """
        Cycle through all the Gather Data for one gather and output it ot the Works Sheet.
        """
        sheet_obj = self.work_book[sheet_name]
        next_row = xls.next_available_row(sheet_obj)
        # If string either data not available of Feature is not enabled
        if isinstance(data, str) or data is None:
            xls.rw_cell(sheet_obj, next_row, 1, hostname)
            if data is not None and 'is not enabled' in data:
                xls.rw_cell(sheet_obj, next_row, 2,  data.replace('\n', ''))
            else:
                xls.rw_cell(sheet_obj, next_row, 2, 'No Data')
            return

        # Need to build simple sheet_map
        sheet_map = self.__build_sheet_map(gather_name, data[0].keys())

        # Cycle through the data
        for i, data_entry in enumerate(data):
            xls.rw_cell(sheet_obj, next_row + i, 1, hostname)
            for key in sheet_map:
                # Input the data only if in sheet_map
                if key in data_entry.keys():
                    wr_val = data_entry[key]
                    # Clean up any Values that are list to comma seperated
                    if isinstance(wr_val, list):
                        wr_val = ', '.join(wr_val)
                    xls.rw_cell(sheet_obj, next_row+i, sheet_map[key], wr_val)

    def output_data_to_json(self):
        output_folder = self.output_path.child('json')
        for net_dev in self.network_devices:
            output_file = output_folder.child(net_dev.hostname+'_'+net_dev.host+'.json')
            net_dev.verbose_msg("Saving Data to JSON File")
            json_data = net_dev.gather_data.copy()
            # Popping off the Commands that are going to their own file data as it is not parsed and can be huge
            for pop_key in self.gathers_to_txt_file:
                if pop_key in json_data:
                    json_data.pop(pop_key)
            json_data = json.dumps(json_data, indent=4)
            output_file.write_file(json_data)

    def __build_sheet_map(self, gather_name, data_keys):
        """
        Need to simpolify complex OUTPUT_KEY_MAP to a more {key:column} simplicity,
        this will extract all values that do match.
        """
        sheet_map = dict()
        for mapper in self.output_key_map[gather_name]:
            for key in mapper['keys']:
                if key in data_keys:
                    sheet_map[key] = mapper['column']
        return sheet_map

    def _build_ws_obj(self, sheet_name, list_of_dict_w_col_name, first_col_names=None):
        if first_col_names is None:
            first_col_names = ['hostname']

        # Build out Work Sheet if it does not exist, will build out with all the Column Names
        # the Default value for first Column is 'Hostname'
        if sheet_name not in self.work_book.sheetnames:
            self.work_book.create_sheet(sheet_name)
            sheet_obj = self.work_book[sheet_name]
            for i, col_name in enumerate(first_col_names, 1):
                xls.rw_cell(sheet_obj, 1, i, col_name)
                xls.set_width(sheet_obj, self.dflt_col_width, i)
            # Build out the features
            for mapper in list_of_dict_w_col_name:
                # Add the Column Name
                xls.rw_cell(sheet_obj, 1, mapper['column'], mapper['column_name'])

                # Set Width, if value available, otherwise set to default
                if 'width' in mapper.keys():
                    xls.set_width(sheet_obj, mapper['width'], mapper['column'])
                else:
                    xls.set_width(sheet_obj, self.dflt_col_width, mapper['column'])

                # Enable the filter option
                sheet_obj.auto_filter.ref = sheet_obj.dimensions

    def _intf_count_to_main(self, net_dev: networkdevice.NetworkDevice):
        ws_obj = self.work_book['Main']
        if 'interface_count' in net_dev.device_info.keys():
            for intf_name, vals in net_dev.device_info['interface_count'].items():
                if intf_name in self.interface_count_map.keys():
                    for key, col in self.interface_count_map[intf_name].items():
                        xls.rw_cell(ws_obj, net_dev.main_row, col, vals[key])

    def _net_dev_to_main(self, net_dev: networkdevice.NetworkDevice):
        ws_obj = self.work_book['Main']
        # Write the Status of the Device
        xls.rw_cell(ws_obj, net_dev.main_row, 2, net_dev.status)
        for mapper in self.device_info_map:
            for key in mapper['keys']:
                if key in net_dev.device_info.keys():
                    wr_val = net_dev.device_info[key]
                    if isinstance(wr_val, list):
                        wr_val = ', '.join(wr_val)
                    xls.rw_cell(ws_obj, net_dev.main_row, mapper['column'], wr_val)
                    break

    def save_work_book(self):
        xls.save_xls_retry_if_open(self.work_book, self.output_file, self.output_path)

    def update_network_devices(self):
        devices_params = self.read_net_devs_from_xls()
        self.build_network_device(devices_params)

    def build_network_device(self, devices_params):
        for one_device_param in devices_params:
            # Build a NetworkDevice Class from the one generated here
            # That class will hold the NetCapt object in netcapt_handle
            net_dev = networkdevice.NetworkDevice(**one_device_param)
            self.network_devices.append(net_dev)

    def read_net_devs_from_xls(self):
        return_list_device_params = list()
        sheet_obj = self.work_book['Main']
        # Cycle from row start to the end of the available rows
        for row in range(self.row_start, sheet_obj.max_row + 1):
            device_params = dict()
            host = xls.rw_cell(sheet_obj, row, 1)
            # Clean up any spaces
            if host is not None:
                host = host.replace(' ', '')
            dev_status = xls.rw_cell(sheet_obj, row, 2)
            if dev_status:
                dev_status = dev_status.lower()
            # Only add to list if host is not empty and status is yes
            if dev_status == 'yes' and host:
                device_params = {
                    'host': host,
                    'main_row': row,
                    'status': dev_status,
                    'device_type': xls.rw_cell(sheet_obj, row, 3),
                    'verbose': self.verbose,
                    'username': self.global_username,
                    'password': self.global_password,
                    'secret': self.global_secret,
                    'gather_f_bools': self.gather_f_bools,
                    'output_path': self.output_path,
                    'output_raw_cli': self.output_raw_cli,
                    'other_commands': self.other_commands,
                }
                # Add in default parameters
                device_params.update(self.default_device_settings)

                # Optional protocols on a per device bassis
                protocol = xls.rw_cell(sheet_obj, row, 4)
                port = xls.rw_cell(sheet_obj, row, 5)
                username = xls.rw_cell(sheet_obj, row, 6)
                password = xls.rw_cell(sheet_obj, row, 7)

                # Overwrite or default to SSH
                if device_params['device_type'] != 'autodetect':
                    if protocol:
                        device_params['device_type'] = device_params['device_type'] + '_' + protocol
                    else:
                        device_params['device_type'] = device_params['device_type'] + '_ssh'

                # Overwrite the following parameters
                if port:
                    device_params['port'] = port
                if username:
                    device_params['username'] = username
                if password:
                    device_params['password'] = password
                    device_params['secret'] = password
                return_list_device_params.append(device_params)
        return return_list_device_params

    def update_gather_functions(self):
        """
        Cycle through all the gather functions and determine if we need to run the gather function.
        """
        sheet_row_start = 5
        for i, gather_f in enumerate(self.gather_f_bools):
            value = xls.rw_cell(self.work_book['Settings'], sheet_row_start+i, 2)
            if value.lower() == 'yes':
                self.gather_f_bools[gather_f] = True

    def verbose_msg(self, msg):
        """
            print a message if verbose option is true.
        """
        if self.verbose:
            print(msg)

    def _read_one_device(self, sheet_obj, row):
        host = xls.rw_cell(sheet_obj, row, 1)
        if host is not None:
            host = host.replace(' ', '')
        dev_status = xls.rw_cell(sheet_obj, row, 2)
        if host and dev_status.lower() == 'yes':
            device_settings = {
                'main_row': row,
                'host': host,
                'device_type': xls.rw_cell(sheet_obj, row, 3),
                'protocol': xls.rw_cell(sheet_obj, row, 4),
                'port': xls.rw_cell(sheet_obj, row, 5),
                'username': xls.rw_cell(sheet_obj, row, 5),
                'password': xls.rw_cell(sheet_obj, row, 5),
                'verbose': self.verbose,
            }
            device_settings += self.default_device_settings

    def __update_cli_arg(self, cli_args):
        # if None, then option is defaulted and no value entered
        # Only options with default value are Verbose and Input File
        # Loops through Keys and updates
        for key in cli_args:
            if cli_args[key] is not None:
                setattr(self, key, cli_args[key])

    def update_cli_xlsx_attributes(self, cli_args):
        """
        Update the Variables based on the Excel sheet and finally on CLI Arguments
        :return:
        """
        self.__update_all_xlsx_attributes()
        # Goes last, as the CLI_ARGS take priority, so it can overwrite the EXCEL Args
        self.__update_cli_arg(cli_args)

        # Will use global_password is no global_secret was provided
        if not self.global_secret:
            self.global_secret = self.global_password

    def __read_xls_var(self, variable, row, column):
        # update from Cell value
        cell_val = xls.rw_cell(self.work_book["Main"], row, column)
        # if no cell value then it will not change default value
        if cell_val is not None:
            setattr(self, variable, cell_val)

    def __update_all_xlsx_attributes(self):
        attribute_list = ['global_username', 'global_password', 'global_secret', 'output_path', 'output_file']
        # WIll only work if the attributes are in order without a Value in between.
        for i, key in enumerate(attribute_list, 1):
            self.__read_xls_var(key, i, 2)

    def write_device_errors_to_wb(self):
        """
        Write all the Device Errors and Comments to the workbook.
        :return:
        """
        for net_dev in self.network_devices:
            xls.list_of_list_to_ws(self.work_book['Comments'], net_dev.comment_messages)

    def clear_credentials(self):
        # Clear the Global Credentials
        for i in range(1,3):
            xls.clear_cell(self.work_book['Main'], i, 2)
        # Clear Per Device Username
        for cell in self.work_book["Main"]["F"][self.row_start-1:]:
            cell.value = None
        # Clear Per Device Password
        for cell in self.work_book["Main"]["G"][self.row_start-1:]:
            cell.value = None

    def save_everything(self):
        """
        Rubn through all of the save functions
        :return:
        """
        if self.data_to_json:
            self.output_data_to_json()
        self.save_gathers_to_txt()
        self.wr_net_dev_to_wb()
        self.write_device_errors_to_wb()
        self.clear_credentials()
        self.save_work_book()

    def open_xls_file_in_app(self):
        """
        Open the WorkBook with Native Application
        """
        file_location = xls.add_xls_tag(self.output_file)
        file_location = self.output_path.child(file_location)
        # Open command for Windows
        self.verbose_msg('Opening the file: %s' % file_location)
        if sys.platform == 'win32':
            os.system('start ' + str(file_location))
        # Open Command for Mac
        elif sys.platform == 'darwin':
            os.system('open "' + str(file_location) + '"')

    def save_gathers_to_txt(self):
        for net_dev in self.network_devices:
            self._dev_data_to_txt(net_dev)

    def _dev_data_to_txt(self, net_dev):
        for gather_fun_name in self.gathers_to_txt_file:
            if gather_fun_name in net_dev.gather_data.keys():
                data = net_dev.gather_data[gather_fun_name]
                f_path = self.output_path.child(gather_fun_name.replace('gather_', ''))
                f_path.mkdir(parents=True)
                f_path = f_path.child(net_dev.host+'_'+gather_fun_name+'.txt')
                if isinstance(data, str):
                    f_path.write_file(data)
                elif isinstance(data, dict):
                    msg = str()
                    for cmd, output in data.items():
                        msg = '\n'+hf.center_message(cmd, 60, '-') + '\n' + output
                    f_path.write_file(msg)
