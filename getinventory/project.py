import os
import sys
import netcapt
from . import helper_functions as hf
from . import xls
from . import networkdevice
from . import threadable_functions as thread_f
from unipath import Path

class GetInventoryProject:
    # Start Row of Main column to begin reading Network Devices
    row_start = 8
    default_device_settings = {
        'conn_timeout': 30,
        'keepalive': 10,
        'global_delay_factor': 5,
        'auto_connect': False
    }

    def __init__(self):
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
            'gather_version': False,
            'gather_arp': False,
            'gather_mac': False,
            'gather_interface': False,
            'gather_cdp': False,
            'gather_lldp': False,
            'gather_route': False,
            'gather_bgp': False,
            'gather_inventory': False,
            'gather_commands': False,
            'gather_ip_mroute': False
        }

        # List of netcapt Network Devices, will only be in list of Status is set to 'Yes'
        self.network_devices = list()

        # Need to update path to a Path object
        self.output_path = Path(self.output_path)

        self.load_the_project()
        # for key, val in vars(self).items():
        #     print(key, ':', val)

    def load_the_project(self):
        """
        Run all the Methods to get the project loaded
        """
        self.build_out_directories()
        self.update_gather_functions()
        self.update_network_devices()

    def build_out_directories(self):
        print(self.output_path)
        self.output_path.mkdir(parents=True)
        subdirectories = ['raw_cli_logs', 'json_output', 'gather_commands']
        for subdir in subdirectories:
            self.output_path.child(subdir).mkdir(parents=True)


    def start(self):
        """
        Contains all the functionality for the project.
        """
        hf.start_multithreading(self.network_devices, thread_f.run_gathers, self.max_threads)

    def end(self):
        self.save_everything()
        self.open_xls_file_in_app()


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
        for i, gather_f in enumerate(self.gather_f_bools):
            value = xls.rw_cell(self.work_book['Settings'], 5+i, 2)
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

    def save_work_book(self):
        xls.save_xls_retry_if_open(self.work_book, self.output_file, self.output_path)

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
        # self.write_device_info_to_wb()
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
        if sys.platform == 'win32':
            os.system('start ' + str(file_location))
        # Open Command for Mac
        elif sys.platform == 'darwin':
            os.system('open "' + str(file_location) + '"')