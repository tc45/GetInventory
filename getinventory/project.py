import netcapt
from . import helper_functions as hf
from . import xls
from unipath import Path

class GetInventoryProject:
    # Start Row of Main column to begin reading Network Devices
    row_start = 8
    default_device_settings = {
        'conn_timeout': 30,
        'keepalive': 10,
        'global_delay_factor': 5,
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
        self.output_directory = ''
        self.global_username = ''
        self.global_password = ''
        self.global_secret = ''

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


        x = vars(self).copy()
        # x.pop('work_book')
        print(x)
        hf.print_json(x, True)


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

    def read_net_devs_from_xls(self):
        sheet_obj = self.work_book['Main']

        for row in range(self.row_start, sheet_obj.max_row + 1):
            pass

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
        self.__update_all_xlsx_attributes
        # Goes last, as the CLI_ARGS take priority, so it can overwrite the EXCEL Args
        self.__update_cli_arg(cli_args)

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
        for i, key in enumerate(attribute_list):
            self.__read_xls_var(key, i, 2)

