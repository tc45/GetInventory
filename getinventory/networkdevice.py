import sys
import traceback
from typing import Union
from netcapt.cisco import CiscoNxosDevice, CiscoXrDevice, CiscoWlcDevice, CiscoIosDevice
from unipath import Path
import netcapt
from . import helper_functions as hf
from netmiko.ssh_exception import AuthenticationException
from netcapt.netcapt_exceptions import GatherAttributeError
from datetime import datetime

from threading import Semaphore
SCREEN_LOCK = Semaphore(value=1)


class InitiatingConnectionException(Exception):
    pass


class NetworkDevice:
    netcapt_handle: Union[None, CiscoIosDevice, CiscoXrDevice, CiscoNxosDevice, CiscoWlcDevice]

    def __init__(
            self, gather_f_bools, status,
            main_row, output_path,
            output_raw_cli, verbose, other_commands, **kwargs
    ):
        self.status = status
        self.gather_f_bools = gather_f_bools
        self.main_row = main_row
        self.output_raw_cli = output_raw_cli
        self.output_path = Path(output_path)
        self.start_time = None
        self.end_time = None
        self.other_commands = other_commands

        self.verbose = verbose
        self.host = kwargs['host']
        # The autodetect option will be handled by the GetNetworkDevice function
        self.netcapt_handle = netcapt.GetNetworkDevice(**kwargs)

        self.gather_data = dict()

        self.collection_time = str()
        self.hostname = str()

        # Adding all Device Information here instead of instantiating a variable
        self.device_info = dict()

        # Stores a list of all the error/Comments
        self.comment_messages = list()

    def add_exception_error(self, e, e_type='Error', set_status='Error | See Comment(s)'):
        """
        Adds a detected Exception Error to the comment messages as an Error, will handle the data parsing to
        ensure consistency.

        If need to add in a message please see add_cmnt_msg method
        :param e_type: Error Type
        :param e: Exception
        :return:
        """
        if set_status:
            self.status = set_status
        exc_tb = sys.exc_info()[2]
        exc_type = sys.exc_info()[0]
        exc_line = exc_tb.tb_lineno
        f_name = traceback.extract_tb(exc_tb, 1)[0][2]
        traceback_str = traceback.format_exc()
        t_err_msg = "{} | Exception Type: {} | At Function: {} | Line No: {} | Error Message: {} | FULL MESSAGE:\n{}"
        t_err_msg = t_err_msg.format(self.host, exc_type, f_name, exc_line, e, traceback_str)
        self.add_cmnt_msg(e, e_type, exc_type, t_err_msg)
        self.verbose_msg(
            'Exception ERROR: Exception Type: {} | Error Message: {}'
            '(See Excel Output for full message)'.format(exc_type, f_name, e)
        )

    def add_cmnt_msg(self, msg, cmnt_type, exception_type=None, debug_msg=str()):
        """
        Add a comment to the list, to keep track of any Exceptions or comments with the device.
        :param debug_msg: Message for Debugging purposes, such as the traceback string
        :param msg: str() Message of the comment
        :param cmnt_type: str() comment type, i.g. Error, Comment, Issue, Concern, etc.
        :param exception_type: optional Exception type, to have a column identifying the Exception type.
        :return: None
        """
        comment = str(len(self.comment_messages) + 1) + " | "
        comment += str(msg)
        t_time = hf.format_datetime("%Y-%m-%d_%Hh%Mm%Ss")
        self.comment_messages.append(
            [self.host, self.hostname, self.main_row, cmnt_type, t_time, msg, exception_type, debug_msg]
        )

    def go_gather(self):
        """
        Loops through all the Gather_Bool items and gets attribute from netcapt_handle.
        Will only run if True.
        """
        for gather_fun_name, gather_true_false in self.gather_f_bools.items():
            try:
                if gather_true_false:
                    self.verbose_msg('Running: ' + gather_fun_name)
                    if gather_fun_name == 'gather_commands':
                        data = self.gather_commands(self.other_commands)
                        self.gather_data[gather_fun_name] = data
                    else:
                        self.gather_data[gather_fun_name] = getattr(self.netcapt_handle, gather_fun_name)()
            except GatherAttributeError as e:
                self.add_exception_error(e, 'Unsupported Gather Method', "")
            except Exception as e:
                self.add_exception_error(e, 'Error: '+gather_fun_name)

    def gather_commands(self, commands):
        cmd_output = dict()
        for cmd in commands:
            try:
                output = self.netcapt_handle.gather_commands([cmd])
                cmd_output.update(output)
            except Exception as e:
                self.add_exception_error(e)
        return cmd_output


    def print_msg(self, msg):
        """
        Print Message with main_row value
        """
        line1 = str(self.main_row)
        if len(line1) < 6:
            line1 += (6 - len(line1)) * " "
        line1 = " " + line1
        line2 = str(self.host)
        if len(line2) < 15:
            line2 += (15 - len(line2)) * " "
        SCREEN_LOCK.acquire()
        print(line1, "|", line2, "|", msg)
        SCREEN_LOCK.release()

    def verbose_msg(self, msg):
        """
        Method to handle Verbose Message option.
        :param msg: msg to print if Global Verbose is true
        :return: None
        """
        if self.verbose:
            self.print_msg(msg)

    def update_dev_info(self):
        self.hostname = self.netcapt_handle.hostname
        self.device_info['hostname'] = self.hostname
        self.device_info['device_type'] = self.netcapt_handle.classname

        # Load 'show version info to device_info
        self.verbose_msg('Running: gather_version')
        data = self.netcapt_handle.gather_version()
        if isinstance(data, list):
            self.device_info.update(data[0])

        # Add the Interface Counts
        self.verbose_msg('Getting Interface Counts')
        data = self.netcapt_handle.count_intf()
        if isinstance(data, dict):
            self.device_info['interface_count'] = data

        # Get SFP Count
        self.verbose_msg('Getting SFP Count')
        data = self.netcapt_handle.get_sfp()
        self.device_info['sfp_count'] = len(data)

    def end_connection(self):
        """ Close Connection"""
        self.netcapt_handle.end_connection()
        self.verbose_msg('Ending CLI connection')

    def start_raw_cli_log(self):
        if self.output_raw_cli:
            raw_cli_f_path = self.output_path.child('raw_cli_logs').child(self.host+'_raw_cli.log')
            hf.add_time_stamp_to_file(raw_cli_f_path)
            self.netcapt_handle.start_raw_cli_log(raw_cli_f_path)

    def start_connection(self, max_attempts=3):
        """
        Start Connection and handle attempts, This method will allow connection reattempts
        :param max_attempts: int default is 3
        """
        self.verbose_msg('Starting Connection')
        attempt = 0
        for attempt in range(max_attempts):
            try:
                self.netcapt_handle.start_connection()
                self.start_raw_cli_log()
                return
            except AuthenticationException as e:
                self.add_exception_error(e, traceback.format_exc())
            except Exception as e:
                self.add_exception_error(e, traceback.format_exc())
            self.verbose_msg('LOGIN FAILURE: Attempt to Establish Connection Failed, Attempt: ' + str(attempt))
        raise InitiatingConnectionException(
            'LOGIN FAILURE: Attempt to Establish Connection Failed, Attempt: ' + str(attempt)
        )

    def update_time(self, start_end):
        if start_end == 'start':
            self.start_time = datetime.now()
            self.verbose_msg('Start Time: {}'.format(self.start_time))
            self.device_info['start_time'] = str(self.start_time)
        elif start_end == 'end':
            self.end_time = datetime.now()
            self.verbose_msg('End Time: {}'.format(self.end_time))
            self.device_info['end_time'] = str(self.end_time)
            if self.start_time is not None and self.end_time is not None:
                self.device_info['elapsed_time'] = str(self.end_time - self.start_time)
                self.verbose_msg('Elapsed Time: {}'.format(self.device_info['elapsed_time']))
