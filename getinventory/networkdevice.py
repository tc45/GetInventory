import sys
import traceback
from unipath import Path
import netcapt
from netmiko.ssh_exception import AuthenticationException
from . import helper_functions as hf


class InitiatingConnectionException(Exception):
    pass


class NetworkDevice:
    def __init__(self, gather_bool, status, main_row, **kwargs):
        self.status = status
        self.gather_bool = gather_bool
        self.main_row = main_row
        self.elapsed_time = int()

        self.verbose = kwargs['verbose']
        self.host = kwargs['host']
        self.netcapt_handle = netcapt.GetNetworkDevice(**kwargs)

        self.gather_data = dict()

        # Below Needs to go into device info
        # TODO: Need to work on Migrating this to device_attributes
        self.collection_time = str()
        self.hostname = str()
        self.cpu_5_sec = str()
        self.cpu_1_min = str()
        self.cpu_5_min = str()
        self.cpu_15_min = str()
        self.version = str()

        # Adding all
        self.device_info = dict()

        # Stores a list of all the error/Comments
        self.comment_messages = list()
        self.output_path = Path()

    def add_exception_error(self, e, e_type='Error'):
        """
        Adds a detected Exception Error to the comment messages as an Error, will handle the data parsing to
        ensure consistency.

        If need to add in a message please see add_cmnt_msg method
        :param e_type: Error Type
        :param e: Exception
        :return:
        """
        exc_tb = sys.exc_info()[2]
        exc_type = sys.exc_info()[0]
        exc_line = exc_tb.tb_lineno
        f_name = traceback.extract_tb(exc_tb, 1)[0][2]
        traceback_str = traceback.format_exec()
        t_err_msg = "{} | Exception Type: {} | At Function: {} | Line No: {} | Error Message: {} | FULL MESSAGE:\n{}"
        t_err_msg = t_err_msg.format(self.host, exc_type, f_name, exc_line, e, traceback_str)
        self.add_cmnt_msg(e, e_type, exc_type, t_err_msg)
        self.verbose_msg(
            'Exception ERROR: Exception Type: {} | At Function: {} | Error Message: {}'
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
        for gather_fun_name, gather_true_false in self.gather_bool.items():
            try:
                if gather_true_false:
                    self.gather_data[gather_fun_name] = getattr(self.netcapt_handle, gather_fun_name)()

            except Exception as e:
                self.add_exception_error(e)

    def print_msg(self, msg):
        """
        Print Message with main_row value
        """
        line1 = str(self.main_row)
        if len(line1) < 8:
            line1 += (8 - len(line1)) * " "
        line1 = " " + line1
        line2 = str(self.host)
        if len(line2) < 15:
            line2 += (15 - len(line2)) * " "
        print(line1, "|", line2, "|", msg)

    def verbose_msg(self, msg):
        """
        Method to handle Verbose Message option.
        :param msg: msg to print if Global Verbose is true
        :return: None
        """
        if self.VERBOSE:
            self.print_msg(msg)

    def update_dev_info(self):
        pass

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
                return
            except AuthenticationException as e:
                self.add_exception_error(e, traceback.format_exc())
            except Exception as e:
                self.add_exception_error(e, traceback.format_exc())
            self.verbose_msg('LOGIN FAILURE: Attempt to Establish Connection Failed, Attempt: ' + str(attempt))
        raise InitiatingConnectionException(
            'LOGIN FAILURE: Attempt to Establish Connection Failed, Attempt: ' + str(attempt)
        )
