import optparse
from unipath import Path
import sys
from getpass import getpass
from threading import Thread
import json
from openpyxl.workbook import Workbook
from datetime import datetime


def cli_args():
    """
    Reads the CLI options provided and returns them using the OptionParser
    Will return the Values as a dictionary
    """
    parser = optparse.OptionParser()
    # Following Values are defaulted here
    parser.add_option('-i', '--input_file',
                      dest="input_file",
                      default="GetInventory - Default.xlsx",
                      action="store",
                      help="Input file name of excel sheet"
                      )
    parser.add_option('-v', '--verbose',
                      dest="verbose",
                      default=False,
                      action="store_true",
                      help="Enable Verbose Output"
                      )
    parser.add_option('-r', '--output_raw_cli',
                      dest="output_raw_cli",
                      default=False,
                      action="store_true",
                      help="Capture the raw CLI output"
                      )
    parser.add_option('-t', '--max_threads',
                      dest='max_threads',
                      action='store',
                      default=25,
                      help='Maximum number of concurrent thread',
                      type='int'
                      )
    parser.add_option('-j', '--data_to_json',
                      dest='data_to_json',
                      action='store_true',
                      default=False,
                      help='Maximum number of concurrent threads',
                      )

    # Following values have default values in the GetInventoryProject class
    parser.add_option('-o', '--output_file',
                      dest="output_file",
                      action="store",
                      help="Output file name of excel sheet"
                      )
    parser.add_option('-d', '--output_directory',
                      dest="output_dir",
                      action="store",
                      help="Output Directory of excel sheet"
                      )
    parser.add_option('-u', '--username',
                      dest="global_username",
                      action="store",
                      help="Global Username"
                      )
    parser.add_option('-p', '--password',
                      dest="global_password",
                      action="store_true",
                      help="Global Password"
                      )
    parser.add_option('-s', '--secret',
                      dest="global_secret",
                      action="store_true",
                      help="Global Secret"
                      )

    options, remainder = parser.parse_args()

    # Print Error message and help if too many arguments
    if remainder:
        print('\nERROR:\n\tInvalid inputs detected:', remainder, '\n')
        parser.print_help()
        sys.exit()

    # Obtain Password if options are selected
    if options.global_password is True:
        options.global_password = getpass('\nGlobal Password: ')
    if options.global_secret is True:
        options.global_secret = getpass('\nGlobal Secret: ')

    # Utilizing the vars() method we can return the options as a dictionary
    return vars(options)


def print_net_dev_msg(row, net_dev, msg):
    line1 = str(row)
    if len(line1) < 8:
        line1 += (8 - len(line1)) * " "
    line1 = " " + line1
    line2 = str(net_dev.host)
    if len(line2) < 15:
        line2 += (15 - len(line2)) * " "
    print(line1, "|", line2, "|", msg)


def start_multithreading(network_devices, thread_function, max_threads=10, verbose=True, *args):
    """
    Multi Threading function for a list of devices.
    :param network_devices: List of Network Device
    :param thread_function: Function to Thread
    :param max_threads:
    :return:
    """
    thread_list = list()
    for n, net_dev in enumerate(network_devices):
        thread_list.append(Thread(target=thread_function, args=((net_dev,)+args)))
        if len(thread_list) == max_threads or n == len(network_devices) - 1:
            if verbose:
                print('Starting {} Threads out of {}'.format(len(thread_list), len(network_devices)))
            for thread in thread_list:
                thread.start()
            for thread in thread_list:
                thread.join()
            thread_list = []


def print_json(json_data, default=False):
    if default:
        print(json.dumps(json_data, indent=4, default=serialize))
    else:
        print(json.dumps(json_data, indent=4))


def serialize(obj):
    if isinstance(obj, Workbook):
        return obj.__repr__()
    return obj.__dict__


def format_datetime(raw_string='%Y-%m-%d %Hh:%Mm:%Ss'):
    """
    Format a String with the datetime special character replacements
    :param raw_string: String to format the current time to
    :return: Formatted string with current time.
    """
    return datetime.now().strftime(raw_string)


def add_time_stamp_to_file(file_path):
    """
    Adds a timestamp as an append to a file
    """
    f = open(file_path, "a+")
    f.write("\n!"+"#"*80+"!\n")
    f.write('!'+" "*30+format_datetime("%Y-%m-%d %H:%M:%S")+" "*30)
    f.write("\n!"+"#"*80+"!\n")
    f.close()


def center_message(msg, max_len, filler=" "):
    spacer = int((max_len - len(msg))/2)-1
    msg = spacer*filler + " " + msg + " " + spacer*filler
    if max_len-len(msg):
        msg += filler
    return msg
