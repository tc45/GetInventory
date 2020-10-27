# GetInventory

GetInventory is a network device collector script.  The purpose of this tool is to quickly and efficiently allow
network engineers and operators to gather details from multiple network devices.  Two primary sets out output will be
produced per device.  The first is an Excel spreadsheet.  The spreadsheet serves as the input for the application, copies itself to an
output directory, and updates the output file with a standard set of information collected per device.  The information gathered is
parsed and added into tables on individual tabs within the file.  In addition to the Excel file,
each device can have a list of user input commands added into the 'Command' tab which will then be run against every device and a
unique file will be output per device.  These can be any standard command that the platform will support.  

[![published](https://static.production.devnetcloud.com/codeexchange/assets/images/devnet-published.svg)](https://developer.cisco.com/codeexchange/github/repo/tc45/GetInventory)

**Excel Spreadsheet output**
| Output | Decription | Tab | Note |
| --- | --- | --- | --- |
| Device details | Hostname, model, IOS, etc | Main |
| Interface count | count of IF by type | Main
| SFP Counts  | count of SFP by type | Main | COMING SOON! |
| Inventory Details | parsed output of 'show inventory' | Inventory |
| ARP Tables | parsed output of 'show arp' | ARP | VRF Aware |
| MAC Addresses | parsed output of 'show mac-address' | MAC |
| Routing tables | parsed output of 'show ip route' | Routes | VRF Aware |
| BGP tables | parsed output of 'show ip bgp' | BGP | **NOT** VRF Aware (COMING SOON) |
| Interface info | parsed output of 'show interface' and 'show interface status' | Interfaces | VRF Aware |
| CDP Details | parsed output of 'show cdp neighbor detail' | CDP |
| LLDP Details | parsed output of 'show lldp neighbor detail' | LLDP |

**Note**: GetInventory is NOT designed to push any configurations, only pull using show commands.  The command 'config t' will cause
unexpected behavior.

## Device Support

GetInventory uses the python library Netmiko for SSH and Telnet connectivity.  Although the platform is extensible to support multiple
vendors, only the vendors and OSes listed below are validated.  The application may run in 'command only' mode (if all parsing
features are turned off), against vendor OSes not listed here.

**Vendor OS Supported**
* Cisco IOS
* Cisco NX-OS
* Cisco IOS XR **Beta**
* ExtremeXOS **Limited functionality**

## Platform Support

GetInventory is intended to be cross-platform supportable, and has tested on Windows and MAC platforms. LINUX is in Beta mode, please report any issues when utilizing with these platforms.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine or VM for deployment.  

### Prerequisites

Netmiko runs on Python and uses input from an Excel spreadsheet.  The following details the minimum requirements to run this application.

* Microsoft Excel XLSX editor
* Python 3.6 or higher (recommend 3.7.6 as of 3/26/2020)
* Python module - Netmiko v3.0.0
* Python module - Openpyxl v3.0.3

### Installing

Once Python is installed, the following steps should be followed to get the application functioning.

1) Download the repository from **[Github](https://github.com/tc45/GetInventory)**.  Download using the ZIP file option, or
you can use GIT or SVN to pull the repository using the URL https://github.com/tc45/GetInventory.git.  

![Download GetInventory](https://github.com/tc45/images/blob/master/GetInventory_Steps_download.gif)

If downloading via ZIP:
2) Copy the ZIP file to a directory of your choice

```
copy GetInventory.zip c:\python\GetInventory
```

3) Extract the files into the directory created
4) Install python modules using pip

```
c:\python\GetInventory>pip install -r requirements.txt
```
5) Once requirements install without error, the application is ready to use.

## Using GetInventory

GetInventory relies on an input Excel spreadsheet to execute.  A default XLSX file is included in the repository and should be
used to get started.  At a minimum, a list of IP address/hostnames need to be added to get
 get started.  The following is a full list of variables that may be needed at execution time.

| Variable | Required | Input area | Description |
| ---- | --- | --- | --- |
| IP Address/Hostnames | X | Spreadsheet only | Devices that need data collected from |
| Username | X | Spreadsheet or command line | SSH/Telnet username|
| Password | X | Spreadsheet or command line | SSH/Telnet password |
| Secret | X | Spreadsheet or command line | Enable password (Defaults to regular password if not specified) |
| Output Directory | X |  Spreadsheet or command line | Directory where output spreadsheet and command files will be stored |
| Output Filename | X | Spreadsheet or command line | Specifies name of output spreadsheet |
| Parse Method | | Spreadsheet only | Specifies 'device type' to be parsed.  Defaults to cisco_ios) |
| Protocol | | Spreadsheet only | Specify SSH or Telnet (Defaults to SSH if not specified) |
| Port | | Spreadsheet only | Specify port if not standard port (Defaults 22 for SSH and 23 for Telnet |
| Device Username | | Spreadsheet only | Per device override for username |
| Device Password | | Spreadsheet only | Per device override for password |


### Open spreadsheet

Open the default spreadsheet, **GetInventory - Default.xlsx**, in the default directory.  The file should default to the Main tab,
but if not go ahead and click on the 'Main' tab.  A sample of the beginning file is shown below.

![](https://github.com/tc45/images/blob/master/GetInventory_Tab_Main_3.26.2020.jpg)



**Note:** Any columns with a black header are currently not functioning.**

### Add General Connectivity Details

The following values can be specified in the spreadsheet or added at run time to the command line.  See section below on ** Command Line
Options ** for additional details on how to add via command line.  Any values specified at command line will override the value
in the spreadsheet.

Change the following values to the requirements of your own project.

| Variable | Location | Default |
| ---- | --- | --- |
| Username | Cell B1 | local |
| Password | Cell B2 | local |
| Secret | Cell B3 | local |
| Output Directory | Cell B4 | c:\temp\GetInventory\lab_test\run1 |
| Output Name | Cell B5 | Lab Testing |

### Add Devices Details


Update the list of Hosts that need data collected from in column A starting at row 8.  Add one device IP address/hostname per line.

If the device is not a Cisco IOS device, go ahead and update Parse Method to the appropriate parser.  Refer to section above on 'Device Support'
to see what platforms are supported.  Currently only cisco_ios and cisco_nxos are supported.


**Starting at row 8:**

| Variable | Column | Required | Description |
| --- | --- | --- |  --- |
| Host | A | Yes | Hostname/IP Address of device to connect |
| Active | B | No | Ignores hosts set to 'No'.  <br/>Options: Yes/No <br/>Default: blank |
| Parse Method | C | No | Defines OS parser to be used.  <br/>Options: cisco_ios, cisco_nxos, autodetect* <br/>Default: cisco_ios(blank) |
|Protocol | D | No | Toggle connection protocol.  <br/>Options: Telnet, ssh  <br/>Default: ssh(blank)
| Port Override | E | No | Override default ports of 22(ssh) or 23 (telnet).  0-65535  |

* Autodetect - Logs into device to try and determine IOS before making selection.  Takes approx 10 sec additional per device

Once device details including at least Hostnames/IP addresses have been added to the spreadsheet, save and close the Excel file.  


### Commands tab

All commands entered into the 'Commands' tab will be executed per device and output to an individual file in the
output directory specified.  Each file will be named after the hostname of the device that was entered in the Main tab.  
The commands should be entered one per line in column A.  All other columns will be ignored.  Any command that doesn't
require input should be valid.  The default file has examples of typical commands, but this can be expanded to anything
you need.  Regular expressions, pipe | begin|include|section also work in these commands.  Note that these commands
are not platform specific.  If you put in a switch command, it will be run on a router, but just output with '^ Invalid
input detected' message.

![GetInventory - Commands Tab](https://github.com/tc45/images/blob/master/GetInventory_tabs_commands.jpg)

### Settings tab

Certain built in functions can be toggled on and off depending on the requirements of the project.  Click the 'Settings'
tab and choose the options relevant.

| Value |  Default | Notes |
| --- | --- | --- |
| Gather version info | Yes |
| Gather ARP tab info | Yes |
| Gather MAC tab info | Yes |
| Gather interface tab info | Yes |
| Gather CDP tab info | Yes |
| Gather LLDP tab info | Yes |
| Gather route info | Yes |
| Gather BGP info | Yes |
| Gather inventory | Yes |
| Gather commands | Yes | Gather commands tab |
|Number of Concurrent Connections| 10 | Change the maximum number of concurrent connections to device to capture all the required show commands, max of 100 |
| Gather logging (to CSV) | FUTURE | Output logging |

![Download GetInventory](https://github.com/tc45/images/blob/master/GetInventory_tabs_settings.jpg)

### Executing the script

To execute the script, drop to a command prompt and navigate to the script directory.  Launch python with the command line argument 'main.py' to execute the script.  If no command line arguements are applied, the script will look in the local directory for the 'GetInventory - Default.xlsx' file to load as the source.  Command line arguments listed below can be supplied at runtime to override some of the behavior of the spreadsheet.

![Execute script - GetInventory](https://github.com/tc45/images/blob/master/GetInventory_steps_execute.gif)


### Command line arguments

All global options in the Spreadsheet can be overriden with command line options which are revelead with either a -h
or --help flag.  

| Arguement | Flag | Description |
| --- | --- | --- |
| Username | -u \<USERNAME\> | Override global username |
| Password | -p \<PASSWORD\> | Override global password |
| Secret | -s \<SECRET\> | Override global secret |
| Input file | -i \<XLS_INPUT_FILE\> | Override default input file |
| Output directory | -o \<OUTPUT_DIRECTORY\> | Override file output directory |
| Output file | -f \<XLS_OUTPUT_FILE\> | Override output file |

```
Usage: main.py [options]

Options:
  -h, --help            show this help message and exit
  -v, --verbose         Enable Verbose Output
  -r, --raw_cli_output  Capture the raw CLI output
  -i INPUT_FILE, --input_file=INPUT_FILE
                        Input file name of excel sheet
  -o OUTPUT_FILE, --output_file=OUTPUT_FILE
                        Output file name of excel sheet
  -d OUTPUT_DIR, --output_directory=OUTPUT_DIR
                        Output Directory of excel sheet
  -u USERNAME, --username=USERNAME
                        Global Username
  -p PASSWORD, --password=PASSWORD
                        Global Password
  -s SECRET, --secret=SECRET
                        Global Secret

```

**Override input file**
```
c:\Python\GetInventory>python main.py -i "MY_CUSTOM_XLS_INPUT.xlsx"
```

**Override username/password**
```
c:\Python\GetInventory>python main.py -u "USER123" -p "ITSASECRET"
```

**Override output file and directory**
```
c:\Python\GetInventory>python main.py -o "c:\my documents\Project123" -f "Site1.xlsx"
```

### Open output file

The output file will be in the directory specified in either the spreadsheet or via command line.

Depending on the options selected under settings check the following data tabs for output:

| Tab Name | Items |
| ---- | ---- |
| Main | Original input, Make, Model, Interface Count, etc |
| Inventory | Part ID, Device, Description, Serial |
| Interfaces | IF, Description, Type, VRF, Trunk/Access, VLAN, etc |
| Routes | VRF, Protocol, Route, Subnet, CIDR, Next Hop IP, Distance, Metric, Uptime |
| BGP | Status, Path Selection, Route Source, Network, Next Hop, Metric, etc |
| ARP | VRF, IP Address, Age, Hardware/MAC, Type, Interface |
| MAC | Destination Address, Type, VLAN, Destination Port |
| CDP | Local Port, remote Port, Remote Host, Interface IP, MGMT IP, etc |
| LLDP | Chassis ID, Local Port, Remote Host, Remote Host, etc |
| Errors | Any errors encountered while executing parsing commands or connectivity issues |

## Future additions

Below are some of the planned future enhancements.  I welcome any assistance.

* Cross OS Support (JUNOS, Aruba, Palo Alto, Cisco ASA, etc)
* OSPF table parsing
* EIGRP Table parsing
* Output log to CSV
* CDP Crawling
* Network Discovery
* Subnet input instead of single IPs

## Built With

* [Netmiko](https://pynet.twb-tech.com/blog/automation/netmiko.html) - SSH/Telnet connection handler
* [Openpyxl](https://openpyxl.readthedocs.io/en/stable/) - Excel document module
* [ntcTemplates](https://github.com/networktocode/ntc-templates) - network to code templates (JSON parsers)
* [textFSM](https://github.com/google/textfsm) - raw data parser to JSON (Included in Netmiko)

## Contributing

Please reach out to Tony if you are interested in contributing to this project.  

## Authors

* **Tony Curtis** - *Initial work* - [Github-tc45](https://github.com/tc45)
Network architect turned automation programmer
* **Ruben GuitierrezMartinez** - [Github-rubengm13](https://github.com/rubengm13)
Network engineer/NetDevOps developer

## License

Licensed under the MIT license.

## Acknowledgments

* To all the engineers that have run the tool and provided feedback to improve it.

