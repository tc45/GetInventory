# GetInventory

GetInventory is a network device collector script.  The purpose of this tool is to quickly and efficiently allow 
network engineers and operators to be more efficient in gathering details from network devices.  Two primary sets out output will be 
produced per device.  The first is an Excel spreadsheet.  The spreadsheet serves as the input for the application, copies itself to an 
output directory, and updates the output file with a standard set of information collected per device.  The information gathered is 
parsed and added into tables on individual tabs within the file.  In addition to the Excel file, 
each device can have a list of user input commands added into the 'Command' tab which will then be run against every device and a 
unique file will be output per device.  These can be any standard command that the platform will support.  


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

## Platform Support

GetInventory is intended to be cross-platform supportable, but has only been tested on Windows platforms.  Any testing/coding help 
to test on LINUX or MAC would be greatly appreciated.

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
| Username | X | Spreadsheet or command line | SSH/Telnet username |
| Password | X | Spreadsheet or command line | SSH/Telnet password |
| Secret | X | Spreadsheet or command line | Enable password (Defaults to regular password if not specified) |
| Output Directory | X |  Spreadsheet or command line | Directory where output spreadsheet and command files will be stored |
| Output Filename | X | Spreadsheet or command line | Specifies name of output spreadsheet |
| Parse Method | | Spreadsheet only | Specifies 'device type' to be parsed.  Defaults to cisco_ios) |
| Protocol | | Spreadsheet only | Specify SSH or Telnet (Defaults to SSH if not specified) | 
| Port | | Spreadsheet only | Specify port if not standard port (Defaults 22 for SSH and 23 for Telnet |
| Device Username | | Spreadsheet only | Per device override for username - COMING SOON |
| Device Password | | Spreadsheet only | Per device override for password - COMING SOON |


### Open spreadsheet

Open the default spreadsheet, **GetInventory - Default.xlsx**, in the default directory.  The file should default to the Main tab, 
but if not go ahead and click on the 'Main' tab.  A sample of the beginning file is shown below.

![](https://github.com/tc45/images/GetInventory_Tab_Main_3.26.2020.jpg)



**Note:** Any columns with a black header are currently not functioning.**

### Add General Connectivity Details

The following values can be specified in the spreadsheet or added at run time to the command line.  See section below on ** Command Line 
Options ** for additional details on how to add via command line.  Any values specified at command line will override the value 
in the spreadsheet. 

Chagne the following values to the requirements of your own project.

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
| Gather logging (to CSV) | FUTURE | Output logging

### Executing the script

To execute the script, drop to a command prompt and navigate to the script directory.  Launch python with the command line argument 'main.py' to execute the script.  If no command line arguements are applied, the script will look in the local directory for the 'GetInventory - Default.xlsx' file to load as the source.  Command line arguments listed below can be supplied at runtime to override some of the behavior of the spreadsheet. 

![Execute script - GetInventory](https://github.com/tc45/images/GetInventory_Steps_execute.gif)


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


![](https://github.com/tc45/images/GetInventory_CLI_Arguements.jpg)

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
* Multithreading
* CDP Crawling
* Network Discovery
* Subnet input instead of single IPs

## Deployment

Add additional notes about how to deploy this on a live system

## Built With

* [Dropwizard](http://www.dropwizard.io/1.0.2/docs/) - The web framework used
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Tony Curtis** - *Initial work* - [Github-tc45](https://github.com/tc45)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* To all the engineers that have run the tool and provided feedback to improve it.
* Inspiration
* etc
