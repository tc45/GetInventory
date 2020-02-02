import textfsm
import getopt, sys


def main():
    str_days = '859 day(s), 6 hour(s), 48 minute(s), 41 second(s)'
    str_new_format = format_uptime(str_days)
    print(str_new_format)

    test_textfsm()

    inputfile = ''
    outputfile = ''
    failed_args = False



    try:
        opts, args = getopt.getopt(sys.argv[1:], "hi:o:", ["ifile=","ofile"])
    except getopt.GetoptError:
        print("test_textfsm.py -i <inputfile> -o <outputfile>")
        sys.exit(2)
        failed_args = True

    if failed_args is False:
        for opt, arg in opts:
            if opt == "-h":
                print("test_textfsm.py -i <inputfile> -o <outputfile>")
            elif opt in ("-i", "--ifile"):
                inputfile = arg
            elif opt in ("-o", "--ofile"):
                outputfile = arg

        print("Input file is " + inputfile)
        print("Input file is " + outputfile)


def test_textfsm():
    print("\nLinux Unix (Intel-x86) processor with 863257K bytes of memory.\n\n")
    asr_input_file = open("sample_ios_version_ASR.txt", encoding="utf-8")
    asr_raw_text_data = asr_input_file.read()
    asr_input_file.close()
    c800_input_file = open("sample_ios_version_c800.txt", encoding="utf-8")
    c800_raw_text_data = c800_input_file.read()
    c800_input_file.close()
    isr4k_input_file = open("sample_ios_version_isr4k.txt", encoding="utf-8")
    isr4k_raw_text_data = isr4k_input_file.read()
    isr4k_input_file.close()
    virtual_input_file = open("sample_ios_version_virtual.txt", encoding="utf-8")
    virtual_raw_text_data = virtual_input_file.read()
    virtual_input_file.close()

    # Run the string through textFSM
    template = open("ntc-templates/templates/cisco_ios_show_version.textfsm")
    asr_re_table = textfsm.TextFSM(template)
    c800_re_table = textfsm.TextFSM(template)
    isr4k_re_table = textfsm.TextFSM(template)
    virtual_re_table = textfsm.TextFSM(template)
    print("ASR DATA:\n" + str(asr_re_table.ParseText(asr_raw_text_data)))
    print("C800 DATA:\n" + str(c800_re_table.ParseText(c800_raw_text_data)))
    print("ISR4k DATA:\n" + str(isr4k_re_table.ParseText(isr4k_raw_text_data)))
    print("Virtual DATA:\n" + str(virtual_re_table.ParseText(virtual_raw_text_data)))


def format_uptime(uptime):
    str_years, str_weeks, str_days, str_hours, str_minutes = 0, 0, 0, 0, 0
    str_input = uptime.split(",")
    for i in str_input:
        i = i.strip()
        str_split = i.split(" ")
        if left(str_split[1], 3) == "yea":
            str_years = int(str_split[0])
        if left(str_split[1], 3) == "wee":
            str_weeks = int(str_split[0])
        if left(str_split[1], 3) == "day":
            str_days = int(str_split[0])
        if left(str_split[1], 3) == "hou":
            str_hours = int(str_split[0])
        if left(str_split[1], 3) == "min":
            str_minutes = int(str_split[0])

    if str_days > 365:
        years = str_days / 365
        if not years.is_integer():
            years = int(str(years).split(".")[0])
        str_days = str_days - years * 365
        str_years = str_years + years
    if str_days > 7:
        weeks = str_days / 7
        if not weeks.is_integer():
            weeks = int(str(weeks).split(".")[0])
        str_days = str_days - weeks * 7
        str_weeks = str_weeks + weeks
    if str_weeks > 52:
        years = str_weeks/52
        if not years.is_integer():
            years = years.split(".")
            years = years[0]
        str_weeks = str_weeks - years * 52
        str_years = str_years + years

    return (str(str_years) + "y " +
            str(str_weeks) + "w " +
            str(str_days) + "d " +
            str(str_hours) + "h " +
            str(str_minutes) + "m "
            )


def left(s, amount):
    return s[:amount]


def right(s, amount):
    return s[-amount:]


def mid(s, offset, amount):
    return s[offset:offset+amount]


main()
