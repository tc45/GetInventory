from .networkdevice import NetworkDevice


def run_gathers(net_dev: NetworkDevice):
    """
    Function to call all the NetworkDevice methods to capture all the data.
    """
    try:
        net_dev.update_time('start')
        net_dev.start_connection()
        net_dev.update_dev_info()
        net_dev.go_gather()
        net_dev.end_connection()
        net_dev.update_time('end')
        if net_dev.status.lower() == 'yes':
            net_dev.status = 'Complete'
    except Exception as e:
        net_dev.add_exception_error(e)
