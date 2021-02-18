from .networkdevice import NetworkDevice


def run_gathers(net_dev: NetworkDevice):
    """
    Function to call all the NetworkDevice methods to capture all the data.
    """
    try:
        net_dev.start_connection()
        net_dev.update_dev_info()
        net_dev.end_connection()
    except Exception as e:
        net_dev.add_exception_error(e)
