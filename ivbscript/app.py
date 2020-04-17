from jupyter_console.app import ZMQTerminalIPythonApp


def main():
    ZMQTerminalIPythonApp.launch_instance(argv=['--kernel', 'vbscript'])


if __name__ == '__main__':
    main()
