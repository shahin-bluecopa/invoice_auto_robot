from rpa.invoice_robot import InvoiceRobot
import sys
from bluecopa_rpa_sdk.entrypoint import launch

if __name__ == "__main__":
    source = InvoiceRobot()
    launch(source, sys.argv[1:])
