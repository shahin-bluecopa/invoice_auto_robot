import logging
from typing import Mapping, Any, Union, List, MutableMapping

from bluecopa_rpa_sdk.robots.abstract_robot import AbstractRobot
from bluecopa_rpa_sdk.utils.robot_protocol import RobotStateMessage

from .tasks import process_invoices


class InvoiceRobot(AbstractRobot):

    def get_config_spec(self) -> dict:
        return {}

    def run_robot(self, logger: logging.Logger, config: Mapping[str, Any], input_file_path: str, output_folder_path: str, state: Union[List[RobotStateMessage], MutableMapping[str, Any]] = None):
        logger.info(f"Starting {self.name}")
        
        process_invoices(logger)
        
        logger.info(f"Finished {self.name}")
