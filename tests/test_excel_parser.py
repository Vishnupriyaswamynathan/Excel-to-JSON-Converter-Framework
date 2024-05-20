import unittest
from excel_parser import ExcelParser
from config_manager import ConfigManager

class TestExcelParser(unittest.TestCase):
    def setUp(self):
        config_manager = ConfigManager('config.yaml')
        config = config_manager.load_config()
        self.parser = ExcelParser(config)

    def test_parse(self):
        data = self.parser.parse('test.xlsx')
        self.assertIsNotNone(data)

if __name__ == '__main__':
    unittest.main()
