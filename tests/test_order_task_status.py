import unittest
from unittest.mock import patch, MagicMock

from fangzheng_web_app import order_entry_service as service


class OrderTaskStatusTests(unittest.TestCase):
    def test_entry_completion_and_reply_are_independent(self):
        for completed, replied, label in (
            (False, False, '待录单'),
            (False, True, '待录单'),
            (True, False, '录单完成待回复'),
            (True, True, '录单完成已回复'),
        ):
            with self.subTest(completed=completed, replied=replied):
                cursor = MagicMock()
                cursor.__enter__.return_value.execute.return_value.fetchone.return_value = {'id': 1} if replied else None
                with patch.object(service, '_entry_progress', return_value={'completed': completed, 'label': '待批量料号查询'}), patch.object(service, 'db_cursor', return_value=cursor):
                    result = service.template_progress(1, 'test-user')
                self.assertEqual(result['label'], label)
                self.assertEqual(result['closed'], completed)
                self.assertEqual(result['replied'], replied)


if __name__ == '__main__':
    unittest.main()
