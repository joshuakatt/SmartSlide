# test_textScraper.py
import unittest
from unittest.mock import patch, call
import subprocess

# Import the function you want to test from the module you want to test.
# You might need to adjust the import statement depending on the structure of your project.
from ..main import change_slide

class TestSlideChange(unittest.TestCase):

    @patch('subprocess.Popen')
    def test_change_slide(self, mock_popen):
        mock_popen.return_value.communicate.return_value = (b'output', b'')
        change_slide(1)
        expected_command = '''
        tell application "Microsoft PowerPoint"
            activate
            show slide 2 of active presentation
        end tell
        '''
        mock_popen.assert_called_with(["osascript", "-e", expected_command], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

if __name__ == '__main__':
    unittest.main()
