import unittest
from src.main import main

class TestMain(unittest.TestCase):
    def test_main_runs(self):
        # This test just checks that main() runs without error
        main()

if __name__ == "__main__":
    unittest.main()
