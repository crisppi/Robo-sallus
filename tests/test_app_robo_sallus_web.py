import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from app_robo_sallus_web import HTML, default_state  # noqa: E402


class AppRoboSallusWebTests(unittest.TestCase):
    def test_execution_timer_state_and_card_exist(self):
        state = default_state()
        self.assertIsNone(state["execution_started_at"])
        self.assertIsNone(state["execution_finished_at"])
        self.assertIn('id="total_timer"', HTML)
        self.assertIn("formatDuration", HTML)

    def test_launch_rows_are_sorted_with_active_and_latest_first(self):
        self.assertIn("orderedLaunchRows", HTML)
        self.assertIn("activeDifference", HTML)
        self.assertIn("launchTimestamp(b) - launchTimestamp(a)", HTML)


if __name__ == "__main__":
    unittest.main()
