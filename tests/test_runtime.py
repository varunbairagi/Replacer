import unittest

from landdocs.runtime import check_tk_version


class RuntimeChecks(unittest.TestCase):
    def test_old_mac_tk_is_rejected_before_a_window_can_open(self):
        with self.assertRaisesRegex(RuntimeError, "blank white window"):
            check_tk_version(8.5, "darwin")

    def test_supported_mac_tk_and_other_platforms_are_accepted(self):
        for version, platform in ((8.6, "darwin"), (9.0, "darwin"), (8.5, "linux"), (8.5, "win32")):
            with self.subTest(version=version, platform=platform):
                check_tk_version(version, platform)


if __name__ == "__main__":
    unittest.main()
