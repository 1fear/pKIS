import unittest

from sheets import format_google_sheets_error


class GoogleErrorMessageTests(unittest.TestCase):
    def test_permission_error_gets_actionable_message(self):
        message = format_google_sheets_error(PermissionError())

        self.assertIn("Нет доступа к Google-таблице", message)
        self.assertIn("service account", message)

    def test_invalid_jwt_gets_actionable_message(self):
        message = format_google_sheets_error(
            RuntimeError("invalid_grant: Invalid JWT Signature.")
        )

        self.assertIn("Google-ключ", message)
        self.assertIn("Invalid JWT Signature", message)


if __name__ == "__main__":
    unittest.main()
