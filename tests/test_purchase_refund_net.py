import unittest

from vinted_mail_to_excel import _build_cancel_remaining, _safe_imap_logout, _should_use_imap


class TestPurchaseRefundNeutralization(unittest.TestCase):
    def test_purchase_and_refund_same_title_amount_cancel_across_sources(self):
        records = [{
            "kind": "zakup",
            "title_original": "Pantaloni negri de piele cu pietre",
            "amount_total": "90.94",
        }]
        history_records = [{
            "kind": "zwrot",
            "title": "Pantaloni negri de piele cu pietre",
            "price": "90.94",
        }]

        cancel_remaining = _build_cancel_remaining(records + history_records)
        key = ("pantaloni negri de piele cu pietre", "90.94")
        self.assertEqual(cancel_remaining.get(key), 1)

        cancel_remaining[key] -= 1
        self.assertEqual(cancel_remaining.get(key), 0)

    def test_safe_imap_logout_ignores_timeout_errors(self):
        class BrokenMail:
            state = "SELECTED"

            def logout(self):
                raise TimeoutError("timed out")

        self.assertIsNone(_safe_imap_logout(BrokenMail()))

    def test_history_only_bypasses_imap(self):
        self.assertFalse(_should_use_imap("t", True))
        self.assertTrue(_should_use_imap("n", True))
        self.assertFalse(_should_use_imap("n", False))


if __name__ == "__main__":
    unittest.main()
