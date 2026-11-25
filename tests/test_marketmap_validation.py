import json
import os
import re
import sys
import unittest
ROOT = os.path.dirname(os.path.dirname(__file__))
sys.path.insert(0, ROOT)
from validate_data import check_marketmap

class PrivateDataValidationTests(unittest.TestCase):
    def setUp(self):
        data_path = os.path.join(ROOT, 'data.js')
        private_path = os.environ.get('PNW_APPROVED_RATIO_FILE', os.path.join(ROOT, '.private', 'approved_ratios.json'))
        if not os.path.isfile(data_path) or not os.path.isfile(private_path):
            self.skipTest('비공개 실자료 또는 승인 기준 없음')
        with open(data_path, encoding='utf-8') as f:
            self.content = f.read()
        self.match = re.search(r'^const DATA_MM = (.*);$', self.content, re.M)
        self.assertIsNotNone(self.match)
        self.mm = json.loads(self.match.group(1))
        with open(private_path, encoding='utf-8') as f:
            self.approved = json.load(f)

    def test_approved_country_shares_and_fractions(self):
        basis = self.mm['kr_split']['basis']['combined_by_country']
        self.assertTrue(self.approved['combined_percent'])
        for country, percent in self.approved['combined_percent'].items():
            numerator, denominator, present = basis[country]
            self.assertTrue(present, country)
            self.assertAlmostEqual(100 * numerator / denominator, percent, delta=0.0051, msg=country)
        for country, fraction in self.approved['exact_fraction'].items():
            numerator, denominator, present = basis[country]
            self.assertTrue(present, country)
            self.assertAlmostEqual(numerator, fraction['numerator'], places=4, msg=country)
            self.assertAlmostEqual(denominator, fraction['denominator'], places=4, msg=country)

    def test_swapped_series_amount_is_detected_even_when_total_is_unchanged(self):
        record = next(r for r in self.mm['kr_split']['original_kr'] if r['series_amount_usd'][0] > 1)
        record['series_amount_usd'][0] -= 1
        record['series_amount_usd'][1] += 1
        corrupted = (self.content[:self.match.start(1)] + json.dumps(self.mm, ensure_ascii=False, separators=(',', ':'))
                     + self.content[self.match.end(1):])
        errors, _ = check_marketmap(corrupted)
        self.assertTrue(any('계열별 원금액 대사 실패' in e for e in errors), errors)

if __name__ == '__main__':
    unittest.main()
