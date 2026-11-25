import json
import os
import sys
import tempfile
import unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'pipeline'))
from marketmap_estimates import build_local, month

class LocalTests(unittest.TestCase):
    def test_effective_months_negative_and_actual_zero(self):
        y = 2024
        capa = [{'product': 'WFC - Commodity', 'location_desc': '중국', 'capa_mty': 1200,
                 'start': month(y, 2), 'end': month(y, 2)}]
        cn = [{'y': y, 'm': 1, 'v': 0}, {'y': y, 'm': 2, 'v': 100}, {'y': y, 'm': 3, 'v': 0}]
        actuals = [{'country': '중국', 'start': month(y, 1), 'end': month(y, 1),
                    'product': 'CWF', 'value_t': 0, 'source': 'test'}]
        out = build_local(capa, {'중국': 0}, cn, [], '', actuals=actuals)
        rows = {r['start']: r for r in out['entries']}
        self.assertEqual(rows[month(y, 1)]['status'], 'actual')
        self.assertEqual(rows[month(y, 1)]['value_t'], 0)
        self.assertAlmostEqual(rows[month(y, 2)]['value_t'], -10)
        self.assertEqual(rows[month(y, 2)]['status'], 'conflict')
        self.assertEqual(rows[month(y, 3)]['value_t'], 0)

    def test_japan_annual_world_only_and_no_monthly_extension(self):
        with tempfile.TemporaryDirectory() as directory:
            data_dir = os.path.join(directory, '데이터_공급국')
            os.makedirs(data_dir)
            for year in (2023, 2024, 2025):
                data = [{'partnerCode': 0, 'cmdCode': code, 'netWgt': 1000} for code in ('481013', '481014', '481019')]
                data += [{'partnerCode': 1, 'cmdCode': '481013', 'netWgt': 999999}]
                with open(os.path.join(data_dir, f'comtrade_JP_X_{year}.json'), 'w', encoding='utf8') as f:
                    json.dump({'data': data}, f)
            capa = [{'product': 'WFC - Commodity', 'location_desc': '일본', 'capa_mty': 120}]
            result = build_local(capa, {'일본': 0}, [], [], directory)
            self.assertEqual(len(result['entries']), 3)
            for entry in result['entries']:
                self.assertEqual(entry['end'] - entry['start'], 11)
                self.assertEqual(entry['exports_t'], 3)
                self.assertAlmostEqual(entry['value_t'], 105)
                self.assertEqual(entry['status'], 'reference')

    def test_mismatched_product_does_not_override(self):
        capa = [{'product': 'WFC - Commodity', 'location_desc': '중국', 'capa_mty': 1200}]
        actuals = [{'country': '중국', 'start': month(2024, 1), 'end': month(2024, 1),
                    'product': 'other', 'value_t': 1, 'source': 'test'}]
        result = build_local(capa, {'중국': 0}, [{'y': 2024, 'm': 1, 'v': 0}], [], '', actuals=actuals)
        self.assertEqual(result['entries'][0]['status'], 'reference')
        self.assertAlmostEqual(result['entries'][0]['value_t'], 90)

if __name__ == '__main__':
    unittest.main()
