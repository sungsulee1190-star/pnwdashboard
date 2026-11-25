import os
import sys
import unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'pipeline'))
from marketmap_estimates import canonical, month, split_kr

class SplitTests(unittest.TestCase):
    def test_split_preserves_quantity_and_original_amount(self):
        m = month(2023, 8)
        vol = {(m, '미국', 1): 60, (m, '미국', 2): 10, (m, '미국', 4): 30}
        amt = {(m, '미국', 1): 600, (m, '미국', 2): 100, (m, '미국', 4): 300}
        rows, original = split_kr(vol, amt, {'미국': 0}, {'미국': [20, 100, True]})
        self.assertEqual(rows[0]['status'], 'estimated')
        self.assertAlmostEqual(vol[(m, '미국', 3)], 20)
        self.assertAlmostEqual(sum(vol.values()), 100)
        self.assertEqual(original[0]['amount_usd'], 1000)
        self.assertEqual(original[0]['series_amount_usd'], [600, 100, 0, 300])

    def test_excess_preserves_pool(self):
        m = month(2023, 8)
        vol = {(m, '미국', 1): 90, (m, '미국', 4): 10}
        rows, _ = split_kr(vol, {}, {'미국': 0}, {'미국': [50, 100, True]})
        self.assertEqual(rows[0]['status'], 'conflict')
        self.assertEqual(vol[(m, '미국', 4)], 10)

    def test_zero_missing_and_boundary(self):
        old, new = month(2023, 7), month(2023, 8)
        vol = {(old, 'A', 4): 5, (new, 'A', 4): 5, (new, 'B', 4): 5}
        rows, _ = split_kr(vol, {}, {'A': 0, 'B': 1}, {'A': [0, 100, True], 'B': [0, 100, False]})
        self.assertEqual(vol[(old, 'A', 4)], 5)
        self.assertEqual([r['status'] for r in rows], ['confirmed_zero', 'missing'])
        self.assertEqual(vol[(new, 'B', 4)], 5)

    def test_aliases(self):
        self.assertEqual(canonical('인도(인디아)'), '인도')
        self.assertEqual(canonical('터키'), '튀르키예')

if __name__ == '__main__':
    unittest.main()
