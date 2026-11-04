"""Portable checks for demo fixtures, animation budgets and documentation wiring."""

import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from tools.generate_feature_demos import (
    DEMOS, HEIGHT, MAX_BYTES, ROOT, WIDTH, exposed_regions, gif_frame_duration, inspect_gif,
    prepare_fixture, prepare_launcher,
)


class DemoFixtureTests(unittest.TestCase):
    def test_long_recordings_are_accelerated_without_losing_frames(self):
        self.assertEqual(gif_frame_duration(60), 200)
        for count in (91, 163, 300, 900):
            with self.subTest(count=count):
                duration = gif_frame_duration(count)
                self.assertEqual(duration % 10, 0)
                self.assertLessEqual(count * duration, 18000)
                self.assertGreaterEqual(duration, 20)
        with self.assertRaises(ValueError):
            gif_frame_duration(0)

    def test_missing_launcher_compiler_fails_before_building_or_launching(self):
        with tempfile.TemporaryDirectory() as temporary, patch.dict(os.environ, WINDIR=temporary):
            with patch('tools.generate_feature_demos.subprocess.run') as run:
                with self.assertRaisesRegex(RuntimeError, 'C# compiler'):
                    prepare_launcher(Path(temporary))
                run.assert_not_called()
                self.assertFalse((Path(temporary) / 'Demo-Tools').exists())

    def test_homepage_starts_with_basics_then_tools_then_advanced_features(self):
        homepage = (ROOT / 'README.md').read_text(encoding='utf-8')
        self.assertLess(homepage.index('### 基础上手'), homepage.index('### 效率工具'))
        self.assertLess(homepage.index('### 效率工具'), homepage.index('### 进阶功能'))
        self.assertLess(homepage.index('(docs/demos/tabs.gif)'), homepage.index('(docs/demos/split.gif)'))
        self.assertTrue({'tab-basics', 'bookmarks', 'folder-drop', 'launcher', 'screenshot'}.issubset(DEMOS))

    def test_privacy_masks_leave_foreground_dialogs_uncovered(self):
        rect = (0, 0, 400, 600)
        cover = (100, 100, 500, 500)
        regions = exposed_regions(rect, [cover])
        self.assertEqual(sum((r - l) * (b - t) for l, t, r, b in regions), 120000)
        for left, top, right, bottom in regions:
            self.assertTrue(right <= cover[0] or left >= cover[2] or bottom <= cover[1] or top >= cover[3])
        self.assertEqual(exposed_regions(rect, [rect]), [])
        self.assertEqual(exposed_regions(rect, [(500, 500, 600, 600)]), [rect])

    def test_fixture_is_self_contained_and_disables_external_services(self):
        with tempfile.TemporaryDirectory() as temporary:
            data = Path(temporary)
            prepare_fixture(data)
            config = json.loads((data / 'config' / 'config.json').read_text(encoding='utf-8'))
            self.assertFalse(config['enable_explorer_monitor'])
            self.assertFalse(config['auto_update_check'])
            self.assertFalse(config['hotkeys']['screenshot'])
            self.assertEqual(config['cached_tabs'], [{'path': str(data / 'Project-A')}])
            for folder in ('Project-A', 'Project-B', 'Documents'):
                self.assertTrue((data / folder / 'README.txt').is_file())
                source = (data / folder / 'src' / 'main.py').read_text(encoding='utf-8')
                self.assertIn('TabEx demo match', source.splitlines()[1])

    def test_every_demo_is_linked_from_homepage_and_guide(self):
        homepage = (ROOT / 'README.md').read_text(encoding='utf-8')
        guide = (ROOT / 'docs' / 'FEATURE_DEMOS.md').read_text(encoding='utf-8')
        for name in DEMOS:
            with self.subTest(demo=name):
                self.assertIn(f'(docs/demos/{name}.gif)', homepage)
                self.assertIn(f'(demos/{name}.gif)', guide)
                self.assertTrue((ROOT / 'docs' / 'demos' / f'{name}.gif').is_file())


@unittest.skipUnless(importlib.util.find_spec('PIL'), 'Install requirements-demo.txt to validate GIF encoding')
class DemoGifTests(unittest.TestCase):
    def setUp(self):
        from PIL import Image
        self.Image = Image
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.path = Path(self.temporary.name) / 'sample.gif'

    def write_gif(self, size=(WIDTH, HEIGHT + 84), duration=4000, loop=0):
        first = self.Image.new('RGB', size, 'white')
        second = self.Image.new('RGB', size, 'blue')
        first.save(self.path, save_all=True, append_images=[second], duration=duration, loop=loop)

    def test_accepts_exact_duration_boundary(self):
        for duration in (4000, 10000):
            with self.subTest(duration=duration):
                self.write_gif(duration=duration)
                self.assertIn(f'{duration / 500:.1f}s', inspect_gif(self.path))

    def test_rejects_short_and_long_animations(self):
        for duration in (3990, 10010):
            with self.subTest(duration=duration):
                self.write_gif(duration=duration)
                with self.assertRaisesRegex(AssertionError, '8-20 seconds'):
                    inspect_gif(self.path)

    def test_rejects_wrong_dimensions_and_non_looping_animation(self):
        for options in ({'size': (100, 100)}, {'loop': 1}):
            with self.subTest(options=options):
                self.write_gif(**options)
                with self.assertRaisesRegex(AssertionError, 'dimensions or loop'):
                    inspect_gif(self.path)

    def test_rejects_static_image(self):
        self.Image.new('RGB', (WIDTH, HEIGHT + 84), 'white').save(self.path)
        with self.assertRaisesRegex(AssertionError, '8-20 seconds'):
            inspect_gif(self.path)

    def test_rejects_oversized_file(self):
        self.write_gif()
        with self.path.open('ab') as stream:
            stream.write(b'\0' * MAX_BYTES)
        with self.assertRaisesRegex(AssertionError, '3 MiB'):
            inspect_gif(self.path)

    def test_delivered_gifs_meet_all_budgets(self):
        for name in DEMOS:
            with self.subTest(demo=name):
                inspect_gif(ROOT / 'docs' / 'demos' / f'{name}.gif')


if __name__ == '__main__':
    unittest.main()
