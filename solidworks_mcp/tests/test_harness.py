"""
Mocked-COM test harness for the SolidWorks MCP server.

SolidWorks COM automation is Windows-only, so these tests run on Linux
against fake COM objects. The harness stubs ``win32com`` / ``pythoncom``
in ``sys.modules`` BEFORE importing the package, then exercises the
automation mixins with a fake document that records API calls.

Run:  python3 -m unittest discover -s solidworks_mcp/tests -v
      (from the repository root)
"""

import sys
import types
import unittest
from unittest.mock import MagicMock

# ---------------------------------------------------------------------------
# Stub win32com / pythoncom before any package import
# ---------------------------------------------------------------------------

win32com = types.ModuleType("win32com")
win32com_client = types.ModuleType("win32com.client")


class _FakeVariant:
    def __init__(self, *args, **kwargs):
        self.args = args


def _fake_variant(*args, **kwargs):
    return _FakeVariant(*args, **kwargs)


win32com_client.VARIANT = _fake_variant
win32com_client.GetObject = MagicMock(return_value=None)
win32com_client.Dispatch = MagicMock(return_value=None)
win32com_client.dynamic = types.SimpleNamespace(Dispatch=MagicMock(return_value=None))
win32com.client = win32com_client

pythoncom = types.ModuleType("pythoncom")
pythoncom.VT_DISPATCH = 9
pythoncom.VT_I4 = 3
pythoncom.VT_BYREF = 0x4000
pythoncom.Nothing = None
pythoncom.CoInitialize = lambda: None

sys.modules.setdefault("win32com", win32com)
sys.modules.setdefault("win32com.client", win32com_client)
sys.modules.setdefault("pythoncom", pythoncom)

from solidworks_mcp.automation import SolidWorksAutomation  # noqa: E402
from solidworks_mcp.utils import UnitConverter  # noqa: E402


# ---------------------------------------------------------------------------
# Fake COM objects
# ---------------------------------------------------------------------------

class FakeFeature:
    """A node in a fake feature tree (linked list via GetNextFeature)."""

    def __init__(self, name, type_name, nxt=None, suppressed=False):
        self.Name = name
        self.GetTypeName2 = type_name
        self.GetNextFeature = nxt
        self._suppressed = suppressed

    def IsSuppressed(self):
        return self._suppressed


class FakeSketchManager:
    def __init__(self):
        self.calls = []
        self.ActiveSketch = None

    def CreateCircle(self, x, y, z, x2, y2, z2):
        self.calls.append(("CreateCircle", (x, y, z, x2, y2, z2)))
        return object()

    def CreateLine(self, x1, y1, z1, x2, y2, z2):
        self.calls.append(("CreateLine", (x1, y1, z1, x2, y2, z2)))
        return object()

    def Create3PointArc(self, x1, y1, z1, x2, y2, z2, x3, y3, z3):
        self.calls.append(("Create3PointArc",
                           (x1, y1, z1, x2, y2, z2, x3, y3, z3)))
        return object()

    def CreateSpline3(self, point_data, surfs, direction, natural_ends):
        self.calls.append(("CreateSpline3",
                           (point_data, surfs, direction, natural_ends)))
        return object()

    def CreateSketchSlot(self, slot_type, length_type, width,
                         x1, y1, z1, x2, y2, z2, x3, y3, z3,
                         arc_direction, add_dimension):
        self.calls.append(("CreateSketchSlot",
                           (slot_type, length_type, width,
                            x1, y1, z1, x2, y2, z2, x3, y3, z3,
                            arc_direction, add_dimension)))
        return object()

    def SketchTrim(self, option, x, y, z):
        self.calls.append(("SketchTrim", (option, x, y, z)))
        return True

    def InsertSketch(self, *args):
        self.calls.append(("InsertSketch", args))


class FakeExtension:
    def __init__(self):
        self.selected = []

    def SelectByID2(self, name, sel_type, x, y, z, append, mark, callout, options):
        self.selected.append((name, sel_type))
        return True


class FakeDoc:
    """Minimal fake ModelDoc2 supporting the calls the mixins make."""

    def __init__(self, first_feature=None):
        self.FirstFeature = first_feature
        self.SketchManager = FakeSketchManager()
        self.Extension = FakeExtension()
        self.FeatureManager = MagicMock()
        self._cleared = False

    def ClearSelection2(self, *args):
        self._cleared = True

    def InsertSketch2(self, *args):
        pass


def make_automation(doc=None):
    """Build a SolidWorksAutomation wired to a fake doc (no COM needed)."""
    auto = SolidWorksAutomation.__new__(SolidWorksAutomation)
    auto._sw_app = types.SimpleNamespace(ActiveDoc=doc, RevisionNumber="33.0")
    auto._connected = True
    auto._units = UnitConverter("mm")
    return auto


class TestHarness(unittest.TestCase):
    def test_package_imports_with_stubbed_com(self):
        self.assertIsNotNone(SolidWorksAutomation)

    def test_result_shape(self):
        auto = make_automation()
        r = auto._result(True, "ok")
        self.assertTrue(r["success"])
        self.assertEqual(r["error_code"], 0)
        self.assertEqual(r["error_name"], "swSuccess")
        self.assertIn("timestamp", r)

    def test_units_mm_to_meters(self):
        conv = UnitConverter("mm")
        self.assertAlmostEqual(conv.to_meters(25, "mm"), 0.025)
        self.assertAlmostEqual(conv.to_meters(1, "inch"), 0.0254)
        self.assertAlmostEqual(conv.from_meters(0.05, "mm"), 50.0)

    def test_draw_circle_converts_to_meters(self):
        doc = FakeDoc()
        auto = make_automation(doc)
        r = auto.draw_circle(10, 20, 25, unit="mm")
        self.assertTrue(r["success"], r)
        name, args = doc.SketchManager.calls[0]
        self.assertEqual(name, "CreateCircle")
        # center (10mm, 20mm) -> (0.01, 0.02) m; edge point x = center + radius
        self.assertAlmostEqual(args[0], 0.01)
        self.assertAlmostEqual(args[1], 0.02)
        self.assertAlmostEqual(args[3], 0.035)

    def test_draw_polygon_rejects_bad_sides(self):
        auto = make_automation(FakeDoc())
        r = auto.draw_polygon(0, 0, 25, sides=2, unit="mm")
        self.assertFalse(r["success"])
        self.assertEqual(r["error_code"], 107)  # swInvalidInput

    def test_find_last_sketch_walks_tree(self):
        sketch2 = FakeFeature("Sketch2", "ProfileFeature")
        boss = FakeFeature("Boss-Extrude1", "Extrusion", nxt=sketch2)
        sketch1 = FakeFeature("Sketch1", "ProfileFeature", nxt=boss)
        auto = make_automation(FakeDoc(first_feature=sketch1))
        self.assertEqual(auto._find_last_sketch(FakeDoc(sketch1)), "Sketch2")

    def test_extrude_without_sketch_fails_gracefully(self):
        auto = make_automation(FakeDoc(first_feature=None))
        r = auto.extrude_sketch(10)
        self.assertFalse(r["success"])
        self.assertIn("No sketch", r["message"])


if __name__ == "__main__":
    unittest.main()
