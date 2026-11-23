"""
Mocked-COM tests for the SolidWorks 2025 advanced tools.

Runs on Linux against fake COM objects (see test_harness.py for the
stub setup). Each test asserts the automation layer calls the correct
2025 API method with the correct arguments — meters, radians, enum
values and parameter order.

Run:  python3 -m unittest discover -s solidworks_mcp/tests -v
"""

import math
import sys
import types
import unittest
from unittest.mock import MagicMock

# Reuse the win32com/pythoncom stubs from test_harness (idempotent).
import test_harness  # noqa: F401
from test_harness import FakeDoc, FakeFeature, FakeSketchManager, make_automation

import win32com.client  # noqa: E402  (stubbed module)

# Fake swconst values for the CreateDefinition route tests.
win32com.client.constants = types.SimpleNamespace(
    swFmLPattern=101, swFmCirPattern=102)


class FakeFeatureManager:
    """Records feature-manager calls."""

    def __init__(self):
        self.calls = []
        self._definition = None

    def FeatureExtrusion3(self, *args):
        self.calls.append(("FeatureExtrusion3", args))
        return object()

    def FeatureExtrusion2(self, *args):
        self.calls.append(("FeatureExtrusion2", args))
        return object()

    def FeatureCut3(self, *args):
        self.calls.append(("FeatureCut3", args))
        return object()

    def FeatureRevolve2(self, *args):
        self.calls.append(("FeatureRevolve2", args))
        return object()

    def InsertProtrusionBlend2(self, *args):
        self.calls.append(("InsertProtrusionBlend2", args))
        return object()

    def InsertCutBlend(self, *args):
        self.calls.append(("InsertCutBlend", args))
        return object()

    def HoleWizard5(self, *args):
        self.calls.append(("HoleWizard5", args))
        return object()

    def InsertMirrorFeature2(self, *args):
        self.calls.append(("InsertMirrorFeature2", args))
        return object()

    def CreateDefinition(self, type_id):
        self.calls.append(("CreateDefinition", (type_id,)))
        self._definition = types.SimpleNamespace()
        return self._definition

    def CreateFeature(self, data):
        self.calls.append(("CreateFeature", (data,)))
        return object()


class FakeMeasure:
    def __init__(self, distance=0.05):
        self.Distance = distance

    def Calculate(self, entities):
        return True


class FakeMassProperty:
    Volume = 0.001
    Mass = 7.85
    SurfaceArea = 0.06
    CenterOfMass = (0.01, 0.02, 0.03)
    PrincipalMomentsOfInertia = (1.0, 2.0, 3.0)
    Density = 7850.0


class RichFakeDoc(FakeDoc):
    """FakeDoc plus a recording FeatureManager and export/measure."""

    def __init__(self, first_feature=None):
        super().__init__(first_feature)
        self.FeatureManager = FakeFeatureManager()
        self.save_calls = []

    def InsertFeatureShell(self, thickness, outward):
        self.save_calls.append(("InsertFeatureShell", (thickness, outward)))


class FakeSegment:
    def __init__(self, construction=False):
        self.ConstructionGeometry = construction
        self.select_calls = []

    def Select2(self, append, mark):
        self.select_calls.append((append, mark))
        return True


class FakeSketch:
    def __init__(self, segments):
        self._segments = segments

    def GetSketchSegments(self):
        return self._segments


def sketch_feature(name, centerline=True):
    """A ProfileFeature whose sketch has a centerline segment."""
    seg = FakeSegment(construction=centerline)
    feat = FakeFeature(name, "ProfileFeature")
    feat.GetSpecificFeature2 = lambda: FakeSketch([seg])
    return feat, seg


class TestAdvancedSketch(unittest.TestCase):
    def test_draw_arc_3point_meters(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.draw_arc_3point(0, 0, 50, 0, 25, 25, unit="mm")
        self.assertTrue(r["success"], r)
        name, args = doc.SketchManager.calls[0]
        self.assertEqual(name, "Create3PointArc")
        # start (0,0), end (0.05, 0), middle (0.025, 0.025) in meters
        self.assertAlmostEqual(args[3], 0.05)
        self.assertAlmostEqual(args[6], 0.025)
        self.assertAlmostEqual(args[7], 0.025)

    def test_draw_spline_calls_create_spline3(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.draw_spline([(0, 0), (10, 10), (20, 0)], unit="mm")
        self.assertTrue(r["success"], r)
        name, args = doc.SketchManager.calls[0]
        self.assertEqual(name, "CreateSpline3")
        # 3 points x 3 doubles, first point (0,0,0), second (0.01,0.01,0)
        self.assertEqual(len(args[0]), 9)
        self.assertAlmostEqual(args[0][3], 0.01)
        self.assertEqual(args[3], True)  # SimulateNaturalEnds

    def test_draw_slot_straight(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.draw_slot("straight", 10, 0, 0, 50, 0, unit="mm")
        self.assertTrue(r["success"], r)
        name, args = doc.SketchManager.calls[0]
        self.assertEqual(name, "CreateSketchSlot")
        self.assertEqual(args[0], 0)   # straight
        self.assertEqual(args[1], 0)   # center_center length
        self.assertAlmostEqual(args[2], 0.01)  # width in meters

    def test_draw_slot_bad_type(self):
        auto = make_automation(RichFakeDoc())
        r = auto.draw_slot("weird", 10)
        self.assertFalse(r["success"])
        self.assertEqual(r["error_code"], 107)

    def test_trim_entities(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.trim_entities(10, 20, unit="mm")
        self.assertTrue(r["success"], r)
        name, args = doc.SketchManager.calls[0]
        self.assertEqual(name, "SketchTrim")
        self.assertAlmostEqual(args[1], 0.01)
        self.assertAlmostEqual(args[2], 0.02)


class TestAdvancedFeatures(unittest.TestCase):
    def test_revolve_uses_feature_revolve2_and_mark_16(self):
        feat, seg = sketch_feature("Sketch1")
        doc = RichFakeDoc(first_feature=feat)
        auto = make_automation(doc)
        r = auto.revolve_sketch(angle=360)
        self.assertTrue(r["success"], r)
        # Axis selected with mark 16
        self.assertEqual(seg.select_calls, [(True, 16)])
        calls = dict(doc.FeatureManager.calls)
        args = dict(doc.FeatureManager.calls)["FeatureRevolve2"]
        self.assertAlmostEqual(args[8], 2 * math.pi)  # Dir1Angle radians
        self.assertEqual(args[6], 0)  # Dir1Type blind

    def test_revolve_without_centerline_fails(self):
        feat, seg = sketch_feature("Sketch1", centerline=False)
        doc = RichFakeDoc(first_feature=feat)
        auto = make_automation(doc)
        r = auto.revolve_sketch()
        self.assertFalse(r["success"])
        self.assertIn("centerline", r["message"])

    def test_loft_boss_selects_two_profiles(self):
        s2 = FakeFeature("Sketch2", "ProfileFeature")
        s1 = FakeFeature("Sketch1", "ProfileFeature", nxt=s2)
        doc = RichFakeDoc(first_feature=s1)
        auto = make_automation(doc)
        r = auto.loft_boss()
        self.assertTrue(r["success"], r)
        calls = doc.FeatureManager.calls
        self.assertEqual(calls[0][0], "InsertProtrusionBlend2")
        # TessToleranceFactor default is 1.0 (arg index 3)
        self.assertEqual(calls[0][1][3], 1.0)
        selected = [s[0] for s in doc.Extension.selected]
        self.assertEqual(selected, ["Sketch1", "Sketch2"])

    def test_loft_cut(self):
        s2 = FakeFeature("Sketch2", "ProfileFeature")
        s1 = FakeFeature("Sketch1", "ProfileFeature", nxt=s2)
        doc = RichFakeDoc(first_feature=s1)
        auto = make_automation(doc)
        r = auto.loft_cut()
        self.assertTrue(r["success"], r)
        self.assertEqual(doc.FeatureManager.calls[0][0], "InsertCutBlend")

    def test_hole_wizard_27_params(self):
        feat, seg = sketch_feature("Sketch1")
        doc = RichFakeDoc(first_feature=feat)
        auto = make_automation(doc)
        r = auto.hole_wizard(2, 0, 1, "M6", end_condition="through_all",
                             diameter=6, depth=10, unit="mm")
        self.assertTrue(r["success"], r)
        name, args = doc.FeatureManager.calls[0]
        self.assertEqual(name, "HoleWizard5")
        self.assertEqual(len(args), 27)
        self.assertEqual(args[0], 2)    # GenericHoleType
        self.assertEqual(args[3], "M6")  # SSize
        self.assertEqual(args[4], 1)    # EndType through_all
        self.assertAlmostEqual(args[5], 0.006)  # Diameter meters
        self.assertAlmostEqual(args[6], 0.01)   # Depth meters

    def test_linear_pattern_create_definition_route(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.linear_pattern("Front Plane", 20, 4,
                                ["Boss-Extrude1"], unit="mm")
        self.assertTrue(r["success"], r)
        calls = doc.FeatureManager.calls
        self.assertEqual(calls[0], ("CreateDefinition", (101,)))
        data = doc.FeatureManager._definition
        self.assertAlmostEqual(data.Spacing, 0.02)
        self.assertEqual(data.TotalInstances, 4)
        self.assertEqual(calls[-1][0], "CreateFeature")

    def test_circular_pattern_total_angle_radians(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.circular_pattern("Axis1", 180, 3, ["Boss-Extrude1"])
        self.assertTrue(r["success"], r)
        data = doc.FeatureManager._definition
        self.assertAlmostEqual(data.TotalAngle, math.pi)
        self.assertEqual(data.TotalInstances, 3)
        self.assertTrue(data.EqualSpacing)

    def test_shell_uses_insert_feature_shell(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.shell(2, [[0, 0, 10]], unit="mm")
        self.assertTrue(r["success"], r)
        self.assertEqual(doc.save_calls[0][0], "InsertFeatureShell")
        self.assertAlmostEqual(doc.save_calls[0][1][0], 0.002)
        self.assertFalse(doc.save_calls[0][1][1])  # outward=False

    def test_mirror_feature(self):
        doc = RichFakeDoc()
        auto = make_automation(doc)
        r = auto.mirror_feature(["Boss-Extrude1"], "Front Plane")
        self.assertTrue(r["success"], r)
        name, args = doc.FeatureManager.calls[0]
        self.assertEqual(name, "InsertMirrorFeature2")
        self.assertEqual(args, (False, False, True, False, 0))

    def test_extrude_uses_extrusion3_first(self):
        feat, seg = sketch_feature("Sketch1")
        doc = RichFakeDoc(first_feature=feat)
        auto = make_automation(doc)
        r = auto.extrude_sketch(10, unit="mm")
        self.assertTrue(r["success"], r)
        self.assertEqual(r["data"]["api_method"], "FeatureExtrusion3_23p")

    def test_cut_through_all_both_uses_end_cond_9(self):
        feat, seg = sketch_feature("Sketch1")
        doc = RichFakeDoc(first_feature=feat)
        auto = make_automation(doc)
        r = auto.cut_extrude(10, through_all=True, both_directions=True)
        self.assertTrue(r["success"], r)
        name, args = doc.FeatureManager.calls[0]
        self.assertEqual(name, "FeatureCut3")
        self.assertEqual(args[3], 9)  # T1 = swEndCondThroughAllBoth


class TestMeasureExport(unittest.TestCase):
    def test_measure_distance(self):
        doc = RichFakeDoc()
        doc.Extension.CreateMeasure = lambda: FakeMeasure(0.05)
        auto = make_automation(doc)
        r = auto.measure_distance(unit="mm")
        self.assertTrue(r["success"], r)
        self.assertAlmostEqual(r["data"]["distance"], 50.0)

    def test_mass_properties(self):
        doc = RichFakeDoc()
        doc.Extension.CreateMassProperty2 = lambda: FakeMassProperty()
        auto = make_automation(doc)
        r = auto.get_mass_properties(unit="mm")
        self.assertTrue(r["success"], r)
        self.assertAlmostEqual(r["data"]["mass_kg"], 7.85)
        self.assertAlmostEqual(r["data"]["volume_m3"], 0.001)
        self.assertEqual(r["data"]["center_of_mass"], [10.0, 20.0, 30.0])

    def test_export_stl(self):
        doc = RichFakeDoc()
        doc.Extension.SaveAs = MagicMock(return_value=True)
        auto = make_automation(doc)
        r = auto.export_document("/tmp/test_part.stl")
        self.assertTrue(r["success"], r)
        self.assertEqual(r["data"]["format"], "STL")
        args = doc.Extension.SaveAs.call_args[0]
        self.assertTrue(args[0].endswith(".stl"))

    def test_export_bad_format(self):
        auto = make_automation(RichFakeDoc())
        r = auto.export_document("/tmp/test_part.xyz")
        self.assertFalse(r["success"])

    def test_reference_plane_offset(self):
        doc = RichFakeDoc()
        plane = types.SimpleNamespace(Name="Plane1")
        doc.FeatureManager.InsertRefPlane = MagicMock(return_value=plane)
        auto = make_automation(doc)
        r = auto.create_reference_plane(5, "Front Plane", unit="mm")
        self.assertTrue(r["success"], r)
        args = doc.FeatureManager.InsertRefPlane.call_args[0]
        self.assertEqual(args[0], 8)  # Distance constraint
        self.assertAlmostEqual(args[1], 0.005)


if __name__ == "__main__":
    unittest.main()
