"""
SolidWorks Advanced Feature Operations
--------------------------------------
Revolve, loft, Hole Wizard, patterns, shell and mirror.

Version: 4.1.0 (SolidWorks 2025 API groundwork)
Author: Samsaam Ali Baig

API grounding (verified against the official SolidWorks 2025 API help):
- IFeatureManager.FeatureRevolve2 — current in 2025 (older FeatureRevolve*
  variants obsolete). Revolve axis must be pre-selected with mark 16.
- IFeatureManager.InsertProtrusionBlend2 / InsertCutBlend — current loft
  methods in 2025 (InsertProtrusionBlend is obsolete). Profiles must be
  pre-selected.
- IFeatureManager.HoleWizard5 — current in 2025 (HoleWizard..4 obsolete).
- IFeatureManager.InsertMirrorFeature2 — current in 2025 (there is no
  FeatureMirror2). Pre-select the entities and the mirror plane.
- IModelDoc2.InsertFeatureShell(Thickness, Outward) — current in 2025.
  Faces to remove are pre-selected with mark 1.
- Patterns: FeatureLinearPattern5 / FeatureCircularPattern5 are obsolete
  per the 2027 methods list (2025 timing unverified), so the current
  route is IFeatureManager.CreateDefinition(swFeatureNameID_e) ->
  feature-data object -> IFeatureManager.CreateFeature. Feature-data
  integer IDs are resolved at runtime from the installed type library
  (see _sw_const in base.py) — never hard-coded, because the web API
  docs list enum member names but not their integer values.
- Selection marks for circular patterns are per the API docs: 1 =
  direction axis, 4 = features to pattern, 256 = bodies to pattern.
  Linear-pattern marks were not surfaced; the same convention is used.

NOT in this file (unverified against the 2025 docs, do not ship):
- Sweep: the current method is IFeatureManager.InsertProtrusionSwept3,
  but its exact 20-parameter signature did not surface from the docs —
  verify it against the installed type library before implementing.
- InsertRefAxis: no official page surfaced; do not ship.
"""

import logging
import math
import traceback
from typing import Dict, List, Optional

import win32com.client
import pythoncom

from ..constants import SwEndConditions, SwErrors, SwFeatureNameID

logger = logging.getLogger(__name__)


class AdvancedFeatureOperations:
    """
    Mixin class for advanced feature operations.

    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _units: UnitConverter instance
    - _sw_const(name): Runtime SolidWorks constant resolution
    - _find_last_sketch(doc): Last sketch name in the feature tree
    """

    # ========================================================================
    # Helpers
    # ========================================================================

    def _select_by_id(self, doc, name: str, sel_type: str, mark: int = 0,
                      append: bool = False) -> bool:
        """Select an entity by ID, with the repo's null-callout fallback."""
        empty_callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        try:
            if doc.Extension.SelectByID2(
                    name, sel_type, 0, 0, 0, append, mark, empty_callout, 0):
                return True
        except Exception:
            pass
        try:
            return bool(doc.Extension.SelectByID2(
                name, sel_type, 0, 0, 0, append, mark, pythoncom.Nothing, 0))
        except Exception as e:
            logger.debug(f"SelectByID2 failed for {name} ({sel_type}): {e}")
            return False

    def _last_sketches(self, doc, count: int = 2) -> List[str]:
        """Return the names of the last `count` sketches in the tree."""
        names = []
        try:
            feat = doc.FirstFeature
            while feat is not None:
                try:
                    if feat.GetTypeName2 == "ProfileFeature":
                        names.append(feat.Name)
                except Exception:
                    pass
                try:
                    feat = feat.GetNextFeature
                except Exception:
                    break
        except Exception as e:
            logger.debug(f"_last_sketches error: {e}")
        return names[-count:] if len(names) >= count else names

    def _find_centerline_segment(self, doc):
        """
        Find the construction-geometry (centerline) segment of the last
        sketch, for use as a revolve axis.

        Returns:
            (segment, sketch_name) or (None, sketch_name).
        """
        sketch_name = self._find_last_sketch(doc)
        if not sketch_name:
            return None, ""
        try:
            feat = doc.FirstFeature
            while feat is not None:
                try:
                    if feat.Name == sketch_name:
                        sketch = feat.GetSpecificFeature2()
                        for seg in sketch.GetSketchSegments():
                            try:
                                if seg.ConstructionGeometry:
                                    return seg, sketch_name
                            except Exception:
                                continue
                        return None, sketch_name
                except Exception:
                    pass
                try:
                    feat = feat.GetNextFeature
                except Exception:
                    break
        except Exception as e:
            logger.debug(f"_find_centerline_segment error: {e}")
        return None, sketch_name

    # ========================================================================
    # Revolve
    # ========================================================================

    def revolve_sketch(self, angle: float = 360,
                       end_condition: str = "blind",
                       reverse: bool = False) -> Dict:
        """
        Revolve the last sketch around its centerline axis.
        API: IFeatureManager.FeatureRevolve2 (current in SolidWorks 2027).

        The sketch must contain a centerline (draw one with
        draw_centerline first); it is selected as the axis with mark 16.

        Args:
            angle: Revolution angle in degrees (default 360 = full revolve)
            end_condition: "blind" (angle-limited) or "through_all"
            reverse: Reverse the revolution direction

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            end_map = {
                "blind": SwEndConditions.swEndCondBlind,
                "through_all": SwEndConditions.swEndCondThroughAll,
            }
            if end_condition not in end_map:
                return self._result(False,
                    f"Unknown end_condition '{end_condition}'. "
                    f"Use: {list(end_map)}",
                    SwErrors.swInvalidInput)

            # Step 1: close any open sketch and select the profile sketch
            success, sketch_name, error_msg = self._close_and_select_sketch(doc)
            if not success:
                return self._result(False,
                    f"Revolve failed: {error_msg}",
                    SwErrors.swFeatureError)

            # Step 2: select the centerline as the axis (mark 16)
            axis_seg, _ = self._find_centerline_segment(doc)
            if axis_seg is None:
                return self._result(False,
                    f"Revolve failed: no centerline found in sketch "
                    f"'{sketch_name}'. Draw one with draw_centerline first.",
                    SwErrors.swFeatureError)
            try:
                axis_seg.Select2(True, 16)
            except Exception as e:
                return self._result(False,
                    f"Revolve failed: could not select axis: {e}",
                    SwErrors.swSelectionError)

            angle_rad = math.radians(angle)
            end_cond = int(end_map[end_condition])

            feat = doc.FeatureManager.FeatureRevolve2(
                True,            # Sd - single direction
                True,            # IsSolid
                False,           # IsThin
                False,           # IsCut
                bool(reverse),   # ReverseDir
                False,           # BothDirectionUpToSameEntity
                end_cond,        # Dir1Type
                0,               # Dir2Type
                angle_rad,       # Dir1Angle
                0,               # Dir2Angle
                False,           # OffsetReverse1
                False,           # OffsetReverse2
                0,               # OffsetDistance1
                0,               # OffsetDistance2
                0,               # ThinType
                0,               # ThinThickness1
                0,               # ThinThickness2
                True,            # Merge
                True,            # UseFeatScope
                True             # UseAutoSelect
            )

            if feat is None:
                return self._result(False,
                    "Revolve failed - ensure the sketch is a closed profile "
                    "with a centerline axis",
                    SwErrors.swFeatureError)

            return self._result(True,
                f"Revolved sketch '{sketch_name}' {angle}° [{end_condition}]",
                SwErrors.swSuccess,
                {"angle": angle, "end_condition": end_condition,
                 "sketch_name": sketch_name, "api_method": "FeatureRevolve2"})

        except Exception as e:
            logger.error(f"Revolve error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Loft
    # ========================================================================

    def loft_boss(self, closed: bool = False) -> Dict:
        """
        Loft a boss between the last two sketches (profiles).
        API: IFeatureManager.InsertProtrusionBlend2 (current in 2027;
        InsertProtrusionBlend is obsolete). The two profiles are
        pre-selected from the feature tree.

        Args:
            closed: Close the loft (periodic, e.g. for a closed loop)

        Returns:
            Result dictionary
        """
        return self._loft(cut=False, closed=closed)

    def loft_cut(self, closed: bool = False) -> Dict:
        """
        Loft a cut through the last two sketches (profiles).
        API: IFeatureManager.InsertCutBlend (current in 2027).

        Args:
            closed: Close the loft (periodic)

        Returns:
            Result dictionary
        """
        return self._loft(cut=True, closed=closed)

    def _loft(self, cut: bool, closed: bool) -> Dict:
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Close any open sketch, then select the last two profiles
            try:
                if doc.SketchManager.ActiveSketch is not None:
                    doc.SketchManager.InsertSketch(True)
            except Exception:
                pass
            doc.ClearSelection2(True)

            profiles = self._last_sketches(doc, 2)
            if len(profiles) < 2:
                return self._result(False,
                    f"Loft needs 2 profile sketches, found {len(profiles)}: "
                    f"{profiles}",
                    SwErrors.swFeatureError)

            if not self._select_by_id(doc, profiles[0], "SKETCH", 0, False):
                return self._result(False,
                    f"Could not select profile '{profiles[0]}'",
                    SwErrors.swSelectionError)
            if not self._select_by_id(doc, profiles[1], "SKETCH", 0, True):
                return self._result(False,
                    f"Could not select profile '{profiles[1]}'",
                    SwErrors.swSelectionError)

            if cut:
                feat = doc.FeatureManager.InsertCutBlend(
                    bool(closed),  # Closed
                    False,          # KeepTangency
                    False,          # ForceNonRational
                    1.0,            # TessToleranceFactor (default)
                    0,              # StartMatchingType
                    0,              # EndMatchingType
                    False,          # IsThinBody
                    0,              # Thickness1
                    0,              # Thickness2
                    0,              # ThinType
                    True,           # UseFeatScope
                    True            # UseAutoSelect
                )
                method = "InsertCutBlend"
            else:
                feat = doc.FeatureManager.InsertProtrusionBlend2(
                    bool(closed),  # Closed
                    False,          # KeepTangency
                    False,          # ForceNonRational
                    1.0,            # TessToleranceFactor (default)
                    0,              # StartMatchingType
                    0,              # EndMatchingType
                    0,              # StartTangentLength
                    0,              # EndTangentLength
                    False,          # StartTangentDir
                    False,          # EndTangentDir
                    False,          # IsThinBody
                    0,              # Thickness1
                    0,              # Thickness2
                    0,              # ThinType
                    True,           # Merge
                    True,           # UseFeatScope
                    True,           # UseAutoSelect
                    0               # GuideCurveInfluence
                )
                method = "InsertProtrusionBlend2"

            if feat is None:
                return self._result(False,
                    f"Loft failed on profiles {profiles}. Ensure both "
                    f"sketches are closed profiles.",
                    SwErrors.swFeatureError)

            kind = "cut" if cut else "boss"
            return self._result(True,
                f"Loft {kind} between '{profiles[0]}' and '{profiles[1]}' "
                f"[{method}]",
                SwErrors.swSuccess,
                {"profiles": profiles, "closed": closed,
                 "api_method": method})

        except Exception as e:
            logger.error(f"Loft error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Hole Wizard
    # ========================================================================

    def hole_wizard(self, hole_type: int, standard_index: int,
                    fastener_type_index: int, size: str,
                    end_condition: str = "through_all",
                    diameter: float = 0, depth: float = 10,
                    thread_class: str = "", unit: str = None) -> Dict:
        """
        Create a Hole Wizard hole at the points of the selected sketch.
        API: IFeatureManager.HoleWizard5 (current in SolidWorks 2027).

        Workflow: create a sketch on the target face, place a point where
        the hole goes, close the sketch, then call this tool — the hole is
        created at each sketch point.

        Args:
            hole_type: GenericHoleType (swWzdGeneralHoleTypes_e) — index
                of the hole type in the Toolbox database
            standard_index: Hole standard index (swWzdHoleStandards_e)
            fastener_type_index: Fastener type index
                (swWzdHoleStandardFastenerTypes_e)
            size: Fastener size string, e.g. "M6", "1/4-20"
            end_condition: "blind" (depth-controlled) or "through_all"
            diameter: Hole diameter (simple holes); 0 = use Toolbox size
            depth: Hole depth for blind holes (ignored for through_all)
            thread_class: Thread class string, e.g. "6H" (tapped holes)
            unit: Unit for diameter/depth

        Note: the three index parameters are Toolbox-database indices —
        they depend on the installed Toolbox version. Read them off the
        Hole Wizard PropertyManager (Type tab order) for your install.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            end_map = {
                "blind": SwEndConditions.swEndCondBlind,
                "through_all": SwEndConditions.swEndCondThroughAll,
            }
            if end_condition not in end_map:
                return self._result(False,
                    f"Unknown end_condition '{end_condition}'. "
                    f"Use: {list(end_map)}",
                    SwErrors.swInvalidInput)

            # Close any open sketch and select the last (position) sketch
            success, sketch_name, error_msg = self._close_and_select_sketch(doc)
            if not success:
                return self._result(False,
                    f"Hole Wizard failed: {error_msg}. Create a sketch on "
                    f"the target face with a point at the hole position first.",
                    SwErrors.swFeatureError)

            diameter_m = self._units.to_meters(diameter, unit) if diameter else 0
            depth_m = self._units.to_meters(depth, unit)

            feat = doc.FeatureManager.HoleWizard5(
                int(hole_type),            # GenericHoleType
                int(standard_index),       # StandardIndex
                int(fastener_type_index),   # FastenerTypeIndex
                str(size),                  # SSize
                int(end_map[end_condition]),  # EndType
                diameter_m,                 # Diameter
                depth_m,                    # Depth
                0,                          # Length (counterbore/countersink only)
                0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,  # Value1..Value12
                str(thread_class),          # ThreadClass
                False,                      # RevDir
                True,                       # FeatureScope
                True,                       # AutoSelect
                False,                      # AssemblyFeatureScope
                False,                      # AutoSelectComponents
                False                       # PropagateFeatureToParts
            )

            if feat is None:
                return self._result(False,
                    "Hole Wizard failed - check the hole type/standard/"
                    "fastener indices against your Toolbox, and ensure the "
                    "SolidWorks Toolbox add-in is loaded",
                    SwErrors.swFeatureError)

            return self._result(True,
                f"Hole Wizard hole '{size}' ({end_condition}) at sketch "
                f"'{sketch_name}' [HoleWizard5]",
                SwErrors.swSuccess,
                {"size": size, "end_condition": end_condition,
                 "sketch_name": sketch_name, "api_method": "HoleWizard5"})

        except Exception as e:
            logger.error(f"Hole Wizard error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Patterns (CreateDefinition route — 2027 current)
    # ========================================================================

    def _create_pattern(self, feature_name_id: str, data_setup,
                        feature_names: List[str], direction_name: str,
                        direction_type: str, kind: str) -> Dict:
        """
        Shared CreateDefinition pattern flow.

        Selection marks (2027 docs): 1 = direction/axis, 4 = features.
        data_setup(data) sets the numeric properties on the feature-data
        object; returns an error string or "".
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            doc.ClearSelection2(True)

            if not self._select_by_id(doc, direction_name, direction_type,
                                      mark=1, append=False):
                return self._result(False,
                    f"Could not select pattern direction '{direction_name}' "
                    f"({direction_type})",
                    SwErrors.swSelectionError)

            for fname in feature_names:
                if not self._select_by_id(doc, fname, "BODYFEATURE",
                                          mark=4, append=True):
                    return self._result(False,
                        f"Could not select feature '{fname}' to pattern",
                        SwErrors.swSelectionError)

            type_id = self._sw_const(feature_name_id)
            data = doc.FeatureManager.CreateDefinition(type_id)
            if data is None:
                return self._result(False,
                    f"CreateDefinition failed for {feature_name_id} — the "
                    f"constant may not resolve on this SolidWorks version",
                    SwErrors.swFeatureError)

            setup_err = data_setup(data)
            if setup_err:
                return self._result(False, setup_err,
                                  SwErrors.swFeatureError)

            feat = doc.FeatureManager.CreateFeature(data)
            if feat is None:
                return self._result(False,
                    f"{kind} pattern failed - check the direction, feature "
                    f"names and counts",
                    SwErrors.swFeatureError)

            return self._result(True,
                f"{kind} pattern created on {feature_names} [CreateFeature]",
                SwErrors.swSuccess,
                {"features": feature_names, "direction": direction_name,
                 "api_method": "CreateDefinition+CreateFeature"})

        except Exception as e:
            logger.error(f"{kind} pattern error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    def linear_pattern(self, direction_name: str, spacing: float, count: int,
                       feature_names: List[str],
                       direction_type: str = "PLANE",
                       reverse: bool = False, unit: str = None) -> Dict:
        """
        Linear-pattern features along a direction.
        API (2027): CreateDefinition(swFmLPattern) -> ILinearPatternFeatureData
        -> CreateFeature. (FeatureLinearPattern5 is obsolete in 2027.)

        Args:
            direction_name: Name of the direction entity, e.g. "Front Plane"
                or an edge name
            spacing: Distance between instances
            count: Total number of instances (including the seed)
            feature_names: Feature names to pattern, e.g. ["Boss-Extrude1"]
            direction_type: "PLANE", "EDGE" or "AXIS"
            reverse: Reverse the pattern direction
            unit: Unit for spacing

        Returns:
            Result dictionary
        """
        spacing_m = self._units.to_meters(spacing, unit)

        def setup(data):
            try:
                data.Spacing = spacing_m
                data.TotalInstances = int(count)
                data.ReverseDirection = bool(reverse)
            except Exception as e:
                return f"Could not set pattern data: {e}"
            return ""

        return self._create_pattern(
            SwFeatureNameID.LINEAR_PATTERN, setup, feature_names,
            direction_name, direction_type, "Linear")

    def circular_pattern(self, axis_name: str, angle: float, count: int,
                         feature_names: List[str],
                         axis_type: str = "AXIS",
                         equal_spacing: bool = True,
                         reverse: bool = False) -> Dict:
        """
        Circular-pattern features around an axis.
        API (2027): CreateDefinition(swFmCirPattern) ->
        ICircularPatternFeatureData -> CreateFeature.
        (FeatureCircularPattern5 is obsolete in 2027.)

        Args:
            axis_name: Name of the axis entity, e.g. a reference axis
            angle: Total angle in degrees (360 = full circle)
            count: Total number of instances (including the seed)
            feature_names: Feature names to pattern
            axis_type: "AXIS", "EDGE" or "PLANE"
            equal_spacing: Space instances equally over the angle
            reverse: Reverse the pattern direction

        Returns:
            Result dictionary
        """
        angle_rad = math.radians(angle)

        def setup(data):
            try:
                data.TotalAngle = angle_rad
                data.TotalInstances = int(count)
                data.EqualSpacing = bool(equal_spacing)
                data.ReverseDirection = bool(reverse)
            except Exception as e:
                return f"Could not set pattern data: {e}"
            return ""

        return self._create_pattern(
            SwFeatureNameID.CIRCULAR_PATTERN, setup, feature_names,
            axis_name, axis_type, "Circular")

    # ========================================================================
    # Shell (direct method — current in 2025)
    # ========================================================================

    def shell(self, thickness: float, face_points: List[list],
              outward: bool = False, unit: str = None) -> Dict:
        """
        Shell the solid, removing the faces at the given points.
        API: IModelDoc2.InsertFeatureShell(Thickness, Outward) — current
        in SolidWorks 2025 (no FeatureShell2 exists; IShellFeatureData
        is only for the advanced multi-thickness case).

        Faces are selected by 3D coordinates, the same way
        create_sketch_on_face selects a face — each [x, y, z] must lie
        on the face to remove. Faces are selected with mark 1.

        Args:
            thickness: Shell wall thickness
            face_points: List of [x, y, z] points, one per face to remove
            outward: Shell outward instead of inward
            unit: Unit for thickness and coordinates

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            if not face_points:
                return self._result(False,
                    "Shell needs at least one face to remove "
                    "(give [x, y, z] points on the faces)",
                    SwErrors.swInvalidInput)

            thickness_m = self._units.to_meters(thickness, unit)
            unit_str = unit or self._units.default_unit.value

            doc.ClearSelection2(True)
            empty_callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            for pt in face_points:
                x_m = self._units.to_meters(pt[0], unit)
                y_m = self._units.to_meters(pt[1], unit)
                z_m = self._units.to_meters(pt[2], unit)
                selected = doc.Extension.SelectByID2(
                    "", "FACE", x_m, y_m, z_m, True, 1, empty_callout, 0)
                if not selected:
                    return self._result(False,
                        f"No face found at point {pt} {unit_str}",
                        SwErrors.swSelectionError)

            doc.InsertFeatureShell(thickness_m, bool(outward))

            return self._result(True,
                f"Shell: {thickness}{unit_str} wall, {len(face_points)} "
                f"face(s) removed [InsertFeatureShell]",
                SwErrors.swSuccess,
                {"thickness": thickness, "unit": unit_str,
                 "faces_removed": len(face_points),
                 "api_method": "InsertFeatureShell"})

        except Exception as e:
            logger.error(f"Shell error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Mirror (direct method — current in 2025)
    # ========================================================================

    def mirror_feature(self, feature_names: List[str], plane_name: str,
                       mirror_body: bool = False,
                       geometry_pattern: bool = False,
                       merge: bool = True, knit: bool = False) -> Dict:
        """
        Mirror features (or bodies) about a plane.
        API: IFeatureManager.InsertMirrorFeature2 (current in 2025;
        InsertMirrorFeature is obsolete, and there is no FeatureMirror2).

        Args:
            feature_names: Feature names to mirror,
                e.g. ["Boss-Extrude1"]
            plane_name: Mirror plane, e.g. "Front Plane", "Top Plane",
                "Right Plane"
            mirror_body: Mirror the whole body instead of the features
            geometry_pattern: Use a geometry pattern (faster, less robust)
            merge: Merge the mirrored result with the original
            knit: Knit the mirrored surfaces

        Returns:
            Result dictionary

        Note: the selection marks for the pre-selected entities were not
        surfaced from the 2025 docs — features are selected with mark 0
        and the plane appended with mark 0. If the call fails, check the
        InsertMirrorFeature2 Remarks in the API help for the marks.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            if not feature_names and not mirror_body:
                return self._result(False,
                    "Mirror needs at least one feature name, or "
                    "mirror_body=True",
                    SwErrors.swInvalidInput)

            doc.ClearSelection2(True)

            for fname in feature_names:
                if not self._select_by_id(doc, fname, "BODYFEATURE",
                                          mark=0, append=True):
                    return self._result(False,
                        f"Could not select feature '{fname}' to mirror",
                        SwErrors.swSelectionError)

            if not self._select_by_id(doc, plane_name, "PLANE",
                                      mark=0, append=True):
                return self._result(False,
                    f"Could not select mirror plane '{plane_name}'",
                    SwErrors.swSelectionError)

            # ScopeOptions is swFeatureScope_e; 0 = all bodies
            # (secondary source — verify against swconst if it misbehaves).
            feat = doc.FeatureManager.InsertMirrorFeature2(
                bool(mirror_body),       # BMirrorBody
                bool(geometry_pattern),  # BGeometryPattern
                bool(merge),             # BMerge
                bool(knit),              # BKnit
                0                        # ScopeOptions (all bodies)
            )

            if feat is None:
                return self._result(False,
                    "Mirror failed - check the feature names and the "
                    "mirror plane",
                    SwErrors.swFeatureError)

            return self._result(True,
                f"Mirrored {feature_names or 'body'} about '{plane_name}' "
                f"[InsertMirrorFeature2]",
                SwErrors.swSuccess,
                {"features": feature_names, "plane": plane_name,
                 "api_method": "InsertMirrorFeature2"})

        except Exception as e:
            logger.error(f"Mirror error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)
