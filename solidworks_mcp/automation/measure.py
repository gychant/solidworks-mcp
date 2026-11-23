"""
SolidWorks Measurement Operations
---------------------------------
Distance measurement and mass properties.

Version: 4.1.0 (SolidWorks 2027 API groundwork)
Author: Samsaam Ali Baig

API grounding (verified against the official SolidWorks 2027 API help):
- IModelDocExtension.CreateMeasure() -> IMeasure -> Calculate(entities)
  -> Distance / CenterDistance / Angle / Radius ... (current in 2027).
  Calculate(None) measures the current selection. A value of -1 means
  "not applicable" for the measured combination.
- IModelDocExtension.CreateMassProperty2() -> IMassProperty2 (current in
  2027; CreateMassProperty and GetMassProperties/GetMassProperties2 are
  obsolete). Pre-select bodies/components first, or set SelectedItems.
"""

import logging
import traceback
from typing import Dict

import win32com.client
import pythoncom

from ..constants import SwErrors

logger = logging.getLogger(__name__)


class MeasureOperations:
    """
    Mixin class for measurement operations.

    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _units: UnitConverter instance
    """

    def _select_entity(self, doc, name: str, append: bool = False) -> bool:
        """Select a face or edge by name, with the null-callout fallback."""
        empty_callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        for sel_type in ("FACE", "EDGE"):
            try:
                if doc.Extension.SelectByID2(
                        name, sel_type, 0, 0, 0, append, 0, empty_callout, 0):
                    return True
            except Exception:
                pass
            try:
                if doc.Extension.SelectByID2(
                        name, sel_type, 0, 0, 0, append, 0,
                        pythoncom.Nothing, 0):
                    return True
            except Exception:
                pass
        return False

    def measure_distance(self, entity1: str = "", entity2: str = "",
                         unit: str = None) -> Dict:
        """
        Measure the distance between two entities, or between the
        currently selected entities.
        API: IModelDocExtension.CreateMeasure + IMeasure.Calculate
        (current in SolidWorks 2027).

        Args:
            entity1: Name/type of the first entity to measure, e.g.
                "Face1@Boss-Extrude1" — or "" to use the current
                selection (select entities first, e.g. via execute_python)
            entity2: Second entity; "" for a single selected entity or
                the current selection
            unit: Unit for the returned distance

        Returns:
            Result dictionary with distance in the requested unit.
            distance = -1 means the measurement does not apply to the
            selected combination.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            unit_str = unit or self._units.default_unit.value

            # Select the entities if names were given
            if entity1:
                doc.ClearSelection2(True)
                if not self._select_entity(doc, entity1, append=False):
                    return self._result(False,
                        f"Could not select entity '{entity1}'",
                        SwErrors.swSelectionError)
                if entity2:
                    if not self._select_entity(doc, entity2, append=True):
                        return self._result(False,
                            f"Could not select entity '{entity2}'",
                            SwErrors.swSelectionError)

            measure = doc.Extension.CreateMeasure()
            if measure is None:
                return self._result(False, "CreateMeasure failed",
                                  SwErrors.swMeasureError)

            # Calculate(None) measures the current selection
            if not measure.Calculate(None):
                return self._result(False,
                    "Measure failed - select two valid entities first",
                    SwErrors.swMeasureError)

            distance_m = measure.Distance
            if distance_m is None or distance_m < 0:
                return self._result(False,
                    "Distance not applicable to the selected entities",
                    SwErrors.swMeasureError)

            distance = self._units.from_meters(distance_m, unit)

            return self._result(True,
                f"Distance: {distance:.4f} {unit_str}",
                SwErrors.swSuccess,
                {"distance": round(distance, 6), "unit": unit_str,
                 "api_method": "CreateMeasure"})

        except Exception as e:
            logger.error(f"Measure distance error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swMeasureError)

    def get_mass_properties(self, unit: str = None) -> Dict:
        """
        Get mass properties of the active document (volume, mass,
        surface area, center of mass, principal moments of inertia).
        API: IModelDocExtension.CreateMassProperty2 -> IMassProperty2
        (current in SolidWorks 2027; CreateMassProperty is obsolete).

        For a part, the solid bodies are measured. For an assembly,
        pre-select the components to measure first.

        Args:
            unit: Unit for lengths in the returned values

        Returns:
            Result dictionary with mass properties. Volume is in
            cubic meters; mass in kg; center of mass in the requested unit.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            unit_str = unit or self._units.default_unit.value

            mass_prop = doc.Extension.CreateMassProperty2()
            if mass_prop is None:
                return self._result(False,
                    "CreateMassProperty2 failed - the part may be "
                    "surface-only with an overridden mass",
                    SwErrors.swMeasureError)

            volume_m3 = mass_prop.Volume
            mass_kg = mass_prop.Mass
            surface_area_m2 = mass_prop.SurfaceArea
            com = mass_prop.CenterOfMass  # (x, y, z) in meters
            moments = mass_prop.PrincipalMomentsOfInertia  # (Ix, Iy, Iz)

            com_converted = None
            if com is not None:
                com_converted = [
                    round(self._units.from_meters(com[0], unit), 6),
                    round(self._units.from_meters(com[1], unit), 6),
                    round(self._units.from_meters(com[2], unit), 6),
                ]

            moments_list = None
            if moments is not None:
                moments_list = [round(float(m), 6) for m in moments]

            return self._result(True,
                f"Mass: {mass_kg:.4f} kg, Volume: {volume_m3:.6f} m³",
                SwErrors.swSuccess,
                {
                    "mass_kg": round(float(mass_kg), 6) if mass_kg is not None else None,
                    "volume_m3": round(float(volume_m3), 9) if volume_m3 is not None else None,
                    "surface_area_m2": round(float(surface_area_m2), 9) if surface_area_m2 is not None else None,
                    "center_of_mass": com_converted,
                    "center_of_mass_unit": unit_str,
                    "principal_moments_of_inertia": moments_list,
                    "density": mass_prop.Density,
                    "api_method": "CreateMassProperty2",
                })

        except Exception as e:
            logger.error(f"Mass properties error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swMeasureError)
