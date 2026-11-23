"""
SolidWorks Sketch Operations
----------------------------
Create sketches and draw 2D geometry.
"""

import array
import math
import logging
import traceback
from typing import Optional, Dict, List, Tuple

import win32com.client
import pythoncom

from ..constants import SwErrors, SwPlanes, SwSlotCreationType, SwSlotLengthType

logger = logging.getLogger(__name__)


class SketchOperations:
    """
    Mixin class for sketch operations
    
    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _units: UnitConverter instance
    """
    
    def create_sketch(self, plane: str = "Front") -> Dict:
        """
        Create a new sketch on specified plane
        
        Args:
            plane: Plane name - "Front", "Top", or "Right"
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Get plane name
            plane_name = SwPlanes.get(plane)
            
            # Create empty variant for callout parameter
            empty_callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            
            # Select plane
            result = doc.Extension.SelectByID2(
                plane_name, "PLANE", 0, 0, 0, False, 0, empty_callout, 0
            )
            
            if not result:
                return self._result(False, f"Could not select {plane_name}",
                                  SwErrors.swSelectionError)
            
            # Insert sketch
            doc.InsertSketch2(True)
            
            return self._result(True, f"Sketch created on {plane_name}",
                              SwErrors.swSuccess, {"plane": plane_name})
            
        except Exception as e:
            logger.error(f"Create sketch error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def create_sketch_on_face(self, x: float = 0, y: float = 0, z: float = 0,
                              unit: str = None) -> Dict:
        """
        Create a new sketch on a face selected by coordinate.
        ADDED v4.1: Enables cut-extrude on existing body faces (not just ref planes).
        
        Args:
            x, y, z: Coordinates (in user units) of a point ON the target face
            unit: Unit for coordinates (uses default if None)
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Convert to meters
            x_m = self._units.to_meters(x, unit)
            y_m = self._units.to_meters(y, unit)
            z_m = self._units.to_meters(z, unit)
            unit_str = unit or self._units.default_unit.value
            
            # Clear any selection
            doc.ClearSelection2(True)
            
            # Select face at the given coordinates
            empty_callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            selected = doc.Extension.SelectByID2(
                "", "FACE", x_m, y_m, z_m, False, 0, empty_callout, 0
            )
            
            if not selected:
                return self._result(False,
                    f"No face found at ({x}{unit_str}, {y}{unit_str}, {z}{unit_str}). "
                    f"Coordinates must be on an existing body face.",
                    SwErrors.swSelectionError)
            
            # Insert sketch on the selected face
            doc.InsertSketch2(True)
            
            return self._result(True,
                f"Sketch created on face at ({x}, {y}, {z}) {unit_str}",
                SwErrors.swSuccess,
                {"face_point": {"x": x, "y": y, "z": z}, "unit": unit_str})
            
        except Exception as e:
            logger.error(f"Create sketch on face error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def exit_sketch(self) -> Dict:
        """
        Exit the current sketch
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            doc.InsertSketch2(True)
            
            return self._result(True, "Exited sketch", SwErrors.swSuccess)
            
        except Exception as e:
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_line(self, x1: float = 0, y1: float = 0, 
                  x2: float = 100, y2: float = 0, unit: str = None) -> Dict:
        """
        Draw a line in the active sketch
        
        Args:
            x1, y1: Start point coordinates
            x2, y2: End point coordinates
            unit: Unit for coordinates (uses default if None)
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Convert to meters
            x1_m = self._units.to_meters(x1, unit)
            y1_m = self._units.to_meters(y1, unit)
            x2_m = self._units.to_meters(x2, unit)
            y2_m = self._units.to_meters(y2, unit)
            
            # Draw line
            line = doc.SketchManager.CreateLine(x1_m, y1_m, 0, x2_m, y2_m, 0)
            
            if line is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            # Calculate length
            length_m = math.sqrt((x2_m - x1_m)**2 + (y2_m - y1_m)**2)
            length_display = self._units.from_meters(length_m, unit)
            unit_str = unit or self._units.default_unit.value
            
            return self._result(True, f"Line: {length_display:.2f}{unit_str}",
                              SwErrors.swSuccess,
                              {"length": length_display, "unit": unit_str})
            
        except Exception as e:
            logger.error(f"Draw line error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_circle(self, x: float = 0, y: float = 0, 
                    radius: float = 25, unit: str = None) -> Dict:
        """
        Draw a circle in the active sketch
        
        Args:
            x, y: Center point coordinates
            radius: Circle radius
            unit: Unit for dimensions (uses default if None)
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Convert to meters
            x_m = self._units.to_meters(x, unit)
            y_m = self._units.to_meters(y, unit)
            radius_m = self._units.to_meters(radius, unit)
            
            # Draw circle (center point and edge point)
            circle = doc.SketchManager.CreateCircle(
                x_m, y_m, 0,              # Center
                x_m + radius_m, y_m, 0    # Edge point
            )
            
            if circle is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            unit_str = unit or self._units.default_unit.value
            
            return self._result(True, f"Circle: r={radius}{unit_str}",
                              SwErrors.swSuccess,
                              {"radius": radius, "unit": unit_str,
                               "center_x": x, "center_y": y})
            
        except Exception as e:
            logger.error(f"Draw circle error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_rectangle(self, x1: float = -50, y1: float = -25,
                       x2: float = 50, y2: float = 25, unit: str = None) -> Dict:
        """
        Draw a rectangle in the active sketch
        
        Args:
            x1, y1: First corner coordinates
            x2, y2: Opposite corner coordinates
            unit: Unit for dimensions
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Convert to meters
            x1_m = self._units.to_meters(x1, unit)
            y1_m = self._units.to_meters(y1, unit)
            x2_m = self._units.to_meters(x2, unit)
            y2_m = self._units.to_meters(y2, unit)
            
            # Draw rectangle
            rect = doc.SketchManager.CreateCornerRectangle(
                x1_m, y1_m, 0, x2_m, y2_m, 0
            )
            
            if rect is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            width = abs(x2 - x1)
            height = abs(y2 - y1)
            unit_str = unit or self._units.default_unit.value
            
            return self._result(True, f"Rectangle: {width}x{height}{unit_str}",
                              SwErrors.swSuccess,
                              {"width": width, "height": height, "unit": unit_str})
            
        except Exception as e:
            logger.error(f"Draw rectangle error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_arc_center(self, cx: float = 0, cy: float = 0, radius: float = 25,
                        start_angle: float = 0, end_angle: float = 90,
                        unit: str = None) -> Dict:
        """
        Draw an arc by center point and angles
        
        Args:
            cx, cy: Center point
            radius: Arc radius
            start_angle: Start angle in degrees
            end_angle: End angle in degrees
            unit: Unit for radius
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Convert to meters
            cx_m = self._units.to_meters(cx, unit)
            cy_m = self._units.to_meters(cy, unit)
            radius_m = self._units.to_meters(radius, unit)
            
            # Convert angles to radians
            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)
            
            # Calculate start and end points
            x1 = cx_m + radius_m * math.cos(start_rad)
            y1 = cy_m + radius_m * math.sin(start_rad)
            x2 = cx_m + radius_m * math.cos(end_rad)
            y2 = cy_m + radius_m * math.sin(end_rad)
            
            # Draw arc (center, start, end, direction)
            arc = doc.SketchManager.CreateArc(
                cx_m, cy_m, 0,   # Center
                x1, y1, 0,       # Start
                x2, y2, 0,       # End
                1                # Direction (1 = counter-clockwise)
            )
            
            if arc is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            unit_str = unit or self._units.default_unit.value
            arc_angle = abs(end_angle - start_angle)
            
            return self._result(True, f"Arc: r={radius}{unit_str}, {arc_angle}°",
                              SwErrors.swSuccess,
                              {"radius": radius, "start_angle": start_angle,
                               "end_angle": end_angle, "unit": unit_str})
            
        except Exception as e:
            logger.error(f"Draw arc error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_polygon(self, cx: float = 0, cy: float = 0, radius: float = 25,
                     sides: int = 6, unit: str = None) -> Dict:
        """
        Draw a regular polygon
        
        Args:
            cx, cy: Center point
            radius: Circumscribed radius
            sides: Number of sides (3-100)
            unit: Unit for radius
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            if sides < 3 or sides > 100:
                return self._result(False, "Sides must be between 3 and 100",
                                  SwErrors.swInvalidInput)
            
            # Convert to meters
            cx_m = self._units.to_meters(cx, unit)
            cy_m = self._units.to_meters(cy, unit)
            radius_m = self._units.to_meters(radius, unit)
            
            # Draw polygon
            polygon = doc.SketchManager.CreatePolygon(
                cx_m, cy_m, 0,               # Center
                cx_m + radius_m, cy_m, 0,    # Vertex
                sides,                        # Number of sides
                False                         # Inscribed (False = circumscribed)
            )
            
            if polygon is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            unit_str = unit or self._units.default_unit.value
            
            return self._result(True, f"{sides}-sided polygon: r={radius}{unit_str}",
                              SwErrors.swSuccess,
                              {"sides": sides, "radius": radius, "unit": unit_str})
            
        except Exception as e:
            logger.error(f"Draw polygon error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_centerline(self, x1: float = 0, y1: float = -50,
                        x2: float = 0, y2: float = 50, unit: str = None) -> Dict:
        """
        Draw a centerline (for revolve axis)
        
        Args:
            x1, y1: Start point
            x2, y2: End point
            unit: Unit for coordinates
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            # Convert to meters
            x1_m = self._units.to_meters(x1, unit)
            y1_m = self._units.to_meters(y1, unit)
            x2_m = self._units.to_meters(x2, unit)
            y2_m = self._units.to_meters(y2, unit)
            
            # Draw centerline
            line = doc.SketchManager.CreateCenterLine(x1_m, y1_m, 0, x2_m, y2_m, 0)
            
            if line is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            return self._result(True, "Centerline created",
                              SwErrors.swSuccess)
            
        except Exception as e:
            logger.error(f"Draw centerline error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_arc_3point(self, x1: float = 0, y1: float = 0,
                        x2: float = 50, y2: float = 0,
                        x3: float = 25, y3: float = 25,
                        unit: str = None) -> Dict:
        """
        Draw an arc through 3 points.
        API: ISketchManager.Create3PointArc (current in SolidWorks 2027).
        Point 1 = start, point 2 = end, point 3 = a point on the arc.
        
        Args:
            x1, y1: Start point
            x2, y2: End point
            x3, y3: Point on the arc (defines the bulge)
            unit: Unit for coordinates
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            p1 = (self._units.to_meters(x1, unit), self._units.to_meters(y1, unit))
            p2 = (self._units.to_meters(x2, unit), self._units.to_meters(y2, unit))
            p3 = (self._units.to_meters(x3, unit), self._units.to_meters(y3, unit))
            
            arc = doc.SketchManager.Create3PointArc(
                p1[0], p1[1], 0,
                p2[0], p2[1], 0,
                p3[0], p3[1], 0
            )
            
            if arc is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            unit_str = unit or self._units.default_unit.value
            
            return self._result(True,
                f"3-point arc: ({x1},{y1}) -> ({x2},{y2}) via ({x3},{y3}) {unit_str}",
                SwErrors.swSuccess)
            
        except Exception as e:
            logger.error(f"Draw 3-point arc error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_spline(self, points: list, unit: str = None) -> Dict:
        """
        Draw a spline through a list of (x, y) points.
        API: ISketchManager.CreateSpline3 (current in SolidWorks 2027;
        CreateSpline2 is obsolete). For a 2D spline, Surfs and Direction
        are passed as null; SimulateNaturalEnds=True.
        
        Args:
            points: List of (x, y) tuples, at least 2 points
            unit: Unit for coordinates
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            if not points or len(points) < 2:
                return self._result(False, "Spline needs at least 2 points",
                                  SwErrors.swInvalidInput)
            
            point_array = array.array('d')
            for x, y in points:
                point_array.append(self._units.to_meters(x, unit))
                point_array.append(self._units.to_meters(y, unit))
                point_array.append(0)  # Z coordinate
            
            # Null dispatch for Surfs / Direction (2D spline, no surface)
            null_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            
            result = doc.SketchManager.CreateSpline3(
                point_array, null_dispatch, null_dispatch, True
            )
            
            # CreateSpline3 has an out Status parameter; win32com returns
            # (return_value, status) as a tuple in that case.
            spline = result[0] if isinstance(result, tuple) else result
            
            if spline is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            return self._result(True, f"Spline with {len(points)} points created",
                              SwErrors.swSuccess,
                              {"point_count": len(points)})
            
        except Exception as e:
            logger.error(f"Draw spline error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def draw_slot(self, slot_type: str = "straight", width: float = 10,
                  x1: float = 0, y1: float = 0,
                  x2: float = 50, y2: float = 0,
                  x3: float = 0, y3: float = 0,
                  length_type: str = "center_center",
                  arc_direction: int = 1, add_dimension: bool = False,
                  unit: str = None) -> Dict:
        """
        Draw a slot in the active sketch.
        API: ISketchManager.CreateSketchSlot (current in SolidWorks 2027;
        there is no CreateSlot method).
        
        Args:
            slot_type: "straight", "centerline" (centerpoint straight),
                "arc" (centerpoint arc), "3pointarc" (3-point arc)
            width: Slot width (diameter of the end arcs)
            x1, y1: First point (slot start / arc center)
            x2, y2: Second point (slot end)
            x3, y3: Third point (arc slots only; ignored for straight slots)
            length_type: "center_center" (length between arc centers) or
                "full" (overall length)
            arc_direction: -1 = clockwise, 1 = counter-clockwise (arc slots)
            add_dimension: Auto-add dimensions to the slot
            unit: Unit for dimensions
        
        Slot-type integer values are per secondary sources (the official
        enum page lists names only); see constants.SwSlotCreationType.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            type_map = {
                "straight": SwSlotCreationType.STRAIGHT,
                "centerline": SwSlotCreationType.CENTERLINE,
                "arc": SwSlotCreationType.ARC,
                "3pointarc": SwSlotCreationType.THREE_POINT_ARC,
            }
            if slot_type not in type_map:
                return self._result(False,
                    f"Unknown slot_type '{slot_type}'. Use: {list(type_map)}",
                    SwErrors.swInvalidInput)
            
            length_map = {
                "center_center": SwSlotLengthType.CENTER_CENTER,
                "full": SwSlotLengthType.FULL_LENGTH,
            }
            if length_type not in length_map:
                return self._result(False,
                    f"Unknown length_type '{length_type}'. Use: {list(length_map)}",
                    SwErrors.swInvalidInput)
            
            width_m = self._units.to_meters(width, unit)
            p1 = (self._units.to_meters(x1, unit), self._units.to_meters(y1, unit))
            p2 = (self._units.to_meters(x2, unit), self._units.to_meters(y2, unit))
            p3 = (self._units.to_meters(x3, unit), self._units.to_meters(y3, unit))
            
            slot = doc.SketchManager.CreateSketchSlot(
                int(type_map[slot_type]),
                int(length_map[length_type]),
                width_m,
                p1[0], p1[1], 0,
                p2[0], p2[1], 0,
                p3[0], p3[1], 0,
                int(arc_direction),
                bool(add_dimension)
            )
            
            if slot is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)
            
            unit_str = unit or self._units.default_unit.value
            
            return self._result(True,
                f"{slot_type} slot: width {width}{unit_str}",
                SwErrors.swSuccess,
                {"slot_type": slot_type, "width": width, "unit": unit_str})
            
        except Exception as e:
            logger.error(f"Draw slot error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
    
    def trim_entities(self, x: float = 0, y: float = 0, option: int = 0,
                      unit: str = None) -> Dict:
        """
        Trim sketch entities at a pick point.
        API: ISketchManager.SketchTrim (current in SolidWorks 2027;
        IModelDoc2.SketchTrim is obsolete). Trims the selected sketch
        entities near the given point — select the entity first
        (e.g. via execute_python), then trim at the pick location.
        
        Args:
            x, y: Pick point in sketch coordinates
            option: swSketchTrimChoice_e value (0 = default/closest)
            unit: Unit for coordinates
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            x_m = self._units.to_meters(x, unit)
            y_m = self._units.to_meters(y, unit)
            
            trimmed = doc.SketchManager.SketchTrim(int(option), x_m, y_m, 0)
            
            if not trimmed:
                return self._result(False,
                    "Trim failed - select the entity to trim first, and "
                    "pick a point on it",
                    SwErrors.swSketchError)
            
            return self._result(True, f"Trimmed sketch entities at ({x}, {y})",
                              SwErrors.swSuccess)
            
        except Exception as e:
            logger.error(f"Trim entities error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
