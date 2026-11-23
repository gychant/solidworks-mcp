"""
Windows smoke test for the SolidWorks MCP extension (v4.1.0).

Run this ON A WINDOWS MACHINE with SolidWorks 2025 installed:
    python smoke_test_windows.py

It exercises the riskiest new calls and prints a PASS/FAIL report:
  1. _sw_const resolution (swFmLPattern / swFmCirPattern) — the pattern
     tools depend on win32com.client.constants resolving these names
     from the installed type library. If this fails, the pattern tools
     fail with a clear error; everything else still works.
  2. Extrude via FeatureExtrusion3 (2025-current method).
  3. Slot / spline / 3-point arc (slot-type integers are secondary-source).
  4. Reference plane (Distance constraint = 8).
  5. Mass properties (CreateMassProperty2).

No MCP client needed — this talks to SolidWorks directly through the
same automation package the server uses. It creates a throwaway part
and does NOT save it.
"""

import os
import sys
import traceback

# The repo root (this file lives in the repo root).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from solidworks_mcp.automation import SolidWorksAutomation  # noqa: E402


def main():
    results = []

    def check(name, fn):
        try:
            ok, detail = fn()
            results.append((name, bool(ok), detail))
            print(f"[{'PASS' if ok else 'FAIL'}] {name}: {detail}")
        except Exception:
            results.append((name, False, "exception"))
            print(f"[FAIL] {name}: exception")
            traceback.print_exc()

    sw = SolidWorksAutomation()

    def do_connect():
        r = sw.connect()
        return r["success"], r["message"]
    check("connect to SolidWorks", do_connect)

    def do_consts():
        try:
            lp = sw._sw_const("swFmLPattern")
            cp = sw._sw_const("swFmCirPattern")
            return True, f"swFmLPattern={lp}, swFmCirPattern={cp}"
        except Exception as e:
            return False, str(e)
    check("resolve swconst (patterns)", do_consts)

    def do_part():
        r = sw.create_new_part()
        return r["success"], r["message"]
    check("create new part", do_part)

    def do_sketch_circle_extrude():
        r = sw.create_sketch("Front")
        if not r["success"]:
            return False, r["message"]
        r = sw.draw_circle(0, 0, 25)
        if not r["success"]:
            return False, r["message"]
        r = sw.extrude_sketch(10)
        method = (r.get("data") or {}).get("api_method", "?")
        return r["success"], f"{r['message']} [{method}]"
    check("sketch circle + extrude (Extrusion3)", do_sketch_circle_extrude)

    def do_slot_spline_arc():
        r = sw.create_sketch("Front")
        if not r["success"]:
            return False, r["message"]
        r = sw.draw_slot("straight", 10, -40, 0, 40, 0)
        if not r["success"]:
            return False, f"slot: {r['message']}"
        r = sw.draw_spline([(0, 0), (10, 10), (20, 0)])
        if not r["success"]:
            return False, f"spline: {r['message']}"
        r = sw.draw_arc_3point(0, 0, 50, 0, 25, 25)
        return r["success"], r["message"]
    check("slot + spline + 3-point arc", do_slot_spline_arc)

    def do_ref_plane():
        r = sw.create_reference_plane(5, "Front Plane")
        return r["success"], r["message"]
    check("reference plane offset 5mm", do_ref_plane)

    def do_mass_props():
        r = sw.get_mass_properties()
        if not r["success"]:
            return False, r["message"]
        return True, f"mass={r['data'].get('mass_kg')} kg"
    check("mass properties", do_mass_props)

    passed = sum(1 for _, ok, _ in results if ok)
    print(f"\n{passed}/{len(results)} checks passed")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
