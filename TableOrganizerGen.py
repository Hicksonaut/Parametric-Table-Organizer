"""
TableOrganizer Grid Generator - Fusion 360 Add-In
==================================================
Generiert Basis-Platte + optionale Walls direkt in der Root-Komponente.
Körper werden automatisch benannt: Basis_CxR, Walls_CxR_NE

Geometrie:
  Basis : 80x80 mm pro Zelle, 10 mm tief (-Z)
          Ecken 8 mm Verrundung, Unterseite 2 mm Fase
          Magnetlöcher: 6.2x1.86 mm bei 20+60 mm, 8 mm tief von unten
  Walls : 7 mm dick, Höhe = Einheiten x 10 mm (+Z)
          Walls starten bei Z=0 (Oberkante der Basis)
          Ecken 8 mm, Oberkante 2 mm Fase, Innen-Boden 4 mm Abrundung
          Shell-Boden wird per Cut bis Z=0 vollständig entfernt

Installation:
  Windows: %appdata%\Autodesk\Autodesk Fusion 360\API\AddIns\TableOrganizer\
  macOS:   ~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/TableOrganizer/
"""

import adsk.core
import adsk.fusion
import adsk.cam
import traceback

_app      = None
_ui       = None
_handlers = []

COMMAND_ID     = "TableOrganizerGridGen"
PANEL_ID       = "SolidCreatePanel"
BUTTON_LABEL   = "Grid generieren"
BUTTON_TOOLTIP = "Table-Organizer Grid mit optionalen Walls generieren"

# Konstanten (Fusion intern = cm)
UNIT         = 8.0    # 80 mm Zellgröße
THICKNESS    = 1.0    # 10 mm Basis-Dicke nach unten

HOLE_LEN     = 0.62   # 6.2 mm  Magnetloch parallel zur Kante
HOLE_DEPTH   = 0.186  # 1.86 mm Magnetloch senkrecht
WALL_MIN     = 0.1    # 1 mm    Wandstärke bis Loch-Anfang
HOLE_POS_A   = 2.0    # 20 mm
HOLE_POS_B   = 6.0    # 60 mm
HOLE_Z_DEPTH = 0.8    # 8 mm    Lochtiefe von unten

WALL_THICK   = 0.7    # 7 mm Wall-Dicke
WALL_UNIT_H  = 1.0    # 10 mm je Einheit
WALL_CHAMFER = 0.2    # 2 mm Fase Oberkante
WALL_FILLET  = 0.4    # 4 mm Abrundung Innen-Boden


# ══════════════════════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
# ══════════════════════════════════════════════════════════════════════════════

def _fillet(comp, edges_col, r):
    if edges_col.count == 0:
        return
    try:
        fi = comp.features.filletFeatures.createInput()
        fi.addConstantRadiusEdgeSet(
            edges_col, adsk.core.ValueInput.createByReal(r), True
        )
        fi.isRollingBallCorner = True
        comp.features.filletFeatures.add(fi)
    except Exception:
        pass


def _chamfer(comp, edges_col, d):
    if edges_col.count == 0:
        return
    ch = comp.features.chamferFeatures
    try:
        ci = ch.createInput2()
        ci.chamferEdgeSets.addEqualDistanceChamferEdgeSet(
            edges_col, adsk.core.ValueInput.createByReal(d), False
        )
        ch.add(ci)
    except Exception:
        try:
            ci = ch.createInput(edges_col, False)
            ci.setToEqualDistance(adsk.core.ValueInput.createByReal(d))
            ch.add(ci)
        except Exception:
            pass


def _vertical_edges(body):
    col = adsk.core.ObjectCollection.create()
    for e in body.edges:
        try:
            sv, ev = e.startVertex.geometry, e.endVertex.geometry
            if abs(ev.x-sv.x) < 0.001 and abs(ev.y-sv.y) < 0.001 and abs(ev.z-sv.z) > 0.001:
                col.add(e)
        except Exception:
            pass
    return col


def _face_edges(body, z_normal_positive):
    """Kanten der Fläche mit Normale in +Z (True) oder -Z (False)."""
    col = adsk.core.ObjectCollection.create()
    for face in body.faces:
        if face.geometry.objectType == adsk.core.Plane.classType():
            ok, n = face.evaluator.getNormalAtPoint(face.pointOnFace)
            if ok and ((z_normal_positive and n.z > 0.99) or (not z_normal_positive and n.z < -0.99)):
                for e in face.edges:
                    col.add(e)
                break
    return col


# ══════════════════════════════════════════════════════════════════════════════
#  BASIS-PLATTE
# ══════════════════════════════════════════════════════════════════════════════

def _create_base(comp, cols, rows):
    sketches   = comp.sketches
    xy_plane   = comp.xYConstructionPlane
    extrudes   = comp.features.extrudeFeatures
    constructs = comp.constructionPlanes

    # Körper nach unten
    sk = sketches.add(xy_plane)
    sk.name = f"Basis_Skizze_{cols}x{rows}"
    sk.sketchCurves.sketchLines.addTwoPointRectangle(
        adsk.core.Point3D.create(0,         0,         0),
        adsk.core.Point3D.create(UNIT*cols, UNIT*rows, 0)
    )
    ei = extrudes.createInput(
        sk.profiles.item(0),
        adsk.fusion.FeatureOperations.NewBodyFeatureOperation
    )
    ei.setOneSideExtent(
        adsk.fusion.DistanceExtentDefinition.create(
            adsk.core.ValueInput.createByReal(THICKNESS)
        ),
        adsk.fusion.ExtentDirections.NegativeExtentDirection
    )
    feat = extrudes.add(ei)
    body = feat.bodies.item(0)
    body.name = f"Basis_{cols}x{rows}"

    # Ecken abrunden (8 mm, vertikal)
    _fillet(comp, _vertical_edges(body), 0.8)

    # Unterseite fasen (2 mm) – VOR Loch-Cut
    _chamfer(comp, _face_edges(body, False), 0.2)

    # Konstruktionsebene Unterseite für Loch-Skizze
    pi = constructs.createInput()
    pi.setByOffset(xy_plane, adsk.core.ValueInput.createByReal(-THICKNESS))
    bp = constructs.add(pi)
    bp.name = f"Unterseite_{cols}x{rows}"
    bp.isLightBulbOn = False

    sk_h = sketches.add(bp)
    sk_h.name = f"Magnetloecher_{cols}x{rows}"
    hl = sk_h.sketchCurves.sketchLines

    def rect(x1, y1, x2, y2):
        hl.addTwoPointRectangle(
            adsk.core.Point3D.create(x1, y1, 0),
            adsk.core.Point3D.create(x2, y2, 0)
        )

    ws, we = WALL_MIN, WALL_MIN + HOLE_DEPTH

    for col in range(cols):                          # Süd
        xo = col * UNIT
        for cx in (xo+HOLE_POS_A, xo+HOLE_POS_B):
            rect(cx-HOLE_LEN/2, ws, cx+HOLE_LEN/2, we)
    for col in range(cols):                          # Nord
        xo, ye = col*UNIT, rows*UNIT
        for cx in (xo+HOLE_POS_A, xo+HOLE_POS_B):
            rect(cx-HOLE_LEN/2, ye-we, cx+HOLE_LEN/2, ye-ws)
    for row in range(rows):                          # West
        yo = row * UNIT
        for cy in (yo+HOLE_POS_A, yo+HOLE_POS_B):
            rect(ws, cy-HOLE_LEN/2, we, cy+HOLE_LEN/2)
    for row in range(rows):                          # Ost
        yo, xe = row*UNIT, cols*UNIT
        for cy in (yo+HOLE_POS_A, yo+HOLE_POS_B):
            rect(xe-we, cy-HOLE_LEN/2, xe-ws, cy+HOLE_LEN/2)

    pc = adsk.core.ObjectCollection.create()
    for i in range(sk_h.profiles.count):
        pc.add(sk_h.profiles.item(i))
    if pc.count > 0:
        ci = extrudes.createInput(pc, adsk.fusion.FeatureOperations.CutFeatureOperation)
        ci.setOneSideExtent(
            adsk.fusion.DistanceExtentDefinition.create(
                adsk.core.ValueInput.createByReal(HOLE_Z_DEPTH)
            ),
            adsk.fusion.ExtentDirections.PositiveExtentDirection
        )
        extrudes.add(ci)


# ══════════════════════════════════════════════════════════════════════════════
#  WALLS
# ══════════════════════════════════════════════════════════════════════════════

def _create_walls(comp, cols, rows, wall_units):
    """
    Walls werden direkt auf Z=0 (Oberkante der Basis) platziert.
    Höhe = wall_units × 10 mm.

    Kein Shell! Stattdessen:
      1. Skizze mit zwei konzentrischen Rechtecken auf XY (Z=0):
           Außen: 0..tw × 0..td
           Innen: wt..tw-wt × wt..td-wt
         → Fusion erkennt das als Ringprofil (wie ein Bilderrahmen)
      2. Ringprofil nach oben extrudieren  → sofort hohler Rahmen, Boden offen
      3. Außen-Ecken R8mm (vertikale Außenecken)
      4. Fase 2mm an äußeren Oberkanten
    """
    sketches = comp.sketches
    xy_plane = comp.xYConstructionPlane
    extrudes = comp.features.extrudeFeatures

    wall_h = wall_units * WALL_UNIT_H   # Höhe = Einheiten × 10mm, ab Z=0
    wt     = WALL_THICK
    tw, td = UNIT * cols, UNIT * rows

    # ── 1. Skizze: Außen- + Innen-Rechteck → Ringprofil ──────────────────────
    sk = sketches.add(xy_plane)
    sk.name = f"Wall_Skizze_{cols}x{rows}"
    lines  = sk.sketchCurves.sketchLines

    # Außenkontur
    lines.addTwoPointRectangle(
        adsk.core.Point3D.create(0,  0,  0),
        adsk.core.Point3D.create(tw, td, 0)
    )
    # Innenkontur (erzeugt das "Loch" im Profil)
    lines.addTwoPointRectangle(
        adsk.core.Point3D.create(wt,      wt,      0),
        adsk.core.Point3D.create(tw - wt, td - wt, 0)
    )

    # Fusion erzeugt jetzt 2 Profile: das äußere Ringprofil (zwischen den Rechtecken)
    # und das innere Profil (Loch). Wir wollen NUR das Ringprofil extrudieren.
    # Das Ringprofil ist das größere der beiden – bei 2 Profilen ist es item(0)
    # wenn das äußere zuerst gezeichnet wurde, aber wir prüfen per Flächeninhalt.
    ring_profile = None
    for i in range(sk.profiles.count):
        p = sk.profiles.item(i)
        props = p.areaProperties()
        # Ringfläche = tw*td - (tw-2wt)*(td-2wt), Innenfläche = (tw-2wt)*(td-2wt)
        # Das Ringprofil hat die größere Fläche wenn Ring > Innen, d.h. tw*td > 2*(tw-2wt)*(td-2wt)
        # Einfacher: Ringprofil hat area ≈ tw*td - (tw-2wt)*(td-2wt)
        inner_area = (tw - 2*wt) * (td - 2*wt)
        outer_area = tw * td
        ring_area  = outer_area - inner_area
        # Toleranz 1%
        if abs(props.area - ring_area) < ring_area * 0.01:
            ring_profile = p
            break

    # Fallback: nimm das Profil mit der größeren Fläche (Ring > Innen)
    if ring_profile is None and sk.profiles.count >= 2:
        areas = [(sk.profiles.item(i).areaProperties().area, i) for i in range(sk.profiles.count)]
        areas.sort(reverse=True)
        ring_profile = sk.profiles.item(areas[0][1])

    if ring_profile is None:
        _ui.messageBox("Wall-Fehler: Kein Ringprofil gefunden.")
        return

    # ── 2. Ringprofil nach oben extrudieren ───────────────────────────────────
    ei = extrudes.createInput(
        ring_profile,
        adsk.fusion.FeatureOperations.NewBodyFeatureOperation
    )
    ei.setOneSideExtent(
        adsk.fusion.DistanceExtentDefinition.create(
            adsk.core.ValueInput.createByReal(wall_h)
        ),
        adsk.fusion.ExtentDirections.PositiveExtentDirection
    )
    feat      = extrudes.add(ei)
    wall_body = feat.bodies.item(0)
    wall_body.name = f"Walls_{cols}x{rows}_{wall_units}E"

    # ── 3. Innen-Ecken R5mm (vertikale Kanten der Innenkontur) ───────────────
    inner_v = adsk.core.ObjectCollection.create()
    for edge in wall_body.edges:
        try:
            sv = edge.startVertex.geometry
            ev = edge.endVertex.geometry
            if abs(ev.x-sv.x) < 0.001 and abs(ev.y-sv.y) < 0.001 and abs(ev.z-sv.z) > 0.001:
                x, y = sv.x, sv.y
                # Innenkante: liegt NICHT auf der Außenkontur
                on_outer = x < 0.01 or x > tw-0.01 or y < 0.01 or y > td-0.01
                if not on_outer:
                    inner_v.add(edge)
        except Exception:
            pass
    _fillet(comp, inner_v, 0.5)  # R5mm

    # ── 4. Außen-Ecken R8mm ───────────────────────────────────────────────────
    outer_v = adsk.core.ObjectCollection.create()
    for edge in wall_body.edges:
        try:
            sv = edge.startVertex.geometry
            ev = edge.endVertex.geometry
            if abs(ev.x-sv.x) < 0.001 and abs(ev.y-sv.y) < 0.001 and abs(ev.z-sv.z) > 0.001:
                x, y = sv.x, sv.y
                on_corner = (
                    (x < 0.01 or x > tw-0.01) and
                    (y < 0.01 or y > td-0.01)
                )
                if on_corner:
                    outer_v.add(edge)
        except Exception:
            pass
    _fillet(comp, outer_v, 0.8)  # R8mm

    # ── 4. Fase 2mm – nur äußere Oberkanten ───────────────────────────────────
    outer_top = adsk.core.ObjectCollection.create()
    for edge in wall_body.edges:
        try:
            sv = edge.startVertex.geometry
            ev = edge.endVertex.geometry
            if abs(sv.z - wall_h) < 0.01 and abs(ev.z - wall_h) < 0.01:
                mx = (sv.x + ev.x) / 2.0
                my = (sv.y + ev.y) / 2.0
                on_outer = mx < wt*0.9 or mx > tw-wt*0.9 or my < wt*0.9 or my > td-wt*0.9
                if on_outer:
                    outer_top.add(edge)
        except Exception:
            pass
    _chamfer(comp, outer_top, WALL_CHAMFER)


def _combine_and_fillet_inner(comp, cols, rows):
    """
    1. Basis + Walls zu einem Körper zusammenführen (Join).
    2. Innere Bodenkante bei Z=0 mit R4mm abrunden (8 Kanten, Tangentenkette).
       Das entspricht genau der manuellen Einstellung: 8 Kanten, 4mm, Rollende Kugel.
    """
    base_body = None
    wall_body = None
    for body in comp.bRepBodies:
        if body.name.startswith(f"Basis_{cols}x{rows}"):
            base_body = body
        elif body.name.startswith(f"Walls_{cols}x{rows}"):
            wall_body = body

    if base_body is None or wall_body is None:
        return

    tool_col = adsk.core.ObjectCollection.create()
    tool_col.add(wall_body)

    combine_in = comp.features.combineFeatures.createInput(base_body, tool_col)
    combine_in.operation        = adsk.fusion.FeatureOperations.JoinFeatureOperation
    combine_in.isKeepToolBodies = False
    comp.features.combineFeatures.add(combine_in)

    # Nach dem Join: innere Bodenkanten bei Z=0 mit R4mm abrunden.
    # Das sind die 4 horizontalen Kanten an der Innenseite der Walls auf Z=0,
    # also wo die Wall-Innenfläche auf die Basis-Oberfläche trifft.
    wt     = WALL_THICK
    tw, td = UNIT * cols, UNIT * rows

    inner_bottom = adsk.core.ObjectCollection.create()
    for body in comp.bRepBodies:
        if not body.name.startswith(f"Basis_{cols}x{rows}"):
            continue
        for edge in body.edges:
            try:
                sv = edge.startVertex.geometry
                ev = edge.endVertex.geometry
                # Horizontale Kante (Z konstant) auf Z=0
                if abs(sv.z) > 0.005 or abs(ev.z) > 0.005:
                    continue
                # Beide Endpunkte müssen innen liegen (nicht auf Außenkontur)
                def is_inner_pt(p):
                    return (p.x > 0.005 and p.x < tw - 0.005 and
                            p.y > 0.005 and p.y < td - 0.005)
                if is_inner_pt(sv) and is_inner_pt(ev):
                    inner_bottom.add(edge)
            except Exception:
                pass

    if inner_bottom.count > 0:
        try:
            fi = comp.features.filletFeatures.createInput()
            fi.addConstantRadiusEdgeSet(
                inner_bottom,
                adsk.core.ValueInput.createByReal(WALL_FILLET),  # 0.4 cm = 4mm
                True   # tangent chain = True → entspricht "Tangentenkette aktiv"
            )
            fi.isRollingBallCorner = True  # Rollende Kugel
            comp.features.filletFeatures.add(fi)
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
#  COMMAND HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

class GridCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        try:
            cmd  = args.command
            inps = cmd.commandInputs
            cmd.isRepeatable = False

            # ── Haupteinstellungen ─────────────────────────────────────────────
            inps.addIntegerSpinnerCommandInput("cols", "Spalten (X)", 1, 20, 1, 1)
            inps.addIntegerSpinnerCommandInput("rows", "Reihen (Y)",  1, 20, 1, 1)

            # ── Walls ──────────────────────────────────────────────────────────
            wg = inps.addGroupCommandInput("wall_group", "Walls")
            wg.isExpanded = True
            gi = wg.children
            gi.addBoolValueInput("gen_walls", "Walls generieren", True, "", False)
            gi.addIntegerSpinnerCommandInput(
                "wall_units", "Wall-Hoehe (Einheiten à 10 mm)", 1, 50, 1, 2
            )

            # ── Experten-Modus (standardmäßig zugeklappt) ─────────────────────
            eg = inps.addGroupCommandInput("expert_group", "⚙  Experten-Modus")
            eg.isExpanded  = False   # versteckt bis der User aufklappt
            ei = eg.children

            ei.addBoolValueInput("expert_on", "Experten-Modus aktivieren", True, "", False)

            # Benutzerdefinierte Wand-Höhe in mm (statt Einheiten)
            # Sichtbarkeit wird per inputChanged gesteuert
            hi = ei.addValueInput(
                "custom_wall_h", "Benutzerdefinierte Wall-Höhe (mm)",
                "mm",
                adsk.core.ValueInput.createByReal(2.0)   # 20 mm default
            )
            hi.isVisible = False

            # Handlers registrieren
            on_change = GridCommandInputChangedHandler()
            cmd.inputChanged.add(on_change)
            _handlers.append(on_change)

            on_exec = GridCommandExecuteHandler()
            cmd.execute.add(on_exec)
            _handlers.append(on_exec)
            on_dest = GridCommandDestroyHandler()
            cmd.destroy.add(on_dest)
            _handlers.append(on_dest)

        except:
            if _ui:
                _ui.messageBox("Fehler CommandCreated:\n" + traceback.format_exc())


class GridCommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args):
        try:
            changed = args.input
            inps    = args.inputs

            expert_on      = inps.itemById("expert_on")
            custom_wall_h  = inps.itemById("custom_wall_h")
            wall_units_inp = inps.itemById("wall_units")

            if expert_on is None or custom_wall_h is None:
                return

            enabled = expert_on.value

            # Benutzerdefinierte Höhe ein-/ausblenden
            custom_wall_h.isVisible  = enabled
            # Standard-Einheiten-Spinner ausblenden wenn custom aktiv
            wall_units_inp.isVisible = not enabled

        except:
            pass


class GridCommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        try:
            inps      = args.command.commandInputs
            cols      = inps.itemById("cols").value
            rows      = inps.itemById("rows").value
            gen_walls = inps.itemById("gen_walls").value

            expert_on     = inps.itemById("expert_on").value
            custom_wall_h = inps.itemById("custom_wall_h")   # ValueInput in cm intern

            if expert_on and custom_wall_h.isVisible:
                # custom_wall_h.value ist in cm (Fusion intern)
                # Umrechnung auf wall_units (je 1.0 cm = 10 mm = 1 Einheit)
                wall_units_real = custom_wall_h.value / WALL_UNIT_H
            else:
                wall_units_real = inps.itemById("wall_units").value

            design = _app.activeProduct
            if not design or not isinstance(design, adsk.fusion.Design):
                _ui.messageBox("Bitte ein Fusion-Design oeffnen!")
                return

            root = design.rootComponent

            _create_base(root, cols, rows)
            if gen_walls:
                _create_walls(root, cols, rows, wall_units_real)
                _combine_and_fillet_inner(root, cols, rows)

            hole_count  = (cols + rows) * 2 * 2
            wall_mm     = round(wall_units_real * 10, 1)
            _ui.messageBox(
                f"\u2705 Grid_{cols}x{rows} generiert!\n"
                f"Groesse: {cols*80}x{rows*80} mm | 10 mm tief\n"
                f"Magnetloecher: {hole_count} Stueck\n"
                + (f"Walls: {wall_mm} mm hoch (ab Oberkante Basis)"
                   if gen_walls else "Walls: nicht generiert")
            )

        except:
            if _ui:
                _ui.messageBox("Fehler beim Generieren:\n" + traceback.format_exc())


class GridCommandDestroyHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  ADD-IN ENTRY POINTS
# ══════════════════════════════════════════════════════════════════════════════

def run(context):
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui  = _app.userInterface

        ex = _ui.commandDefinitions.itemById(COMMAND_ID)
        if ex:
            ex.deleteMe()

        cmd_def = _ui.commandDefinitions.addButtonDefinition(
            COMMAND_ID, BUTTON_LABEL, BUTTON_TOOLTIP, ""
        )
        on_c = GridCommandCreatedHandler()
        cmd_def.commandCreated.add(on_c)
        _handlers.append(on_c)

        panel = _ui.allToolbarPanels.itemById(PANEL_ID)
        if panel:
            ctrl = panel.controls.addCommand(cmd_def)
            ctrl.isPromotedByDefault = True
            ctrl.isPromoted = True

        adsk.autoTerminate(False)

    except:
        if _ui:
            _ui.messageBox("Add-In Start-Fehler:\n" + traceback.format_exc())


def stop(context):
    try:
        panel = _ui.allToolbarPanels.itemById(PANEL_ID)
        if panel:
            ctrl = panel.controls.itemById(COMMAND_ID)
            if ctrl:
                ctrl.deleteMe()
        cd = _ui.commandDefinitions.itemById(COMMAND_ID)
        if cd:
            cd.deleteMe()
    except:
        pass