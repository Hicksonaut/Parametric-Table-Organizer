"""
TableOrganizer Inlay Generator - Fusion 360 Add-In
===================================================
Zweites Plugin der TableOrganizer-Reihe.
Generiert eine solide Inlay-Basis (OHNE Magnetlöcher) + optionale Walls.

Geometrie:
  Basis    : 80x80 mm pro Zelle, 5 mm tief (-Z)
             Ecken 8 mm Verrundung, Unterseite 2 mm Fase
             Keine Magnetlöcher
  Connector: 65x65 mm, 1.5 mm dick, zentriert unter der Basis
             wird mit der Basis gejoined
  Walls    : 7 mm dick, Höhe = Einheiten x 10 mm (+Z)
             Walls starten bei Z=0 (Oberkante der Basis)
             Außen-Ecken 8 mm, Innen-Ecken 5 mm
             Oberkante 2 mm Fase, Innen-Boden 4 mm Abrundung

Installation:
  Windows: %appdata%\Autodesk\Autodesk Fusion 360\API\AddIns\TableOrganizerInlay\
  macOS:   ~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/TableOrganizerInlay/
"""

import adsk.core
import adsk.fusion
import adsk.cam
import traceback

_app      = None
_ui       = None
_handlers = []

COMMAND_ID     = "TableOrganizerInlayGen"
PANEL_ID       = "SolidCreatePanel"
BUTTON_LABEL   = "TableOrganizer Inlay Generator"
BUTTON_TOOLTIP = "Table-Organizer Inlay mit Connector + optionalen Walls generieren"

# Konstanten (Fusion intern = cm)
UNIT         = 8.0    # 80 mm Zellgröße
THICKNESS    = 0.5    # 5 mm Basis-Dicke nach unten

CONN_SIZE    = 6.5    # 65 mm Connector-Seite (quadratisch)
CONN_THICK   = 0.15   # 1.5 mm Connector-Dicke nach unten

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
#  BASIS-PLATTE (solid, keine Magnete, 5 mm dick)
# ══════════════════════════════════════════════════════════════════════════════

def _create_base(comp, cols, rows):
    sketches = comp.sketches
    xy_plane = comp.xYConstructionPlane
    extrudes = comp.features.extrudeFeatures
    constructs = comp.constructionPlanes

    # ── Basis-Platte ──────────────────────────────────────────────────────────
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

    # Unterseite fasen (2 mm)
    _chamfer(comp, _face_edges(body, False), 0.2)

    # ── Connector: 65x65 mm, 1.5 mm dick, zentriert, an Unterseite ───────────
    # Konstruktionsebene auf Unterseite der Basis (Z = -THICKNESS)
    pi = constructs.createInput()
    pi.setByOffset(xy_plane, adsk.core.ValueInput.createByReal(-THICKNESS))
    bottom_plane = constructs.add(pi)
    bottom_plane.name = f"Unterseite_{cols}x{rows}"
    bottom_plane.isLightBulbOn = False

    # Connector zentriert pro Zelle
    sk_c = sketches.add(bottom_plane)
    sk_c.name = f"Connector_Skizze_{cols}x{rows}"
    cl = sk_c.sketchCurves.sketchLines

    for col in range(cols):
        for row in range(rows):
            # Zellmittelpunkt
            cx = col * UNIT + UNIT / 2.0
            cy = row * UNIT + UNIT / 2.0
            half = CONN_SIZE / 2.0
            cl.addTwoPointRectangle(
                adsk.core.Point3D.create(cx - half, cy - half, 0),
                adsk.core.Point3D.create(cx + half, cy + half, 0)
            )

    # Alle Profile (je ein Connector-Quadrat pro Zelle) extrudieren
    pc = adsk.core.ObjectCollection.create()
    for i in range(sk_c.profiles.count):
        pc.add(sk_c.profiles.item(i))

    if pc.count > 0:
        ci = extrudes.createInput(
            pc,
            adsk.fusion.FeatureOperations.JoinFeatureOperation  # direkt in Basis joinen
        )
        ci.setOneSideExtent(
            adsk.fusion.DistanceExtentDefinition.create(
                adsk.core.ValueInput.createByReal(CONN_THICK)
            ),
            adsk.fusion.ExtentDirections.NegativeExtentDirection  # von Unterseite weiter nach unten
        )
        conn_feat = extrudes.add(ci)

        # Ecken des Connectors R2mm abrunden
        # Vertikale Kanten die auf Z = -THICKNESS bis Z = -(THICKNESS+CONN_THICK) liegen
        conn_edges = adsk.core.ObjectCollection.create()
        z_top = -THICKNESS
        z_bot = -(THICKNESS + CONN_THICK)
        for body in comp.bRepBodies:
            if not body.name.startswith(f"Basis_{cols}x{rows}"):
                continue
            for edge in body.edges:
                try:
                    sv = edge.startVertex.geometry
                    ev = edge.endVertex.geometry
                    # Vertikale Kante im Connector-Z-Bereich
                    if abs(ev.x-sv.x) < 0.001 and abs(ev.y-sv.y) < 0.001:
                        z_min = min(sv.z, ev.z)
                        z_max = max(sv.z, ev.z)
                        if abs(z_max - z_top) < 0.005 and abs(z_min - z_bot) < 0.005:
                            conn_edges.add(edge)
                except Exception:
                    pass
        _fillet(comp, conn_edges, 0.2)  # R2mm


# ══════════════════════════════════════════════════════════════════════════════
#  WALLS (1:1 übernommen)
# ══════════════════════════════════════════════════════════════════════════════

def _create_walls(comp, cols, rows, wall_units):
    """
    Walls werden direkt auf Z=0 (Oberkante der Basis) platziert.
    Höhe = wall_units × 10 mm.

    Zwei konzentrische Rechtecke in einer Skizze → Ringprofil → Extrusion.
    Kein Shell nötig, Boden sofort offen.
    """
    sketches = comp.sketches
    xy_plane = comp.xYConstructionPlane
    extrudes = comp.features.extrudeFeatures

    wall_h = wall_units * WALL_UNIT_H
    wt     = WALL_THICK
    tw, td = UNIT * cols, UNIT * rows

    # ── 1. Skizze: Außen- + Innen-Rechteck → Ringprofil ──────────────────────
    sk = sketches.add(xy_plane)
    sk.name = f"Wall_Skizze_{cols}x{rows}"
    lines  = sk.sketchCurves.sketchLines

    lines.addTwoPointRectangle(
        adsk.core.Point3D.create(0,  0,  0),
        adsk.core.Point3D.create(tw, td, 0)
    )
    lines.addTwoPointRectangle(
        adsk.core.Point3D.create(wt,      wt,      0),
        adsk.core.Point3D.create(tw - wt, td - wt, 0)
    )

    # Ringprofil identifizieren (Fläche ≈ tw*td - (tw-2wt)*(td-2wt))
    ring_profile = None
    inner_area   = (tw - 2*wt) * (td - 2*wt)
    ring_area    = tw * td - inner_area
    for i in range(sk.profiles.count):
        p = sk.profiles.item(i)
        if abs(p.areaProperties().area - ring_area) < ring_area * 0.01:
            ring_profile = p
            break

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

    # ── 3. Innen-Ecken R5mm ────────────────────────────────────────────────────
    inner_v = adsk.core.ObjectCollection.create()
    for edge in wall_body.edges:
        try:
            sv = edge.startVertex.geometry
            ev = edge.endVertex.geometry
            if abs(ev.x-sv.x) < 0.001 and abs(ev.y-sv.y) < 0.001 and abs(ev.z-sv.z) > 0.001:
                x, y = sv.x, sv.y
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

    # ── 5. Fase 2mm – nur äußere Oberkanten ───────────────────────────────────
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


# ══════════════════════════════════════════════════════════════════════════════
#  COMBINE + INNEN-BODEN-FILLET
# ══════════════════════════════════════════════════════════════════════════════

def _combine_and_fillet_inner(comp, cols, rows):
    """
    1. Basis + Walls zu einem Körper zusammenführen (Join).
    2. Innere Bodenkante bei Z=0 mit R4mm abrunden (Tangentenkette, Rollende Kugel).
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

    # Innere Bodenkanten bei Z=0 mit R4mm abrunden
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
                if abs(sv.z) > 0.005 or abs(ev.z) > 0.005:
                    continue
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
                adsk.core.ValueInput.createByReal(WALL_FILLET),
                True
            )
            fi.isRollingBallCorner = True
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

            # ── Experten-Modus ─────────────────────────────────────────────────
            eg = inps.addGroupCommandInput("expert_group", "⚙  Experten-Modus")
            eg.isExpanded = False
            ei = eg.children

            ei.addBoolValueInput("expert_on", "Experten-Modus aktivieren", True, "", False)

            hi = ei.addValueInput(
                "custom_wall_h", "Benutzerdefinierte Wall-Höhe (mm)",
                "mm",
                adsk.core.ValueInput.createByReal(2.0)
            )
            hi.isVisible = False

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
            inps           = args.inputs
            expert_on      = inps.itemById("expert_on")
            custom_wall_h  = inps.itemById("custom_wall_h")
            wall_units_inp = inps.itemById("wall_units")

            if expert_on is None or custom_wall_h is None:
                return

            enabled = expert_on.value
            custom_wall_h.isVisible  = enabled
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
            custom_wall_h = inps.itemById("custom_wall_h")

            if expert_on and custom_wall_h.isVisible:
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

            wall_mm = round(wall_units_real * 10, 1)
            _ui.messageBox(
                f"\u2705 Inlay Grid_{cols}x{rows} generiert!\n"
                f"Groesse: {cols*80}x{rows*80} mm | 5 mm tief\n"
                f"Connector: 65x65 mm | 1.5 mm tief | zentriert pro Zelle\n"
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