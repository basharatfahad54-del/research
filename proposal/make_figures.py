# -*- coding: utf-8 -*-
"""Generates the diagrams for the refined proposal (v2, layout fixes)."""
import os
import math
from PIL import Image, ImageDraw, ImageFont

OUT = r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\proposal\assets"
os.makedirs(OUT, exist_ok=True)

NAVY = (31, 56, 100)
BLUE = (46, 117, 182)
LIGHT = (222, 234, 246)
PALE = (242, 245, 250)
GREY = (89, 89, 89)
LGREY = (217, 217, 217)
ORANGE = (197, 90, 17)
GREEN = (84, 130, 53)
WHITE = (255, 255, 255)
RED = (155, 34, 38)

F = "C:/Windows/Fonts/arial.ttf"
FB = "C:/Windows/Fonts/arialbd.ttf"
FI = "C:/Windows/Fonts/ariali.ttf"


def font(size, bold=False, italic=False):
    return ImageFont.truetype(FB if bold else (FI if italic else F), size)


def wrap(draw, text, fnt, maxw):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if draw.textlength(test, font=fnt) <= maxw:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def text_block(draw, xy, text, fnt, fill, maxw, leading=None):
    x, y = xy
    lead = leading or (fnt.size + 7)
    lines = wrap(draw, text, fnt, maxw)
    for i, ln in enumerate(lines):
        draw.text((x, y + i * lead), ln, font=fnt, fill=fill)
    return y + len(lines) * lead


def box(draw, xy, fill, outline, r=18, width=3):
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)


def arrow(draw, p0, p1, color, width=5, head=18, dash=False):
    x0, y0 = p0; x1, y1 = p1
    ang = math.atan2(y1 - y0, x1 - x0)
    if dash:
        L = math.hypot(x1 - x0, y1 - y0)
        n = max(int(L // 26), 1)
        for i in range(n):
            a = i / n; b = min((i + 0.55) / n, 1)
            draw.line([(x0 + (x1 - x0) * a, y0 + (y1 - y0) * a),
                       (x0 + (x1 - x0) * b, y0 + (y1 - y0) * b)], fill=color, width=width)
    else:
        draw.line([p0, (x1 - head * math.cos(ang) * 0.8, y1 - head * math.sin(ang) * 0.8)], fill=color, width=width)
    hx, hy = x1 - head * math.cos(ang), y1 - head * math.sin(ang)
    left = (hx - head * 1.4 * math.sin(ang), hy + head * 1.4 * math.cos(ang))
    right = (hx + head * 1.4 * math.sin(ang), hy - head * 1.4 * math.cos(ang))
    draw.polygon([(x1, y1), left, right], fill=color)


# ---------------------------------------------------------------- figure 1
def fig_design():
    W, H = 2400, 1100
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)

    box(d, (60, 50, W - 60, 175), LIGHT, BLUE)
    d.text((90, 72), "THEORETICAL LENS", font=font(24, True), fill=NAVY)
    d.text((90, 112), "Critical posthumanism: distributed agency, intra-action, situated knowledge",
           font=font(28), fill=NAVY)
    d.text((1230, 72), "CONCEPTUAL LENS", font=font(24, True), fill=NAVY)
    d.text((1230, 112), "SETILL: ten components, justice criteria, feedback", font=font(28), fill=NAVY)

    stages = [
        ("STAGE 1", "Clarification of requirements",
         ["Integrative literature review", "Theory and concept refinement", "Site scoping and access", "Ethics application"]),
        ("STAGE 2", "Descriptive study I",
         ["Key-informant interviews", "Resident focus groups", "Transect walks and photovoice", "Household energy diaries"]),
        ("STAGE 3", "Prescriptive study",
         ["Co-design workshops", "SETILL model specification", "Biogas system requirements", "Prototype and test"]),
        ("STAGE 4", "Descriptive study II",
         ["Pilot evaluation", "Justice assessment", "Model revision", "Feedback to community"]),
    ]
    xs = [60, 650, 1240, 1830]
    top, h, w = 250, 470, 510
    for i, (num, name, items) in enumerate(stages):
        x = xs[i]
        box(d, (x, top, x + w, top + h), WHITE, NAVY, r=22, width=4)
        d.rounded_rectangle((x, top, x + w, top + 120), radius=22, fill=NAVY)
        d.rectangle((x, top + 80, x + w, top + 120), fill=NAVY)
        d.text((x + 24, top + 18), num, font=font(22, True), fill=(200, 215, 240))
        d.text((x + 24, top + 50), name, font=font(28, True), fill=WHITE)
        y = top + 150
        for it in items:
            d.ellipse((x + 26, y + 11, x + 40, y + 25), fill=BLUE)
            y = text_block(d, (x + 54, y), it, font(24), GREY, w - 80)
            y += 16
        if i < 3:
            arrow(d, (x + w + 8, top + h // 2), (xs[i + 1] - 10, top + h // 2), BLUE, width=6, head=20)

    y_fb = 790
    d.line([(xs[3] + w // 2, top + h), (xs[3] + w // 2, y_fb)], fill=ORANGE, width=4)
    d.line([(xs[3] + w // 2, y_fb), (xs[1] + w // 2, y_fb)], fill=ORANGE, width=4)
    arrow(d, (xs[1] + w // 2, y_fb), (xs[1] + w // 2, top + h + 6), ORANGE, width=4)
    d.text((xs[1] + w // 2 + 60, y_fb - 40),
           "Iterative revision: findings feed back into the model, the requirements and the next co-design cycle",
           font=font(23, italic=True), fill=ORANGE)

    box(d, (60, 900, W - 60, 1000), PALE, LGREY)
    d.text((88, 925), "EMBEDDED SINGLE CASE", font=font(23, True), fill=NAVY)
    d.text((430, 923), "Informal settlement in Greater Tzaneen, Limpopo: one case, four embedded units of analysis "
                       "(households, women's groups, municipal services, technology team)",
           font=font(24), fill=GREY)

    im.save(os.path.join(OUT, "fig1_design.png"), dpi=(300, 300))
    print("fig1")


# ---------------------------------------------------------------- figure 2
def fig_energy():
    W, H = 2400, 1300
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)

    box(d, (60, 40, W - 60, 150), LIGHT, BLUE)
    d.text((90, 68), "Fuel stacking is normal", font=font(30, True), fill=NAVY)
    d.text((430, 72), "Households combine carriers according to cost, reliability, cooking practice and habit, "
                      "not according to a single fuel preference.",
           font=font(25), fill=GREY)

    cols = [
        (60, "ENERGY SOURCES", ["Grid electricity", "Paraffin", "LPG", "Firewood", "Charcoal", "Dry-cell batteries"]),
        (860, "CONVERSION DEVICES", ["Prepaid meter", "Hotplate / stove", "Paraffin stove", "Gas burner", "Open fire / brazier", "Torch, radio, phone"]),
        (1660, "ENERGY SERVICES", ["Lighting", "Cooking", "Space heating", "Water heating", "Phone charging", "Refrigeration"]),
    ]
    top, w, h = 210, 620, 820
    for x, title, items in cols:
        box(d, (x, top, x + w, top + h), WHITE, NAVY, r=20, width=4)
        d.rounded_rectangle((x, top, x + w, top + 80), radius=20, fill=NAVY)
        d.rectangle((x, top + 50, x + w, top + 80), fill=NAVY)
        d.text((x + 24, top + 24), title, font=font(26, True), fill=WHITE)
        y = top + 115
        for it in items:
            d.rounded_rectangle((x + 30, y, x + w - 30, y + 78), radius=12, fill=PALE, outline=LGREY, width=2)
            d.text((x + 50, y + 24), it, font=font(27), fill=GREY)
            y += 105

    arrow(d, (700, 620), (840, 620), BLUE, width=7, head=22)
    d.text((700, 560), "carriers", font=font(22, italic=True), fill=BLUE)
    arrow(d, (1500, 620), (1640, 620), BLUE, width=7, head=22)
    d.text((1500, 560), "useful energy", font=font(22, italic=True), fill=BLUE)

    box(d, (60, 1080, W - 60, 1240), (253, 242, 233), (244, 177, 131))
    d.text((90, 1105), "WHAT THE DIAGRAM DOES NOT SHOW", font=font(24, True), fill=ORANGE)
    d.text((90, 1145), "Unpaid collection and cooking labour · indoor smoke and respiratory risk · fire and paraffin-poisoning risk · "
                       "time poverty · arrears and disconnection · the landlord's meter",
           font=font(24), fill=GREY)
    d.text((90, 1185), "These burdens are distributed unevenly within the household, and they are the reason a purely technical "
                       "efficiency account is insufficient.",
           font=font(24, italic=True), fill=GREY)

    im.save(os.path.join(OUT, "fig2_energy.png"), dpi=(300, 300))
    print("fig2")


# ---------------------------------------------------------------- figure 3
def fig_setill():
    W, H = 2400, 1600
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)

    box(d, (60, 40, W - 60, 170), LIGHT, BLUE, r=16)
    d.text((88, 62), "ANTECEDENTS (CONTEXT)", font=font(23, True), fill=NAVY)
    d.text((88, 104), "Insecure tenure · dense self-built housing · unreliable and costly grid · irregular income · "
                      "organic waste · water access · municipal capacity",
           font=font(26), fill=GREY)

    box(d, (60, 220, 500, 1170), WHITE, NAVY, r=20, width=4)
    d.rounded_rectangle((60, 220, 500, 300), radius=20, fill=NAVY)
    d.rectangle((60, 270, 500, 300), fill=NAVY)
    d.text((84, 244), "STAKEHOLDER SET", font=font(25, True), fill=WHITE)
    stake = ["Residents and women's groups", "Municipal officials (waste, energy, planning)",
             "Technology developers and artisans", "University researchers", "CBOs and NGOs",
             "Traditional and ward leadership"]
    y = 335
    for s in stake:
        y = text_block(d, (85, y), "• " + s, font(24), GREY, 380)
        y += 46

    steps = [
        ("1  CO-DIAGNOSE", "Walk the settlement, map energy practices, identify waste and water constraints"),
        ("2  CO-DESIGN", "Agree the problem frame, the service model and the technical requirements"),
        ("3  PROTOTYPE AND TEST", "Build a community-scale digester; test cooking, safety and maintenance"),
        ("4  EVALUATE AND ADAPT", "Assess justice outcomes, revise the configuration, agree the next cycle"),
    ]
    ys = [220, 455, 690, 925]
    for i, (title, sub) in enumerate(steps):
        y = ys[i]
        box(d, (580, y, 1400, y + 190), WHITE, BLUE, r=18, width=3)
        d.text((610, y + 20), title, font=font(27, True), fill=NAVY)
        text_block(d, (610, y + 68), sub, font(23), GREY, 750)
        if i < 3:
            arrow(d, (990, y + 193), (990, ys[i + 1] - 8), BLUE, width=5, head=16)

    box(d, (1480, 220, 2340, 430), PALE, BLUE, r=18, width=3)
    d.text((1510, 245), "TECHNOLOGY ARTEFACT", font=font(24, True), fill=NAVY)
    text_block(d, (1510, 290), "Community-scale biogas system: feedstock, water, space, digestion, gas, digestate, "
                               "maintenance and spare parts.", font(23), GREY, 790)

    box(d, (1480, 470, 2340, 720), PALE, BLUE, r=18, width=3)
    d.text((1510, 495), "ACCEPTANCE AS RELATIONAL ACHIEVEMENT", font=font(24, True), fill=NAVY)
    text_block(d, (1510, 540), "Not an attitude but a pattern of use that holds when fuel, labour, space, cost, "
                               "smell, safety and social norms line up.", font(23), GREY, 790)

    box(d, (1480, 760, 2340, 1010), (235, 244, 233), GREEN, r=18, width=3)
    d.text((1510, 785), "OUTCOMES", font=font(24, True), fill=(56, 87, 35))
    text_block(d, (1510, 830), "More secure cooking energy; less smoke; lower fire risk; time saved; "
                               "shared control of the service.", font(23), GREY, 790)

    box(d, (1480, 1050, 2340, 1170), PALE, LGREY, r=18, width=2)
    d.text((1510, 1085), "R1 LEARNING LOOP: use feeds back into design", font=font(22, italic=True), fill=BLUE)

    # barrier pathway along the bottom of the middle column
    arrow(d, (520, 1205), (1450, 1205), RED, width=4, head=18, dash=True)
    d.rectangle((525, 1130, 1450, 1178), fill=WHITE)
    d.text((535, 1140), "Barrier pathway (potestas): cost, tenure, space, water, maintenance, elite capture",
           font=font(21, italic=True), fill=RED)

    box(d, (60, 1255, W - 60, 1390), LIGHT, NAVY, r=16)
    d.text((88, 1277), "JUSTICE CRITERIA AND GOVERNANCE", font=font(23, True), fill=NAVY)
    d.text((88, 1321), "Distributive · Procedural · Recognition · Restorative   |   Reflexive monitoring · "
                       "Benefit sharing · Grievance and redress · Exit strategy",
           font=font(26), fill=GREY)

    d.text((60, 1470), "Arrows indicate the direction of influence. The framework is not a stage model, and the "
                       "four steps may overlap.", font=font(20, italic=True), fill=GREY)

    im.save(os.path.join(OUT, "fig3_setill.png"), dpi=(300, 300))
    print("fig3")


# ---------------------------------------------------------------- figure 4
def fig_justice():
    W, H = 2400, 1020
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)

    d.text((W // 2, 45), "Justice is assessed on four dimensions at once", font=font(36, True), fill=NAVY, anchor="ma")
    d.text((W // 2, 105), "A project can score well on one dimension and poorly on another. The framework therefore "
                          "reports a profile, not a single score.", font=font(24, italic=True), fill=GREY, anchor="ma")

    quads = [
        (80, 180, "DISTRIBUTIVE JUSTICE", "Who receives which service, at what cost, and with what reliability?",
         ["Cost incidence across tenure types", "Reliability and duration of supply", "Access for tenants and female-headed households"]),
        (1220, 180, "PROCEDURAL JUSTICE", "Who is in the room, when, and who controls the agenda?",
         ["Timing and location of meetings", "Voice and decision rights", "Grievance and redress channels"]),
        (80, 560, "RECOGNITION", "Whose knowledge and practices count as legitimate?",
         ["Women's and tenants' knowledge", "Local language and idiom", "Cultural cooking practices"]),
        (1220, 560, "RESTORATIVE JUSTICE", "What past harm is repaired, and what future capability is built?",
         ["Repair of past exclusion", "Skills and local enterprise", "Ecological repair and closure plan"]),
    ]
    colors = [NAVY, BLUE, (56, 87, 35), ORANGE]
    for i, (x, y, title, q, items) in enumerate(quads):
        box(d, (x, y, x + 1080, y + 350), WHITE, colors[i], r=22, width=5)
        d.rounded_rectangle((x, y, x + 1080, y + 92), radius=22, fill=colors[i])
        d.rectangle((x, y + 60, x + 1080, y + 92), fill=colors[i])
        d.text((x + 30, y + 24), title, font=font(28, True), fill=WHITE)
        y2 = text_block(d, (x + 30, y + 112), q, font(23, italic=True), GREY, 1010)
        y2 += 12
        for it in items:
            d.ellipse((x + 34, y2 + 10, x + 50, y2 + 26), fill=colors[i])
            y2 = text_block(d, (x + 66, y2), it, font(22), GREY, 960)
            y2 += 9

    d.text((W // 2, 955), "The four dimensions are reported together in Chapter 7 and revisited in the final "
                          "evaluation chapter.", font=font(21, italic=True), fill=GREY, anchor="ma")

    im.save(os.path.join(OUT, "fig4_justice.png"), dpi=(300, 300))
    print("fig4")


# ---------------------------------------------------------------- figure 5
def fig_timeline():
    W, H = 2400, 1150
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)

    phases = [
        ("1. Foundation and ethics", 1, 6, "Proposal approved; ethics certificate", NAVY),
        ("2. Fieldwork", 7, 14, "Interviews, focus groups, diaries, walks", BLUE),
        ("3. Analysis and modelling", 15, 21, "Coded data, situational maps, CLDs", (56, 87, 35)),
        ("4. Co-design and prototyping", 22, 27, "Workshops, specification, pilot build", ORANGE),
        ("5. Evaluation and drafting", 28, 33, "Justice assessment, draft thesis", (112, 48, 160)),
        ("6. Finalisation and submission", 34, 36, "Examination-ready thesis", GREY),
    ]
    left, right = 560, 1870
    top = 150
    rowh = 140
    monthw = (right - left) / 36.0

    for m in range(0, 37, 3):
        x = left + m * monthw
        d.line([(x, top - 20), (x, top + rowh * len(phases) + 10)],
               fill=(150, 150, 150) if m % 6 == 0 else LGREY, width=3 if m % 6 == 0 else 2)
        if m % 6 == 0:
            d.text((x, top - 55), ("M%d" % m) if m else "0", font=font(21), fill=GREY, anchor="ma")
    d.text(((left + right) / 2, 60), "Project month", font=font(24, True), fill=NAVY, anchor="ma")

    for i, (name, s, e, deliverable, col) in enumerate(phases):
        y = top + i * rowh
        d.text((60, y + 30), name, font=font(24, True), fill=GREY)
        x0 = left + (s - 1) * monthw
        x1 = left + e * monthw
        box(d, (x0, y + 18, x1, y + 82), col, col, r=12, width=2)
        d.text(((x0 + x1) / 2, y + 50), "M%d–M%d" % (s, e), font=font(20, True), fill=WHITE, anchor="mm")
        text_block(d, (x1 + 16, y + 30), deliverable, font(19, italic=True), GREY, 470)
        d.polygon([(x1, y + 100), (x1 + 12, y + 112), (x1, y + 124), (x1 - 12, y + 112)], fill=RED)

    d.text((60, top + rowh * len(phases) + 40),
           "Diamonds mark milestones. Fieldwork and analysis overlap because transcription and coding begin while "
           "data collection continues.", font=font(20, italic=True), fill=GREY)

    im.save(os.path.join(OUT, "fig5_timeline.png"), dpi=(300, 300))
    print("fig5")


if __name__ == "__main__":
    fig_design()
    fig_energy()
    fig_setill()
    fig_justice()
    fig_timeline()
    print("DONE")
