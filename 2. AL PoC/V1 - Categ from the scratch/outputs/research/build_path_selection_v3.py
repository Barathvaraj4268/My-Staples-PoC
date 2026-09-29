"""Build the L3/L4 path-selection workbook, v3 (research deliverable, 2026-09-29).

Every Staples path is looked up from the navigation tree by Category ID, so the
paths in the workbook match `Excels/1. Staples_Navigation_Tree_repaired.xlsx`
exactly (asserted below). Scores are analyst inputs (1-5) backed by the desk
research in the 'Evidence & Sources' sheet; weights, thresholds, weighted scores,
ranks, tiers and competitor roles are live Excel formulas.

v2 (2026-09-29) adds the Hero / Probable Hero / Non-Hero lens: a Staples
Strength Index (SSI) and a Market Attractiveness Index (MAI) per path, a
segment + marketplace play per path, a Non-Hero cap on tiers, and a segment map.
Earlier outputs (v1) are left untouched.

v3 (2026-09-29): Top-20 recommendations chosen by segment quota (default 7 Hero,
7 Probable Hero, 6 Non-Hero; minimum 6 each), Non-Hero play = EXPLORE, four
competitors per path (2 primary from the broad-retailer pool, 2 secondary from the
enthusiast / design-specialist pool) with separate criteria, and dedicated
'Hero Criteria' and 'Competitor Criteria' sheets. v1 / v2 outputs and the v2
script (build_path_selection_xlsx.py) are left untouched.

Run:  python outputs/research/build_path_selection_v3.py

Run:  python outputs/research/build_path_selection_xlsx.py
"""
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[2]
TREE = ROOT / "Excels" / "1. Staples_Navigation_Tree_repaired.xlsx"
SKU = ROOT / "data" / "interim" / "sku_staples.parquet"
OUT = ROOT / "outputs" / "research" / "Staples_L3L4_Path_Selection_v3.xlsx"
FIG = ROOT / "outputs" / "research" / "segment_map_v3.png"

# --------------------------------------------------------------------------- data
tree = pd.read_excel(TREE, "Navigation Tree")
tree["path"] = tree[["L1", "L2", "L3", "L4"]].fillna("").apply(
    lambda r: " > ".join(x for x in r if x), axis=1)
tree = tree.set_index("Category ID", drop=False)
in_hand = pd.read_parquet(SKU, columns=["leaf_path"])["leaf_path"].value_counts()
_nav = pd.read_excel(TREE, "Summary by L1", header=3)
NAV = dict(zip(_nav.iloc[:, 0], _nav.iloc[:, 1]))   # L1 -> "Yes"/"No" (in Staples header nav)

W = [20, 15, 15, 15, 10, 10, 5, 10]            # C1..C8 default weights
KW = [30, 25, 15, 15, 15]                      # K1..K5 primary-competitor weights
ES = [25, 25, 15, 20, 15]                      # E1..E5 secondary-competitor weights
MIN_ITEMS, FAST_MIN = 20, 100
QUOTA = {"Hero": 7, "Probable Hero": 7, "Non-Hero": 6}   # Top-20 split (minimum 6 per segment)
RESERVE_N = 3                                  # reserves (swap-ins) per segment
SEG_ORDER = ["Hero", "Probable Hero", "Non-Hero"]
SW = [30, 15, 25, 30]                          # S1..S4 Staples Strength weights
MW = [35, 25, 25, 15]                          # M1..M4 Market Attractiveness weights
H_CUT, P_CUT, PROTECT_C3 = 65, 65, 2
S1_BANDS = [1000, 500, 200, 75]                # items >= band -> 5,4,3,2 else 1
NAV_YES, NAV_NO = 5, 2

# key, category id, link to Sai's list, G2, G4, scores C1..C8, rationale, tier-3 competitor text
L = [
 ("Accent chairs", "CL70203", "1. Chairs & Seating", [5,5,5,4,2,5,3,5], "Pat's own 'White Chair' shelf; pilot: 5 recs, 2 Strong, 94% confident mapping.", ""),
 ("Bar stools", "CL166182", "1. Chairs & Seating", [4,4,4,3,2,4,2,5], "Pilot: best method agreement (rho 0.73 on bar-height split); café/breakroom, not the core work mission.", ""),
 ("Benches", "CL163750", "1. Chairs & Seating", [4,5,4,4,1,3,2,5], "Pilot: 2 Strong recs; thin Staples baseline (49); oversize freight.", ""),
 ("Ottomans", "CL161807", "1. Chairs & Seating", [4,5,4,2,3,2,2,4], "Wayfair 1,069 vs Staples 29 in pilot; residential, weak work mission.", "Wayfair / Amazon"),
 ("Breakroom chairs", "CL166181", "1. Chairs & Seating", [4,5,3,4,2,3,2,5], "Pilot: 5 recs but methods disagree (rho 0.08); RTO breakroom story.", ""),
 ("Office chairs", "CL166253", "1. Chairs & Seating", [3,3,1,5,1,5,4,5], "Core 1P, highest cannibalisation risk. Keep as the CONTROL shelf that shows the gate working.", "Wayfair / Amazon (control shelf)"),
 ("Kids seating", "CL210316", "1. Chairs & Seating", [3,3,3,3,2,2,4,4], "Pilot: no recommended archetypes.", "Wayfair / Target"),
 ("Gaming chairs", "CL210317", "1. Chairs & Seating", [3,3,2,3,1,2,2,3], "Pilot thin (22 families); spec-led.", "Best Buy / Amazon"),
 ("Office desks", "CL210795", "2. Desks", [4,5,3,5,1,5,4,5], "1,501 items; pilot Computer split: 3 recs (1 Strong). Split by Staples' own type facet.", ""),
 ("Sit & stand", "CL215523", "2. Desks", [3,4,2,5,1,4,3,4], "Spec-led (motor/frame); design gaps small, CRS high.", "Amazon / Wayfair"),
 ("Hutches", "CL353865", "2. Desks", [2,3,3,4,1,2,2,3], "Component shelf, little design variance.", "Wayfair / Amazon"),
 ("Desk lamps", "CL167867", "3. Lamps & Lighting", [5,4,4,5,4,4,3,4], "Aesthetic lighting on the work desk; Staples SKUs in hand.", ""),
 ("Floor lamps", "CL216710", "3. Lamps & Lighting", [5,4,4,3,2,3,2,4], "Strong design variance; living-room leaning.", ""),
 ("Table lamps", "CL216712", "3. Lamps & Lighting", [5,4,4,3,3,3,2,4], "Reception / home-office decor lighting.", ""),
 ("Desk organizers", "CL140991", "New", [5,5,4,5,5,5,5,2], "Header-nav core, low ticket, coordinated design collections at Target/Container Store.", ""),
 ("Desk pads", "CL140589", "New", [4,4,4,5,4,3,3,2], "Leather/felt desk pads; Wayfair has a Desk Pads node.", ""),
 ("Planners", "CL166388", "9. Planning & Craft Organization", [5,5,4,5,5,5,5,4], "Deep Staples shelf (1,037) and Staples' own Sincerely Jules move; enthusiast formats missing.", ""),
 ("Calendars", "CL166389", "9. Planning & Craft Organization", [4,3,3,5,5,3,5,2], "Dated goods are core 1P; design calendars are the adjacency.", ""),
 ("Journals", "CL166390", "9. Planning & Craft Organization", [5,4,4,4,5,3,4,2], "Gen-Z journaling habit; pairs with planners.", ""),
 ("Pencil cases", "CL141253", "New", [5,4,5,4,5,4,5,2], "First 'aesthetic' item in the BTS basket; Staples' Recess Club proves demand.", ""),
 ("Pens", "CL110001", "New", [3,4,1,5,5,4,4,2], "Core 1P writing instruments: design pens would substitute best-sellers.", "JetPens / Amazon"),
 ("Label makers", "CL90400", "8. Labels & Packaging Accessories", [3,4,3,5,4,3,3,4], "Mini thermal label printers went viral; Brother/DYMO are core 1P.", ""),
 ("Washi tape", "CL215702", "9. Planning & Craft Organization", [5,5,5,3,5,2,3,1], "Only 32 Staples items: near-pure whitespace, tiny ticket.", ""),
 ("Stickers", "CL167442", "9. Planning & Craft Organization", [5,5,5,3,5,2,3,1], "Planner-companion consumable = repeat commission.", ""),
 ("Craft storage", "CL166121", "9. Planning & Craft Organization", [3,4,4,3,3,2,2,3], "Thin (66) and a hobbyist mission.", "Michaels / Container Store"),
 ("Storage baskets", "CL354705", "New", [5,5,5,4,4,3,5,2], "95 items vs Wayfair's Woven/Decorative/Fabric basket nodes; style-led storage.", ""),
 ("Storage bins", "CL354710", "9. Planning & Craft Organization", [4,4,3,4,4,3,5,2], "Plastic utility bins are core; fabric/acrylic are the adjacency.", ""),
 ("Water bottles", "CL354650", "7. Hydration & Drinkware", [5,5,4,4,5,5,4,4], "Owala/Stanley colourway gap; data in hand; everyone relates.", ""),
 ("Backpacks", "CL142745", "6. Backpacks & Bags", [4,5,3,5,4,5,5,4], "BTS anchor (85% of parents buy one); laptop packs are core 1P.", ""),
 ("Lunch", "CL165124", "6. Backpacks & Bags", [4,5,4,4,5,4,5,4], "Bento/leak-proof designs lead BTS & work-lunch; data in hand.", ""),
 ("Laptop sleeves", "CL219256", "6. Backpacks & Bags", [4,4,4,5,5,3,4,4], "Low-ticket attach to every laptop sale.", ""),
 ("Work totes", "CL160451", "6. Backpacks & Bags", [5,5,4,4,3,3,3,4], "Lifestyle version of the briefcase.", ""),
 ("Keyboards", "CL215783", "New", [4,4,3,5,4,4,3,2], "Typewriter/pastel keyboards mainstream (Logitech POP); business keyboards are core.", ""),
 ("Desk mats", "CL141474", "New", [4,4,4,5,5,3,3,2], "Desk mat is now a decor object; gel wrist rests are core.", ""),
 ("Laptop stands", "CL165071", "New", [4,4,4,5,4,3,3,2], "Wood/aluminium stands are desk-setup objects.", ""),
 ("Monitor stands", "CL141470", "New", [3,4,3,5,3,3,3,2], "Mounts/arms are spec-led; only risers are style-led.", "Amazon / Wayfair"),
 ("Headphones", "CL210844", "New", [4,4,2,4,4,3,4,2], "BTS growth category but Staples already carries 782 items: substitution risk.", "Best Buy / Target"),
 ("Bulletin boards", "CL142403", "New", [4,4,4,5,4,3,4,2], "Letter/felt/fabric memo boards vs cork utility.", ""),
 ("Calendar boards", "CL142105", "New", [4,3,4,5,4,2,4,2], "Acrylic/glass calendars; pairs with planners.", ""),
 ("Whiteboards", "CL166382", "New", [3,3,2,5,3,3,3,2], "Quartet-led core 1P; glass boards the only adjacency.", "Amazon / Wayfair"),
 ("Clocks", "CL161533", "New", [5,5,4,4,4,2,2,2], "Wayfair runs a whole Clocks department.", ""),
 ("Faux plants", "CL355087", "New", [5,5,5,3,4,3,3,2], "Biophilic / resimercial office default.", ""),
 ("Frames", "CL161523", "New", [5,5,4,3,4,3,3,2], "Examples doc already names Michaels on framing.", ""),
 ("Accent tables", "CL161009", "New", [4,5,4,3,1,2,2,2], "Reception-area furniture; oversize.", "Wayfair / Target"),
 ("Coat trees", "CL167079", "New", [4,4,4,3,2,2,2,2], "Entry/office coat racks; niche.", "Wayfair / Amazon"),
 ("Partitions", "CL70600", "New", [4,4,4,4,1,3,3,2], "Decorative dividers & acoustic panels (RTO); freight.", "Wayfair / Amazon"),
 ("Privacy panels", "CL166381", "New", [4,3,4,5,3,2,3,2], "Acoustic felt desk dividers.", "Amazon / Wayfair"),
 ("Trash cans", "CL141928", "New", [4,4,3,4,3,3,2,2], "Facilities whitespace (Pat's doc); design bins for lobbies.", "Wayfair / Target"),
 ("Classroom decor", "CL219245", "New", [5,4,4,5,5,4,5,2], "Teachers self-fund ~$895/yr and buy themed sets: Pat's BTS anchor.", ""),
 ("Classroom storage", "CL207294", "New", [3,3,3,5,4,3,4,2], "Functional; weaker design variance.", "Amazon / Oriental Trading"),
 ("Retail boxes", "CL167951", "8. Labels & Packaging Accessories", [4,4,4,4,4,3,3,4], "Boutique packaging for small sellers: Staples' core SMB customer.", ""),
 ("Mailers", "CL142648", "8. Labels & Packaging Accessories", [3,4,3,5,4,3,3,4], "Coloured/printed/recycled mailers; utility mailers are core 1P.", ""),
 ("Fitness machines", "CL205616", "5. Fitness Equipment", [3,5,5,3,1,3,4,2], "Only as 'active workstation' (walking pads, desk bikes); bulky, thin baseline.", "Amazon / Dick's Sporting Goods"),
 ("Coffee makers", "CL140860", "New", [4,4,3,3,3,2,2,2], "Breakroom design appliances; moderate fit.", "Target / Best Buy"),
 ("Kettles", "CL142180", "New", [4,4,4,3,4,2,2,2], "Design electric kettles for breakrooms.", "Target / Amazon"),
 ("Floor mats", "CL214671", "New", [3,4,3,4,3,2,2,3], "Wayfair doormat pages already scraped; commercial mats are core.", "Wayfair / Amazon"),
 ("Air purifiers", "CL140858", "New", [3,3,3,3,2,2,2,2], "Spec-led wellness appliance.", "Best Buy / Target"),
 ("Diffusers", "CL354884", "New", [4,4,5,2,5,1,2,2], "Wellness decor; weak Staples authority.", "Target / Amazon"),
 ("Coffee organizers", "CL214775", "New", [4,4,4,4,5,2,2,2], "Breakroom coffee station: design-led pod drawers and coffee-bar sets vs utility carousels.", "Amazon / Target"),
 ("Massage tools", "CL214214", "New", [3,5,5,3,3,1,3,2], "Desk-worker recovery tech (massage guns), one of Ben's Scheels examples; odd Staples placement under Spa & Salon.", "Amazon / Best Buy"),
 ("Outdoor games", "CL353835", "New", [4,5,4,2,3,2,3,2], "Company-picnic yard games vs Scheels' premium cornhole / Spikeball depth (Ben's example).", "Amazon / Walmart"),
 ("Drink mixes", "CL350770", "New", [3,4,3,3,5,3,3,2], "Hydration powders (Liquid I.V., Nuun), Ben's Scheels example; food compliance for 3P sellers.", "Amazon / Target"),
 ("Hot glue", "CL354680", "New", [3,4,4,3,4,1,2,1], "Michaels-style temperature-controlled glue guns vs basic guns.", "Amazon / Walmart"),
 ("Portable speakers", "CL167926", "New", [4,4,3,3,4,2,3,2], "Retro / design speakers vs utility speakers.", "Best Buy / Amazon"),
 ("Kitchen storage", "CL142181", "New", [4,4,4,3,4,1,3,2], "Breakroom pantry organisation (acrylic, bamboo).", "Amazon / Target"),
 ("Mirrors", "CL163751", "New", [5,5,5,2,2,2,2,2], "Decor mirrors; residential leaning, oversize.", "Wayfair / Target"),
 ("Candles", "CL354882", "New", [4,4,5,2,5,1,3,2], "Home fragrance; weak Staples authority.", "Target / Amazon"),
]
# gate failures kept in the scorecard so the gates are visible
GATED = [
 ("Rugs", "CL167008", "4. Rugs", "Pass", "Pass", "No L3/L4 exists under Rugs (L2 leaf)."),
 ("Chair mats", "CL70205", "New", "Pass", "Pass", "L2 leaf: no L3/L4 to drill into."),
 ("Labels", "CL142725", "8. Labels & Packaging Accessories", "Pass", "Pass", "L2 leaf of 11,493 commodity labels."),
 ("Gift boxes", "CL215556", "8. Labels & Packaging Accessories", "Pass", "Fail", "Overlaps the Party City partner assortment on Staples.com (Apr-2026)."),
 ("Gift wrap", "CL215563", "8. Labels & Packaging Accessories", "Pass", "Fail", "Overlaps the Party City partner assortment on Staples.com (Apr-2026)."),
 ("Kitchen drinkware", "CL212425", "7. Hydration & Drinkware", "Pass", "Pass", "Only 15 Staples items: no baseline."),
 ("Gift stationery", "CL164333", "New", "Pass", "Pass", "L2 leaf."),
 ("Patio furniture", "CL213655", "New", "Fail", "Pass", "Mis-shelved under Gift Shop > Compasses; leisure, off-mission."),
 ("Luggage", "CL140613", "6. Backpacks & Bags", "Fail", "Pass", "Leisure travel; off Staples' work/school mission."),
]

# Hero lens inputs per path: S3 merchandising investment, S4 shopper association (Staples side);
# M1 demand momentum, M2 competitor destination strength, M3 category size / frequency (market side).
# S1 (depth) and S2 (header nav) are computed from the tree; M4 (moment) reuses C7.
SM = {
 "Accent chairs": (2,2,4,5,4), "Bar stools": (2,2,3,5,4), "Benches": (2,2,3,5,3), "Ottomans": (1,1,3,5,3),
 "Breakroom chairs": (2,3,3,5,4), "Office chairs": (5,5,3,4,5), "Kids seating": (2,2,2,3,2), "Gaming chairs": (2,2,3,4,3),
 "Office desks": (4,5,3,5,4), "Sit & stand": (4,4,3,4,3), "Hutches": (2,3,1,2,1), "Desk lamps": (2,3,4,5,4),
 "Floor lamps": (1,1,3,5,3), "Table lamps": (1,1,3,5,3), "Desk organizers": (4,5,4,5,4), "Desk pads": (2,3,4,3,2),
 "Planners": (5,5,5,5,4), "Calendars": (4,5,3,4,4), "Journals": (3,2,5,4,4), "Pencil cases": (3,3,5,4,4),
 "Pens": (5,5,3,4,5), "Label makers": (3,4,4,4,3), "Washi tape": (1,1,4,4,3), "Stickers": (1,2,3,5,3),
 "Craft storage": (1,1,2,4,2), "Storage baskets": (1,2,4,5,4), "Storage bins": (3,3,3,5,4), "Water bottles": (3,2,5,5,5),
 "Backpacks": (4,4,4,5,5), "Lunch": (3,3,5,5,4), "Laptop sleeves": (3,4,3,4,4), "Work totes": (2,2,3,4,3),
 "Keyboards": (4,4,3,5,4), "Desk mats": (2,3,4,4,3), "Laptop stands": (2,3,4,4,3), "Monitor stands": (3,4,3,4,3),
 "Headphones": (3,3,4,5,5), "Bulletin boards": (4,4,3,4,3), "Calendar boards": (3,4,4,3,2), "Whiteboards": (4,5,2,3,4),
 "Clocks": (1,2,2,5,3), "Faux plants": (2,2,4,5,4), "Frames": (2,3,3,5,4), "Accent tables": (1,1,2,5,3),
 "Coat trees": (1,1,2,4,2), "Partitions": (3,3,3,4,2), "Privacy panels": (3,3,3,3,2), "Trash cans": (3,4,2,4,4),
 "Classroom decor": (4,4,4,4,4), "Classroom storage": (3,3,2,3,3), "Retail boxes": (2,3,3,4,3), "Mailers": (4,5,3,4,5),
 "Fitness machines": (1,1,4,4,3), "Coffee makers": (3,3,3,4,4), "Kettles": (2,2,3,4,3), "Floor mats": (3,4,2,3,3),
 "Air purifiers": (2,2,3,4,3), "Diffusers": (1,1,3,4,3),
 "Coffee organizers": (2,2,3,3,2), "Massage tools": (1,1,4,4,3), "Outdoor games": (1,2,3,4,2), "Drink mixes": (2,3,4,4,5),
 "Hot glue": (1,1,2,4,2), "Portable speakers": (2,2,3,5,4), "Kitchen storage": (1,1,3,4,4), "Mirrors": (1,1,3,5,4), "Candles": (1,1,3,4,4),
 "Rugs": (1,1,3,5,4), "Chair mats": (3,5,3,3,3), "Labels": (5,5,2,3,5), "Gift boxes": (4,3,3,4,3), "Gift wrap": (4,3,3,4,4),
 "Kitchen drinkware": (1,1,4,4,4), "Gift stationery": (2,2,3,3,3), "Patio furniture": (1,1,3,5,4), "Luggage": (2,2,3,4,3),
}
SEG_NOTE = {
 "Planners": "Staples: Sincerely Jules for Blue Sky exclusive (Jun-2026), Vera Bradley; deepest dated-goods shelf. Market: planning / journaling habit.",
 "Desk organizers": "Staples: header-nav Office Supplies staple with Staples-brand ranges. Market: 'desk setup' trend, Target Brightroom.",
 "Backpacks": "Staples: BTS feature category. Market: 85% of parents buy one (Circana).",
 "Classroom decor": "Staples: teacher programme (teacher coupon, teacher lists), 607 items. Market: teachers self-fund about $895 a year.",
 "Water bottles": "Staples: carries Stanley / Hydro Flask but no exclusives; not a drinkware destination. Market: Owala / Stanley / HydroJug craze.",
 "Accent chairs": "Staples: 229 items of utilitarian guest chairs. Market: Wayfair's core design category.",
 "Desk lamps": "Staples: task-lamp led. Market: decor-led lighting is a Wayfair department.",
 "Pencil cases": "Staples: Recess Club exclusive, but only 107 items. Market: the BTS aesthetic essential.",
 "Lunch": "Staples: 404 items, basic insulated bags. Market: bento / leak-proof BTS and work-lunch trend.",
 "Storage baskets": "Staples: 95 items. Market: Wayfair runs Woven, Decorative and Fabric basket nodes.",
 "Office chairs": "Staples' core 1P furniture: protect, and use as the control shelf.",
 "Office desks": "Staples' core furniture (incl. Union & Scale own brand).",
 "Pens": "Staples' core writing instruments: protect.",
 "Labels": "Avery / Staples-brand core, 11,493 items: protect.",
 "Gift boxes": "The Party City partnership now supplies this on Staples.com.",
 "Gift wrap": "The Party City partnership now supplies this on Staples.com.",
 "Clocks": "Neither a Staples strength nor a market favourite (wall clocks are a flat category).",
 "Desk pads": "Small, flat category at Staples (82 items); premium desk pads (Grovemade, Orbitkey) are a niche design object.",
 "Coffee organizers": "Staples: 50 items of K-cup carousels. Market: a small but styled 'coffee bar' niche.",
 "Privacy panels": "Staples: 136 utility screens. Market: acoustic felt is a niche design trend (Felt Right).",
 "Partitions": "Staples: office partitions. Market: folding screens are back in decor but a modest category.",
 "Kettles": "Staples: 55 carafes & kettles. Market: design kettles (Fellow Stagg) are a coffee-enthusiast niche.",
 "Diffusers": "Staples: 46 items. Market: home-fragrance niche led by D2C brands (Vitruvi).",
 "Outdoor games": "Staples: PE-style games. Market: yard games are seasonal and infrequent, but Scheels treats them as a destination.",
}
PLAYS = [
 ("Hero", "EXTEND", "Staples is strong and 1P risk is manageable (C3 above the Protect cut-off).",
  "Add style / lifestyle extensions of the hero only (the 'White Chair' move); strict cannibalisation gate; invite brands Staples already buys first."),
 ("Hero", "PROTECT", "Staples is strong and design variants would substitute 1P (C3 at or below the Protect cut-off).",
  "Keep the marketplace out; use as a control shelf that shows the safety gate working."),
 ("Probable Hero", "BUILD", "The market treats it as a top category but Staples is not (yet) a destination.",
  "Fill the depth gap quickly with marketplace sellers recruited from the primary competitor's brands: the biggest commission upside."),
 ("Non-Hero", "EXPLORE", "Neither a Staples strength nor a market favourite.",
  "Low-risk test through the marketplace: no 1P to protect and sellers hold the inventory. Keep the winners, drop the rest."),
]

# Tier 1 + Tier 2 detail: insight, archetypes, 1P watch-out, moment, room hook, where to look
D = {
 "Planners": ("Staples' deepest stationery shelf (1,037 items) is dated-planner led (Blue Sky, At-A-Glance). Enthusiast formats (disc-bound, undated, sticker-ready, wellness) live at Michaels and Amazon.",
   "Disc-bound customisable planners; undated goal & wellness planners; pastel/floral teacher & student planners; vegan-leather premium planners ($30-60).",
   "Dated 12-month planners incl. the Sincerely Jules exclusive: add other formats, not look-alikes.",
   "Jan-2027 New Year planning; Jul-Aug 2027 academic planners.",
   "Staples launched Sincerely Jules for Blue Sky in Jun-2026. This shows the next step beyond one influencer line.",
   "Amazon: Planners, Refills & Covers (breadth + recruitable sellers). Michaels: Papercraft > Planners (The Happy Planner + companion accessories); near-tie with Amazon."),
 "Water bottles": ("Staples carries 463 desk-use items. Specialists sell full colourways, lids, sizes and accessories from Owala, Stanley, YETI, HydroJug and BruMate.",
   "Sip/chug hybrid bottles; 40 oz handle tumblers in seasonal colourways; kids' 12-16 oz school bottles; lids, boots & straw accessories.",
   "Stanley / Hydro Flask SKUs Staples already sells: gap is colourway & size, not the same SKU.",
   "BTS-2027 (Jul-Sep); Q4 gifting; January hydration resolutions.",
   "Everyone in the room knows the Stanley / Owala craze.",
   "Dick's: Drinkware / Sport Water Bottles. Target: water bottles & tumblers incl. exclusive colourways."),
 "Desk organizers": ("Staples' 342 organizers are mostly black mesh and utility. Target (Brightroom) and Container Store sell coordinated desk collections in acrylic, bamboo, marble, linen and pastel.",
   "Clear acrylic modular sets; warm-wood / bamboo sets; pastel & neutral 'aesthetic desk' collections; turntable caddies.",
   "Black mesh sets (Staples brand, Officemate).",
   "January 'get organized' season; BTS desk & dorm refresh.",
   "Desk organization is a header-nav Staples staple and a low-ticket attach to every paper and pen basket.",
   "Target: Brightroom desk organization + Target Plus brands. Amazon: Desk Organizers. Container Store: Office > Desktop Collections (style benchmark)."),
 "Accent chairs": ("The original 'White Chair' shelf. Pilot (Sep-2026): 5 recommended archetypes, 2 Strong; 94% high-confidence mapping.",
   "Boucle / curved barrel chairs; cream & sage upholstered side chairs with wood legs; rattan & cane lounge chairs; $150-350 design tier.",
   "Black vinyl / fabric guest & reception chairs (Flash, Boss, HON).",
   "Q1-2027 office refresh and return-to-office budgets.",
   "Pat's own example: the thesis in one picture.",
   "Wayfair: Living Room > Accent Chairs and Professional > Office Seating > Accent Chairs. Amazon: Accent Chairs."),
 "Pencil cases": ("107 Staples items. The pencil pouch is the first 'aesthetic' purchase in the BTS basket, at under $20.",
   "Large-capacity multi-layer pouches (Japanese style); clear / jelly pouches; corduroy & quilted pouches; kawaii & character cases.",
   "Staples' own Recess Club pouches: extend beyond them, don't duplicate.",
   "BTS-2027 (Jul-Sep).",
   "Staples' Recess Club / Jellygram exclusives (BTS 2026) already prove the demand.",
   "Amazon: Pencil Cases. Target: pencil cases & pouches. JetPens: Pencil Cases (style benchmark)."),
 "Lunch": ("404 Staples items. Bento compartments, stainless and leak-proof designs lead BTS and work-lunch search.",
   "Bento compartment boxes (kids & adult); neutral / vegan-leather insulated lunch totes; stainless & glass meal-prep sets.",
   "Basic insulated lunch bags Staples already stocks.",
   "BTS-2027; January 'bring lunch to the office' resolutions.",
   "BTS is Pat's named seasonal anchor and lunch is on every parent's list.",
   "Amazon: Lunch Boxes & Bags. Target: lunch boxes & bags (Bentgo, Stasher, owned brands)."),
 "Desk lamps": ("210 Staples items, mostly LED task lamps. Wayfair sells mushroom, pleated, rechargeable and sculptural lamps that sit between task light and decor.",
   "Mushroom / dome lamps; pleated-shade & ceramic-base lamps; rechargeable cordless lamps; brass & warm-wood architect lamps.",
   "LED task lamps with USB ports: Staples' core.",
   "Q1-2027 home-office refresh; BTS dorm.",
   "2026 home-office trend reports all name better lighting and warm materials.",
   "Wayfair: Lighting > Table & Floor Lamps > Desk Lamps (+ Teen Desk Lamps). Amazon: Desk Lamps."),
 "Classroom decor": ("607 Staples items. Teachers self-fund about $895 a year and buy decor as coordinated themes (boho, calm, pastel).",
   "Boho / neutral theme sets; calm-corner & SEL decor; pastel bulletin-board kits; plant & nature themes.",
   "Existing bulletin-board sets: compare theme coverage, not SKU swaps.",
   "BTS-2027 classroom set-up (Jul-Aug); Teacher Appreciation Week (May).",
   "Teachers are a Staples loyalty segment (teacher coupon, teacher lists).",
   "Amazon: Classroom Decorations. Oriental Trading: Teaching Supplies > Classroom Decorations. Michaels: Teacher Supplies."),
 "Backpacks": ("585 Staples items, laptop & commuter led. 85% of parents bought a backpack for BTS; lifestyle backpacks grew 3% in 2025 (Circana).",
   "Corduroy / soft-aesthetic backpacks; Scandinavian-minimal daypacks; mini backpacks; personalisable patch & pin packs.",
   "Laptop backpacks (SwissGear, Targus): watch cannibalisation here.",
   "BTS-2027.",
   "JanSport now sells on Target Plus, the curated-marketplace peer Pat can point to.",
   "Amazon: Backpacks. Target: Backpacks (incl. JanSport on Target Plus)."),
 "Storage baskets": ("Only 95 Staples items, while Wayfair has dedicated Woven Baskets, Decorative Bins and Fabric Baskets nodes. Pure style-led storage.",
   "Seagrass / water-hyacinth woven baskets; felt & linen fabric bins; lined wire baskets; lidded decorative boxes.",
   "Plastic bins sit in the sister node: keep baskets style-led.",
   "January 'get organized' season.",
   "Organization is a January ritual every retailer promotes.",
   "Wayfair: Organization > Storage Containers & Drawers (Woven Baskets, Decorative Bins, Fabric Baskets). Target: Brightroom baskets."),
 "Laptop sleeves": ("300 Staples items. A sleeve is a low-ticket attach to every laptop sale.",
   "Felt & vegan-leather sleeves; pastel / pattern sleeves with pouch; puffer sleeves.",
   "Black Targus / Case Logic sleeves.", "BTS; laptop-attach all year.",
   "Directly serves Pat's '1-2 extra items per checkout'.",
   "Amazon: Laptop Sleeves. Target: laptop sleeves."),
 "Work totes": ("173 Staples items. The work tote is the lifestyle version of the briefcase.",
   "Structured work totes with laptop sleeve; neoprene totes; vegan-leather commuter totes.",
   "Briefcases & padfolios (Samsonite, Solo).", "Q1 return-to-office; graduation gifting.",
   "The 'White Chair' of bags.", "Amazon: work totes. Target: tote bags."),
 "Office desks": ("1,501 Staples items. Pilot Computer-desk split: 3 recommended archetypes (1 Strong). Split by Staples' own type facet.",
   "Fluted / curved-edge writing desks; cane & rattan desks; small-space floating desks; warm-wood Scandinavian desks.",
   "Laminate computer desks (Bush, Sauder, Union & Scale).", "Q1 office refresh.",
   "Second half of the White Chair story.",
   "Wayfair: Office Furniture > Desks + Small Space Offices. Amazon: Home Office Desks."),
 "Journals": ("253 Staples items. Journaling is a Gen-Z wellness habit.",
   "Dotted-grid bullet journals; linen & leather-look hardcovers; guided gratitude / wellness journals.",
   "Notebooks Staples already sells.", "Jan-2027.", "Pairs with the planner story.",
   "Amazon: Journals. Michaels: Papercraft journals."),
 "Faux plants": ("203 Staples items. Plants are the most-used desk personalisation in 2026 home-office trend reports.",
   "Trailing pothos / ivy; small potted succulents in ceramic; olive & fiddle-leaf floor trees; plant + planter sets.",
   "Seasonal faux trees & wreaths.", "Q1 office refresh; spring.",
   "Biophilic, 'resimercial' offices are the return-to-office default.",
   "Wayfair: Flowers & Plants > Faux Plants / Faux Trees. Michaels: Floral."),
 "Stickers": ("166 Staples items. Planner stickers are the consumable that follows every planner sale.",
   "Planner sticker books; functional icon stickers; seasonal packs.",
   "Teacher reward stickers (separate node).", "Jan-2027; BTS.",
   "Repeat-purchase consumable = predictable commission.",
   "Michaels: Happy Planner stickers. Amazon: planner stickers."),
 "Desk mats": ("318 Staples items. The desk mat has become a decor object.",
   "Extended felt desk mats; vegan-leather mats; printed aesthetic mats.",
   "Gel wrist-rest mouse pads (Fellowes, Staples brand).", "BTS; Q1 desk refresh.",
   "In every 'desk setup' video.", "Amazon: Desk & Mouse Pads. Wayfair: Office Organization > Desk Pads."),
 "Washi tape": ("Only 32 Staples items. Washi tape is core to planning & craft and nearly absent at Staples.",
   "Washi tape sets; foil & printed decorative tape; planner tape.",
   "None material: 1P is thin.", "BTS; Jan planning.", "Tiny ticket, high attach to planners.",
   "Michaels: washi tape. Amazon: Washi Tape."),
 "Retail boxes": ("99 Staples items. Small sellers (Etsy, TikTok Shop) want boutique-looking packaging.",
   "Coloured kraft gift / retail boxes with windows; boutique paper shopping bags in colours; tissue + sticker unboxing kits.",
   "Party City gift bags now sold on Staples.com: stay on SMB retail packaging.",
   "Q4 holiday prep (Oct-Nov).", "Staples' core customer is the small business.",
   "Amazon: retail & gift packaging. Uline: Retail Bags / Gift Boxes."),
 "Frames": ("435 Staples items, document & certificate led. Michaels is the framing destination.",
   "Gallery-wall frame sets; floating / acrylic frames; shadow boxes; arched & fluted frames.",
   "Certificate & document frames (Staples core).", "Graduation (May-Jun); holiday gifting.",
   "The Examples doc already cites Michaels on framing.", "Michaels: Frames. Wayfair: Picture Frames & Albums."),
 "Clocks": ("229 Staples items, mostly office wall clocks. Wayfair runs a whole Clocks department.",
   "Oversized wood / metal wall clocks; minimalist silent clocks; mantel & desk clocks.",
   "Atomic / commercial wall clocks.", "Q1 office refresh.", "Every office wall has one.",
   "Wayfair: Decor & Pillows > Clocks. Amazon: Wall Clocks."),
 "Bulletin boards": ("284 Staples items, cork & utility led.",
   "Fabric memo boards; felt letter boards; pinnable felt tiles; framed & acrylic boards.",
   "Cork boards (Quartet, U Brands).", "BTS dorm; Jan planning.", "Letter boards are an office & classroom staple.",
   "Amazon: memo & letter boards. Wayfair: Office Organization > Memo Boards."),
 "Laptop stands": ("205 Staples items. Stands are now wood / aluminium desk objects.",
   "Walnut / bamboo risers; foldable aluminium stands in colours; cushioned lap desks.",
   "Black plastic stands.", "BTS; Q1.", "Desk-setup trend.", "Amazon: Laptop Stands. Best Buy: laptop stands."),
 "Desk pads": ("82 Staples items; overlaps with desk mats.",
   "Vegan-leather dual-colour desk pads; felt blotters.", "Paper desk-pad calendars.", "Q1.",
   "Wayfair even has a Desk Pads node.", "Amazon. Wayfair: Office Organization > Desk Pads."),
 "Table lamps": ("211 Staples items; reception and home-office decor lighting.",
   "Ceramic & rattan table lamps; mushroom lamps; cordless rechargeable lamps.", "Basic lamp sets.", "Q1.",
   "Resimercial lobbies.", "Wayfair: Table Lamps. Target: table lamps (Threshold, Studio McGee)."),
 "Keyboards": ("374 Staples items. Typewriter-style and pastel keyboards went mainstream (Logitech POP).",
   "Retro round-key keyboards; pastel compact 75% boards; colour-matched keyboard + mouse sets.",
   "Logitech / Microsoft business keyboards: core 1P, watch CRS.", "BTS; holiday.",
   "Shows the colour gap inside a brand Staples already sells.", "Amazon: Keyboards. Best Buy: Keyboards (colour filter)."),
 "Benches": ("49 Staples items. Pilot: 2 recommended archetypes, both Strong agreement.",
   "Upholstered entry / reception benches; channel-tufted benches; wood-slat benches.",
   "Beam seating for waiting rooms.", "Q1 office refresh.", "Pilot-proven shelf.",
   "Wayfair: Entry & Mudroom / Bedroom benches, Reception. Amazon."),
 "Calendars": ("1,234 Staples items of dated wall & desk calendars.",
   "Art & photography wall calendars; acrylic reusable calendars; minimalist desk calendars.",
   "At-A-Glance / Blue Sky dated calendars: core 1P.", "Q4-2026 / Jan-2027 dated-goods peak.",
   "Calendar season lines up with Q1 planning.", "Amazon: wall calendars. Target."),
 "Breakroom chairs": ("75 Staples items. Pilot: 5 recommended archetypes but the two methods disagree (rho 0.08).",
   "Coloured moulded-plastic chairs; bentwood café chairs; rattan dining chairs.",
   "Stackable breakroom chairs.", "Q1 return-to-office breakroom upgrades.", "Breakrooms are where RTO budgets go.",
   "Wayfair: dining chairs / Breakroom sets. Amazon."),
 "Mailers": ("396 Staples items of utility mailers.",
   "Coloured & printed poly mailers; recycled / compostable mailers; coloured padded kraft.",
   "Staples-brand & Scotch bubble mailers: core 1P.", "Q4 small-business shipping peak.",
   "Small-business unboxing trend.", "Amazon: poly mailers. Uline: coloured poly mailers."),
 "Label makers": ("148 Staples items (Brother, DYMO). Mini Bluetooth thermal printers went viral for home, classroom and seller labelling.",
   "Pastel mini Bluetooth label printers; sticker printers for small sellers; 4x6 shipping-label printers.",
   "Brother P-touch / DYMO LabelManager: core 1P, check CRS.", "January organize; BTS teachers.",
   "Your 'Labels & Packaging' idea, made specific.", "Amazon: Label Makers. Michaels: Cricut Joy / craft machines."),
 "Floor lamps": ("132 Staples items.", "Arc floor lamps; tree / tri-light lamps; paper-lantern floor lamps.",
   "LED torchieres.", "Q1.", "Pairs with desk & table lamps.", "Wayfair: Floor Lamps. Target."),
 "Bar stools": ("617 Staples items. Pilot bar-height split: 2 Strong recommendations and the best method agreement (rho 0.73).",
   "Rattan / cane counter stools; boucle swivel stools; backless wood saddle stools.",
   "Metal / vinyl commercial stools.", "Q1 breakroom & café upgrades.", "Pilot-proven shelf.",
   "Wayfair: Bar & Counter Stools. Amazon."),
 "Storage bins": ("660 Staples items of plastic utility bins.", "Fabric cube bins in colours; clear acrylic stackers; lidded linen boxes.",
   "Sterilite / Iris plastic bins: core.", "January organize.", "Pairs with storage baskets.",
   "Target: Brightroom. Amazon."),
 "Calendar boards": ("63 Staples items.", "Acrylic dry-erase wall calendars; glass monthly planners; felt calendar boards.",
   "In/out boards.", "January.", "Pairs with planners & calendars.", "Amazon: acrylic calendars. Wayfair: Memo Boards."),
 "Coffee organizers": ("50 Staples items, mostly K-cup carousels and condiment caddies. The coffee corner has become a styled 'coffee bar' at home and in breakrooms.",
   "Ash-wood / bamboo pod drawers and coffee stations; glass canister sets; marble & brass coffee-bar trays; tiered coffee-bar shelves.",
   "Keurig carousels and condiment organisers (low 1P risk).", "Q1-2027 breakroom refresh; holiday hosting.",
   "Every office has a coffee corner.", "Amazon / Target: coffee pod storage. Williams Sonoma: Coffee Bar Accessories, Hold Everything. Crate & Barrel: coffee storage."),
 "Privacy panels": ("136 Staples items of utilitarian desktop screens. Acoustic PET-felt dividers double as pinboards and come in colours and patterns.",
   "Patterned acoustic felt desk dividers (sage, terracotta); pinnable felt screens; clamp-on curved dividers.",
   "Clear acrylic / fabric desktop panels (low risk).", "Q1-2027 return-to-office fit-outs.",
   "Open-plan noise is every office manager's complaint.", "Amazon / Wayfair: desk dividers. Felt Right: Room & Desk Dividers. Etsy: felt panels."),
 "Partitions": ("152 Staples items of office partitions. Folding screens are back in home decor, and room-divider searches peaked in Feb-2026.",
   "Rattan / cane folding screens; slatted wood dividers with shelves; acoustic felt room dividers; tension-rod curtain dividers.",
   "Cubicle panels and mobile partitions (Staples core).", "Q1-2027 office fit-outs; WFH corners.",
   "'Broken-plan' offices and home work corners.", "Wayfair: Room Dividers (Consumer + Pro). Amazon. Felt Right: acoustic dividers. Anthropologie: decorative screens."),
 "Kettles": ("55 Staples items of carafes, kettles & decanters. The gooseneck, temperature-control kettle (Fellow Stagg) is a design object in coffee culture.",
   "Matte-black gooseneck kettles with temperature control; retro enamel kettles; glass tea kettles.",
   "Basic electric kettles and airpots (moderate 1P).", "Q4 gifting; Q1 breakroom refresh.",
   "Coffee culture at the office.", "Amazon / Target: electric kettles. Fellow: Stagg range. Williams Sonoma: kettles."),
 "Diffusers": ("46 Staples items. Wellness decor for desks and reception areas.",
   "Porcelain / stone ultrasonic diffusers; USB desk diffusers; reed diffusers.", "None material.",
   "Q1 wellness resolutions; holiday gifting.", "Wellness at work.", "Amazon / Target. Vitruvi: diffusers. Anthropologie: home fragrance."),
 "Outdoor games": ("78 Staples items of PE-style outdoor games. Scheels sells premium cornhole, Spikeball and giant yard games (one of Ben's examples).",
   "Tournament-grade cornhole boards; Spikeball / roundnet sets; giant yard games; pickleball sets.",
   "Classroom PE games (low risk).", "Summer company picnics (May-Aug 2027).",
   "Corporate picnics and team events.", "Amazon / Walmart. Scheels: Yard Games, Cornhole Boards. Etsy: custom cornhole."),
 "Accent tables": ("203 Staples items for reception areas.", "Travertine / marble side tables; rattan & mango-wood accent tables; nesting tables.",
   "Laminate reception tables.", "Q1 office refresh.", "Reception makeovers.", "Wayfair / Amazon. Article, West Elm, CB2."),
}

# what the two secondary (enthusiast / design) competitors show on each shelf: the 'super-extension'
SEC_LENS = {
 "Planners": "Erin Condren and Michaels show the enthusiast planner system: customisable LifePlanners, disc-bound Happy Planner, stickers and accessories sold as an ecosystem.",
 "Desk organizers": "Grovemade shows the premium end (walnut, cork, modular desk systems); The Container Store shows coordinated acrylic / marble collections.",
 "Backpacks": "Bellroy shows premium commuter design; Urban Outfitters shows fashion-led styles (mini packs, corduroy). Scheels and REI (alternates) show technical packs.",
 "Classroom decor": "Etsy shows handmade theme sets (boho, calm); Oriental Trading shows complete themed classroom kits.",
 "Office desks": "Design Within Reach shows iconic design desks (Herman Miller); West Elm shows mid-century home-office desks.",
 "Bulletin boards": "Etsy shows letter boards and custom felt boards; Pottery Barn Teen shows scalloped, fabric and statement pinboards.",
 "Keyboards": "Keychron and Drop show the enthusiast keyboard world: retro layouts, custom keycaps, hot-swap switches.",
 "Water bottles": "Scheels shows wall-to-wall colourways, lids and accessories; YETI shows customisation and premium colour drops.",
 "Accent chairs": "Article and West Elm show the design-led mid-century / boucle end at $300-600.",
 "Desk lamps": "Lamps Plus shows specialist depth (90,000+ designs); Urban Outfitters shows trend lamps (mushroom, pleated).",
 "Pencil cases": "JetPens and Kinokuniya show Japanese stationery culture: stand-up pouches, multi-compartment cases, character collabs.",
 "Lunch": "Bentgo shows the bento system for kids and adults; Pottery Barn Kids shows personalised, coordinated lunch bundles.",
 "Storage baskets": "The Container Store and Pottery Barn show curated basket collections (seagrass, rattan, lined).",
 "Laptop sleeves": "Bellroy shows premium minimal sleeves; Urban Outfitters shows fashion sleeves.",
 "Clocks": "Schoolhouse shows heirloom-design clocks; Anthropologie shows statement decor clocks.",
 "Desk pads": "Grovemade and Orbitkey show the premium desk pad (vegetable-tanned leather, felt, built-in organisers).",
 "Coffee organizers": "Williams Sonoma and Crate & Barrel show the styled coffee bar (ash-wood stations, glass canisters).",
 "Privacy panels": "Felt Right shows acoustic felt panels as design objects; Etsy shows handmade felt screens.",
 "Partitions": "Felt Right shows acoustic room dividers; Anthropologie shows decorative folding screens.",
 "Kettles": "Fellow and Williams Sonoma show the design / coffee-enthusiast kettle.",
 "Calendars": "Erin Condren and Etsy show design and personalised calendars.",
 "Label makers": "Michaels (Cricut) and The Container Store show the organising / crafting use of label makers.",
 "Mailers": "noissue and Packhelp show custom-branded, eco packaging with low minimums.",
 "Work totes": "Bellroy and Nordstrom show premium work totes.",
 "Journals": "JetPens and Kinokuniya show premium Japanese journals (Hobonichi, Traveler's).",
 "Faux plants": "Michaels (floral) and Anthropologie show premium faux botanicals.",
 "Diffusers": "Vitruvi and Anthropologie show diffusers as decor objects.",
 "Outdoor games": "Scheels shows premium cornhole, Spikeball and giant games; Etsy shows custom boards.",
 "Accent tables": "Article and West Elm show design accent tables.",
}

# ---- PRIMARY pool: broad retailers / marketplaces that measure the gap for the same shopper.
# defaults: K4 seller recruitability, K5 data access (plain-HTTP test 2026-09-29)
CK = {"Wayfair": (3, 4), "Amazon": (5, 3), "Target": (3, 3), "Walmart": (5, 2), "Best Buy": (3, 4),
      "Dick's Sporting Goods": (4, 2), "Uline": (2, 4)}
# per path: (competitor, K1 depth, K2 aesthetic range, K3 shopper overlap)
MP = {
 "Planners": [("Amazon",5,4,4),("Target",3,5,5),("Walmart",4,3,4)],
 "Desk organizers": [("Target",4,5,5),("Amazon",5,3,4),("Wayfair",3,4,3)],
 "Backpacks": [("Amazon",5,4,4),("Target",4,5,5),("Dick's Sporting Goods",4,4,3)],
 "Classroom decor": [("Amazon",5,4,4),("Walmart",4,3,4),("Target",2,5,4)],
 "Office desks": [("Wayfair",5,5,5),("Amazon",5,3,4),("Target",2,4,4)],
 "Bulletin boards": [("Amazon",5,4,4),("Wayfair",4,4,4),("Target",2,5,4)],
 "Keyboards": [("Amazon",5,4,4),("Best Buy",5,4,4),("Walmart",3,2,3)],
 "Water bottles": [("Dick's Sporting Goods",5,5,3),("Target",4,5,5),("Amazon",5,3,4)],
 "Accent chairs": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",3,5,4)],
 "Desk lamps": [("Wayfair",5,5,4),("Amazon",5,3,4),("Target",3,5,5)],
 "Pencil cases": [("Amazon",5,4,4),("Target",4,5,5),("Walmart",4,3,4)],
 "Lunch": [("Amazon",5,4,4),("Target",4,5,5),("Walmart",4,3,4)],
 "Storage baskets": [("Wayfair",5,5,4),("Target",4,5,5),("Amazon",5,3,4)],
 "Laptop sleeves": [("Amazon",5,4,4),("Target",3,4,5),("Best Buy",3,3,4)],
 "Clocks": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",3,5,4)],
 "Desk pads": [("Amazon",5,4,4),("Wayfair",4,4,4),("Target",2,4,4)],
 "Coffee organizers": [("Amazon",5,3,4),("Wayfair",4,4,3),("Target",3,4,4)],
 "Privacy panels": [("Amazon",5,3,4),("Wayfair",4,3,4),("Walmart",3,2,3)],
 "Partitions": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",2,4,3)],
 "Kettles": [("Amazon",5,4,3),("Target",4,5,4),("Walmart",4,2,3)],
 "Calendars": [("Amazon",5,4,4),("Target",3,5,4),("Walmart",4,3,4)],
 "Label makers": [("Amazon",5,4,4),("Walmart",4,3,4),("Best Buy",2,3,3)],
 "Mailers": [("Amazon",5,4,4),("Uline",5,3,5),("Walmart",4,2,3)],
 "Work totes": [("Amazon",5,4,4),("Target",3,4,4),("Walmart",4,2,3)],
 "Journals": [("Amazon",5,4,4),("Target",3,5,4),("Walmart",3,3,3)],
 "Faux plants": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",3,5,4)],
 "Diffusers": [("Amazon",5,3,3),("Target",4,5,4),("Walmart",4,2,3)],
 "Outdoor games": [("Amazon",5,3,3),("Walmart",4,3,3),("Dick's Sporting Goods",4,4,3)],
 "Accent tables": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",3,5,4)],
}
# ---- SECONDARY pool: enthusiast / design specialists outside Staples' competitive set (the 'super-extension').
# defaults: E3 distance from Staples' scope, E5 seller / brand transferability (D2C brands can themselves be invited)
CE = {"Michaels": (4, 3), "Scheels": (5, 4), "REI": (5, 4), "JetPens": (4, 4), "The Container Store": (3, 2),
      "Erin Condren": (4, 5), "Grovemade": (5, 5), "Orbitkey": (5, 5), "Pottery Barn Teen": (5, 1), "Pottery Barn Kids": (5, 1),
      "Pottery Barn": (5, 1), "West Elm": (5, 1), "CB2": (5, 1), "Article": (5, 2), "Design Within Reach": (5, 2),
      "Anthropologie": (5, 3), "Lamps Plus": (4, 3), "Schoolhouse": (5, 2), "Keychron": (5, 5), "Drop": (5, 3),
      "Micro Center": (3, 4), "Kinokuniya": (5, 3), "Oriental Trading": (3, 2), "Really Good Stuff": (3, 2), "Etsy": (4, 3),
      "Bellroy": (5, 5), "Urban Outfitters": (5, 3), "Felt Right": (5, 5), "Williams Sonoma": (5, 2), "Crate & Barrel": (5, 2),
      "Fellow": (5, 5), "Vitruvi": (5, 5), "Bentgo": (4, 5), "YETI": (4, 5), "Nordstrom": (5, 3), "Paper Source": (4, 3),
      "noissue": (5, 5), "Packhelp": (5, 5)}
# per path: (competitor, E1 enthusiast depth, E2 aesthetic authority, E4 adjacency relevance)
MS = {
 "Planners": [("Erin Condren",5,5,5),("Michaels",5,5,5),("JetPens",5,4,3),("Paper Source",3,5,4)],
 "Desk organizers": [("Grovemade",4,5,4),("The Container Store",5,5,5),("Pottery Barn Teen",4,5,3),("JetPens",4,4,3)],
 "Backpacks": [("Bellroy",4,5,4),("Urban Outfitters",4,5,4),("Scheels",5,4,3),("REI",5,4,3)],
 "Classroom decor": [("Etsy",4,5,4),("Oriental Trading",5,4,5),("Really Good Stuff",4,4,5),("Michaels",3,4,4)],
 "Office desks": [("Design Within Reach",5,5,3),("West Elm",4,5,4),("Article",3,5,4),("CB2",3,5,3)],
 "Bulletin boards": [("Etsy",5,5,4),("Pottery Barn Teen",4,5,4),("The Container Store",3,4,4),("Michaels",3,4,3)],
 "Keyboards": [("Keychron",5,5,4),("Drop",5,4,3),("Micro Center",4,3,3)],
 "Water bottles": [("Scheels",5,5,4),("YETI",4,5,4),("REI",4,4,4)],
 "Accent chairs": [("Article",4,5,4),("West Elm",4,5,4),("Anthropologie",3,5,3),("CB2",4,5,3)],
 "Desk lamps": [("Lamps Plus",5,4,4),("Urban Outfitters",3,5,4),("West Elm",3,5,4),("Schoolhouse",3,5,3)],
 "Pencil cases": [("JetPens",5,5,4),("Kinokuniya",4,5,4),("Paper Source",2,4,3)],
 "Lunch": [("Bentgo",4,4,5),("Pottery Barn Kids",4,5,4),("Williams Sonoma",3,4,3)],
 "Storage baskets": [("The Container Store",4,5,5),("Pottery Barn",4,5,4),("Anthropologie",3,5,3)],
 "Laptop sleeves": [("Bellroy",4,5,5),("Urban Outfitters",3,4,4),("Nordstrom",3,4,3)],
 "Clocks": [("Schoolhouse",4,5,4),("Anthropologie",3,5,3),("West Elm",3,5,4)],
 "Desk pads": [("Grovemade",5,5,5),("Orbitkey",4,5,5),("Etsy",4,4,4)],
 "Coffee organizers": [("Williams Sonoma",4,5,4),("Crate & Barrel",3,5,4),("The Container Store",4,4,4)],
 "Privacy panels": [("Felt Right",5,5,5),("Etsy",3,4,3),("Design Within Reach",2,4,3)],
 "Partitions": [("Felt Right",4,5,4),("Anthropologie",3,5,3),("CB2",3,5,3)],
 "Kettles": [("Fellow",4,5,4),("Williams Sonoma",5,5,4),("Crate & Barrel",3,4,3)],
 "Calendars": [("Erin Condren",4,5,5),("Etsy",4,5,3),("Paper Source",3,5,4)],
 "Label makers": [("Michaels",4,4,4),("The Container Store",3,4,4)],
 "Mailers": [("noissue",5,5,4),("Packhelp",5,4,4)],
 "Work totes": [("Bellroy",4,5,4),("Nordstrom",4,5,4),("Urban Outfitters",3,4,3)],
 "Journals": [("JetPens",5,5,4),("Kinokuniya",4,5,3),("Paper Source",3,5,4),("Michaels",3,4,3)],
 "Faux plants": [("Michaels",5,4,3),("Anthropologie",3,5,3),("Pottery Barn",3,5,3)],
 "Diffusers": [("Vitruvi",5,5,3),("Anthropologie",3,5,3),("Nordstrom",3,4,2)],
 "Outdoor games": [("Scheels",5,5,4),("Etsy",4,4,3),("REI",2,3,2)],
 "Accent tables": [("Article",4,5,3),("West Elm",4,5,3),("CB2",4,4,3)],
}

# --------------------------------------------------------------- python mirror of the formulas
def score(v):
    return sum(a * b for a, b in zip(v, W)) / (5 * sum(W)) * 100

def s1_score(items):
    for b, v in zip(S1_BANDS, [5, 4, 3, 2]):
        if items >= b:
            return v
    return 1

def kscore(k1, k2, k3, comp):
    k4, k5 = CK[comp]
    return sum(a * b for a, b in zip([k1, k2, k3, k4, k5], KW)) / (5 * sum(KW)) * 100 + k1 / 1000

def escore(e1, e2, e4, comp):
    e3, e5 = CE[comp]
    return sum(a * b for a, b in zip([e1, e2, e3, e4, e5], ES)) / (5 * sum(ES)) * 100 + e1 / 1000

rows = []
for key, cid, link, sc, why, t3 in L:
    n = tree.loc[cid]
    assert n["Is Terminal"] == "Yes", cid
    rows.append(dict(key=key, cid=cid, link=link, g2="Pass", g4="Pass", sc=sc, why=why, t3=t3, n=n))
for key, cid, link, g2, g4, why in GATED:
    rows.append(dict(key=key, cid=cid, link=link, g2=g2, g4=g4, sc=None, why=why, t3="", n=tree.loc[cid]))
for r in rows:
    n = r["n"]
    r["path"], r["level"], r["items"] = n["path"], int(n["Level"]), int(n["Count"])
    r["url"] = n["URL"]
    r["hand"] = int(in_hand.get(r["path"], 0))
    gates = (r["level"] in (3, 4)) and r["g2"] == "Pass" and r["items"] >= MIN_ITEMS and r["g4"] == "Pass"
    r["gate"] = gates
    r["s"] = score(r["sc"]) if gates and r["sc"] else None
    r["nav"] = NAV.get(n["L1"], "No")
    s3, s4, m1, m2, m3 = SM[r["key"]]
    m4 = r["sc"][6] if r["sc"] else 3
    r["ssi"] = (s1_score(r["items"]) * SW[0] + (NAV_YES if r["nav"] == "Yes" else NAV_NO) * SW[1] + s3 * SW[2] + s4 * SW[3]) / (5 * sum(SW)) * 100
    r["mai"] = (m1 * MW[0] + m2 * MW[1] + m3 * MW[2] + m4 * MW[3]) / (5 * sum(MW)) * 100
    r["seg"] = "Hero" if r["ssi"] >= H_CUT else "Probable Hero" if r["mai"] >= P_CUT else "Non-Hero"
    c3 = r["sc"][2] if r["sc"] else 0
    r["play"] = ("PROTECT" if c3 <= PROTECT_C3 else "EXTEND") if r["seg"] == "Hero" else "BUILD" if r["seg"] == "Probable Hero" else "EXPLORE"
    r["elig"] = gates and r["play"] != "PROTECT"
    r["rk"] = r["s"] + r["sc"][3] / 100 + r["sc"][0] / 1000 if r["elig"] else -1   # score, then C4, then C1
for seg in SEG_ORDER:
    grp = sorted([r for r in rows if r["elig"] and r["seg"] == seg], key=lambda r: -r["rk"])
    for i, r in enumerate(grp, 1):
        r["rank_seg"] = i
for r in rows:
    if not r["gate"]:
        r["status"] = "Parked (gate)"
    elif r["play"] == "PROTECT":
        r["status"] = "Protect (control)"
    elif r["rank_seg"] <= QUOTA[r["seg"]]:
        r["status"] = "Recommended"
    elif r["rank_seg"] <= QUOTA[r["seg"]] + RESERVE_N:
        r["status"] = "Reserve"
    else:
        r["status"] = "Backlog"
by_key = {r["key"]: r for r in rows}
seg_sort = lambda r: (SEG_ORDER.index(r["seg"]), r["rank_seg"])
top20 = sorted([r for r in rows if r["status"] == "Recommended"], key=seg_sort)
reserves = sorted([r for r in rows if r["status"] == "Reserve"], key=seg_sort)
assert len(top20) == sum(QUOTA.values()) == 20 and min(QUOTA.values()) >= 6
covered = {r["key"] for r in top20 + reserves}
assert covered <= set(D) and covered <= set(MP) and covered <= set(MS) and covered <= set(SEC_LENS), covered - set(MS)
assert set(SM) == {r["key"] for r in rows}
roles_p, roles_s = {}, {}
for k in MP:
    rp_ = sorted(MP[k], key=lambda c: -kscore(c[1], c[2], c[3], c[0]))
    rs_ = sorted(MS[k], key=lambda c: -escore(c[1], c[2], c[3], c[0]))
    roles_p[k], roles_s[k] = (rp_[0][0], rp_[1][0]), (rs_[0][0], rs_[1][0])
# every path string must exist verbatim in the navigation tree
assert all(r["path"] in set(tree["path"]) for r in rows)

# --------------------------------------------------------------------------- styles
F = "Arial"
NAVY, TEAL = "1F3864", "2E75B6"
f_body = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_title = Font(name=F, size=16, bold=True, color=NAVY)
f_sub = Font(name=F, size=10, italic=True, color="595959")
f_sec = Font(name=F, size=12, bold=True, color=TEAL)
f_input = Font(name=F, size=10, color="0000FF")
fill_head = PatternFill("solid", fgColor=NAVY)
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_band = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="top", wrap_text=True)
TIER_FILL = {"Tier 1": "C6EFCE", "Tier 2": "FFEB9C", "Tier 3": "E7E6E6", "Parked": "FFC7CE"}
SEG_FILL = {"Hero": "BDD7EE", "Probable Hero": "FCE4D6", "Non-Hero": "EDEDED"}
SEG_HEX = {"Hero": "#2a78d6", "Probable Hero": "#eb6834", "Non-Hero": "#9e9d98"}   # dataviz slots 1-2 + muted grey

wb = Workbook()

def sheet(title, heading, sub):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A2"] = heading, sub
    ws["A1"].font, ws["A2"].font = f_title, f_sub
    return ws

def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill, cell.alignment, cell.border = f_head, fill_head, center, box
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30

def put(ws, row, col, val, font=f_body, align=wrap, fill=None, fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font, c.alignment, c.border = font, align, box
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c

def seg_cf(ws, rng, col_letter, first):
    for t, color in SEG_FILL.items():
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${col_letter}{first}="{t}"'], fill=PatternFill("solid", fgColor=color)))

STATUS_FILL = {"Recommended": "C6EFCE", "Reserve": "FFEB9C", "Backlog": "E7E6E6", "Protect": "DDEBF7", "Parked": "FFC7CE"}

def status_cf(ws, rng, col_letter, first):
    for t, color in STATUS_FILL.items():
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'LEFT(${col_letter}{first},{len(t)})="{t}"'], fill=PatternFill("solid", fgColor=color)))

def sec_title(ws, cell, text):
    ws[cell] = text
    ws[cell].font = f_sec

def hdr_row(ws, row, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    ws.row_dimensions[row].height = 30

def text_rows(ws, first, texts, last_col, height=30, bullet=None):
    for i, t in enumerate(texts):
        if bullet:
            ws.cell(row=first + i, column=1, value=bullet if bullet != "n" else f"{i + 1}.").font = f_bold
        c = ws.cell(row=first + i, column=2 if bullet else 1, value=t)
        c.font, c.alignment = f_body, wrap
        ws.merge_cells(start_row=first + i, start_column=2 if bullet else 1, end_row=first + i, end_column=last_col)
        ws.row_dimensions[first + i].height = height

def tier_cf(ws, rng, col_letter, first):
    for t, color in TIER_FILL.items():
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'LEFT(${col_letter}{first},{len(t)})="{t}"'], fill=PatternFill("solid", fgColor=color)))

# --------------------------------------------------------------- Selection Criteria (path score)
cr = sheet("Selection Criteria", "Path selection: gates and the 8 path-score criteria",
           "Yellow cells with blue text are editable. Hero segmentation lives in 'Hero Criteria'; competitor criteria in 'Competitor Criteria'.")
cr.column_dimensions["A"].width = 8
for col, w in zip("BCDEF", [34, 48, 38, 38, 10]):
    cr.column_dimensions[col].width = w
sec_title(cr, "A4", "A. Hard gates: a path must pass all four to be considered")
hdr_row(cr, 5, ["Code", "Gate", "Rule", "Why it matters"])
gates_txt = [
 ("G1", "Granularity", "The node is an L3 or L4 in the Staples navigation tree (Level 3 or 4).", "L2 is too coarse: the gap must be readable on one shelf."),
 ("G2", "Brand authority", "The node sits on a Staples work / school / small-business mission (not mis-shelved or leisure).", "Pat: build momentum 'without straying too far from Staples' brand authority'."),
 ("G3", "Staples baseline", "Staples lists at least the minimum item count (section C).", "Needs a nearest-Staples-family comparison and a share baseline."),
 ("G4", "No partner conflict", "No Staples strategic partner already owns the assortment on Staples.com.", "Party City (Apr-2026) now supplies party decor, gift bags and wrap."),
]
for i, g in enumerate(gates_txt):
    for j, v in enumerate(g, 1):
        put(cr, 6 + i, j, v, font=f_bold if j == 1 else f_body)
    cr.row_dimensions[6 + i].height = 36
sec_title(cr, "A11", "B. Path-score criteria: each scored 1-5, weighted score scaled to 0-100 (ranks shelves inside each segment)")
hdr_row(cr, 12, ["Code", "Criterion", "What it measures", "Scores 5 when", "Scores 1 when", "Weight"])
crit = [
 ("C1", "Design & lifestyle variance ('White Chair' test)", "Do shoppers choose on colour, material, style or 'vibe' rather than spec?", "Style drives the choice; many aesthetic archetypes exist", "Pure spec / commodity"),
 ("C2", "Competitor depth advantage", "How much wider the best competitor's assortment is on this shelf", "Best competitor is several times wider, with archetypes Staples lacks", "Parity with Staples"),
 ("C3", "1P protection (low cannibalisation)", "Can design variants be added without substituting Staples' core best-sellers?", "Design variants clearly distinct from the 1P core", "Direct substitute of core 1P"),
 ("C4", "Core-adjacency & shopper mission", "Same work, school or small-business mission as Staples' core", "Sits on the core Staples mission", "Residential or leisure mission"),
 ("C5", "Basket attach & quick-win economics", "Parcel-shippable, lower ticket, attaches to core baskets (Pat: +1-2 items per checkout)", "Under ~$50, parcel, frequent attach", "Oversize freight, considered purchase"),
 ("C6", "Business familiarity & story", "Would the Staples room recognise it instantly (header nav, Staples' own moves, visible trend)?", "Header-nav shelf + Staples' own recent move or a famous trend", "Obscure shelf"),
 ("C7", "Commercial moment", "Peaks in the Q1-2027 window or BTS-2027 (Pat's seasonal anchor)", "Peaks in Jan-Mar or Jul-Sep 2027", "No clear moment"),
 ("C8", "Data readiness", "Staples SKUs already in hand + competitor already scraped or reachable", "Both sides in hand / reachable", "Neither side available"),
]
for i, c in enumerate(crit):
    for j, v in enumerate(c, 1):
        put(cr, 13 + i, j, v, font=f_bold if j == 1 else f_body)
    put(cr, 13 + i, 6, W[i], font=f_input, fill=fill_input, align=center)
    cr.row_dimensions[13 + i].height = 36
put(cr, 21, 5, "Total weight", font=f_bold)
put(cr, 21, 6, "=SUM(F13:F20)", font=f_bold, align=center)
sec_title(cr, "A23", "C. Thresholds")
hdr_row(cr, 24, ["Code", "Parameter", "Value", "Note"])
for i, (a, b, v, n) in enumerate([("G3", "Minimum Staples items on the leaf", MIN_ITEMS, "Feeds gate G3."),
                                  ("FS", "Fast-start: min Staples SKUs already in the sample", FAST_MIN, "Top-20 shelves above this can start without a new Staples scrape.")]):
    put(cr, 25 + i, 1, a, font=f_bold); put(cr, 25 + i, 2, b)
    put(cr, 25 + i, 3, v, font=f_input, fill=fill_input, align=center); put(cr, 25 + i, 4, n)
sec_title(cr, "A28", "D. How the Top 20 is chosen")
text_rows(cr, 29, [
    "Gates (section A) remove shelves that are not true L3/L4, are off-mission, too thin, or in a partner's space.",
    "Each remaining shelf gets a Hero segment and a marketplace play ('Hero Criteria').",
    "Each shelf gets a path score, 0-100, from the 8 criteria above.",
    "Inside each segment, shelves are ranked by path score (ties: C4 core-adjacency, then C1 design variance). PROTECT shelves are left out: they are control shelves.",
    "The top 7 Hero + 7 Probable Hero + 6 Non-Hero become the Top 20 (quotas editable in 'Hero Criteria', minimum 6 each); the next 3 per segment are reserves.",
    "Each Top-20 and reserve shelf gets 2 primary + 2 secondary competitors ('Competitor Criteria').",
], 6, 28, bullet="n")
WREF = [f"'Selection Criteria'!$F${13 + i}" for i in range(8)]
SC_ = "'Selection Criteria'!"

# --------------------------------------------------------------- Hero Criteria
hc = sheet("Hero Criteria", "Hero segmentation criteria: Hero, Probable Hero, Non-Hero",
           "How every L3/L4 shelf is placed in a segment, how each segment is played, and how many of each go into the Top 20. Yellow cells are editable.")
for col, w in zip("ABCDEFGHIJ", [8, 30, 46, 9, 22, 22, 22, 22, 22, 26]):
    hc.column_dimensions[col].width = w
sec_title(hc, "A4", "1. What the three segments mean")
hdr_row(hc, 5, ["Segment", "In plain words", "The question it answers", "Play", "What Staples Marketplace does"])
hc.merge_cells("E5:I5")
seg_def = [
 ("Hero", "Staples' own top categories: where Staples already has depth, visibility, investment and shopper trust. (Proxy for 'top by sales / profit / assortment', built from external data.)",
  "Is Staples already known for this?", "EXTEND (or PROTECT)",
  "Add design / lifestyle extensions of what Staples already sells well. Keep the marketplace out where variants would copy 1P best-sellers (PROTECT)."),
 ("Probable Hero", "The market's top categories where Staples is not (yet) a destination: room for improvement.",
  "Does the market love it while Staples is still thin?", "BUILD",
  "Fill the depth gap fast with marketplace sellers recruited from the competitors' brands. The biggest commission upside."),
 ("Non-Hero", "Neither a Staples strength nor a market favourite today.",
  "Is it a low-risk experiment worth a small test?", "EXPLORE",
  "Test a small, curated range through sellers: no inventory and no 1P risk. Keep the winners, drop the rest."),
]
for i, sd in enumerate(seg_def):
    rr = 6 + i
    for j, v in enumerate(sd[:4], 1):
        put(hc, rr, j if j == 1 else j, v, font=f_bold if j in (1, 4) else f_body)
    put(hc, rr, 5, sd[4]); hc.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=9)
    hc.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=SEG_FILL[sd[0]])
    hc.row_dimensions[rr].height = 48
sec_title(hc, "A10", "2. The decision rule (applied in this order)")
text_rows(hc, 11, [
    "Compute two scores for every L3/L4 shelf: the Staples Strength Index (SSI, section 3) and the Market Attractiveness Index (MAI, section 4), each 0-100.",
    "If SSI >= the Hero cut-off, the shelf is a HERO. Inside Hero: if the shelf's 1P-protection score (C3 on the Path Scorecard) <= the Protect cut-off, the play is PROTECT (control shelf); otherwise EXTEND.",
    "Otherwise, if MAI >= the Probable Hero cut-off, the shelf is a PROBABLE HERO (play BUILD).",
    "Otherwise the shelf is a NON-HERO (play EXPLORE).",
    "Rank the shelves inside each segment by path score ('Selection Criteria') and take the quota per segment (section 5) = the Top 20. The next shelves are reserves.",
], 9, 30, bullet="n")
hc["A11"].value = "Step 1"; hc["A12"].value = "Step 2"; hc["A13"].value = "Step 3"; hc["A14"].value = "Step 4"; hc["A15"].value = "Step 5"
ANCH_S = [
 ("S1", "Assortment depth", "Staples items on the leaf (the only 'size' signal available without sales data)", SW[0],
  [f"< {S1_BANDS[3]} items", f"{S1_BANDS[3]}-{S1_BANDS[2] - 1}", f"{S1_BANDS[2]}-{S1_BANDS[1] - 1}", f"{S1_BANDS[1]}-{S1_BANDS[0] - 1}", f">= {S1_BANDS[0]:,}"],
  "Computed from the tree (bands in section 5)"),
 ("S2", "Shelf visibility", "Is the shelf's L1 in Staples' header navigation?", SW[1],
  ["-", "L1 not in header nav", "-", "-", "L1 in header nav"], "Computed from the tree ('Summary by L1')"),
 ("S3", "Merchandising investment", "Exclusives, collaborations, own brands, BTS / seasonal feature slots", SW[2],
  ["No visible investment; generic range", "Occasional promos only", "Seasonal features or some own brand", "Own brand or an exclusive line + feature slots",
   "Exclusive collabs + own brand + hero campaigns (e.g. Sincerely Jules planners)"], "Analyst: Staples press releases & site"),
 ("S4", "Shopper association", "Would a shopper think of Staples first for this?", SW[3],
  ["Nobody thinks of Staples for this", "Occasional convenience buy", "One of several options", "Staples is a usual stop", "Staples is the default destination"],
  "Analyst judgement"),
]
ANCH_M = [
 ("M1", "Demand momentum", "Growth and trend signals", MW[0],
  ["Declining", "Flat", "Steady growth", "Clear growth / trending on social", "Viral or double-digit growth (e.g. Owala, Stanley)"],
  "Analyst: Circana, trade press, social"),
 ("M2", "Competitor destination strength", "Do competitors run it as a department with specialist depth?", MW[1],
  ["No competitor features it", "One competitor", "A few mass retailers carry decent depth", "Mass + a specialist run it as a department",
   "Several specialists and mass leaders treat it as a destination"], "Analyst: competitor site structure"),
 ("M3", "Category size & purchase frequency", "How big the category is and how often it is bought", MW[2],
  ["Niche, rare purchase", "Small or infrequent", "Mid-sized, yearly", "Large or several times a year", "Very large and frequent"], "Analyst"),
 ("M4", "Occasion strength", "Seasonal peak in the planning window", MW[3],
  ["No moment", "Weak", "Some seasonality", "Clear Q1 or BTS peak", "Peaks in both Q1-2027 and BTS-2027"], "Same as C7 on the Path Scorecard"),
]
for top, title, items in [(17, "3. Staples Strength Index (SSI): is it already a Staples hero?", ANCH_S),
                          (25, "4. Market Attractiveness Index (MAI): is it a market hero?", ANCH_M)]:
    sec_title(hc, f"A{top}", title)
    hdr_row(hc, top + 1, ["Code", "Component", "What it measures", "Weight", "Score 1", "Score 2", "Score 3", "Score 4", "Score 5", "How it is scored"])
    for i, (code, name, what, wt, anchors, src) in enumerate(items):
        rr = top + 2 + i
        put(hc, rr, 1, code, font=f_bold); put(hc, rr, 2, name, font=f_bold); put(hc, rr, 3, what)
        put(hc, rr, 4, wt, font=f_input, fill=fill_input, align=center)
        for j, a in enumerate(anchors):
            put(hc, rr, 5 + j, a)
        put(hc, rr, 10, src)
        hc.row_dimensions[rr].height = 44
    put(hc, top + 6, 3, "Total weight", font=f_bold)
    put(hc, top + 6, 4, f"=SUM(D{top + 2}:D{top + 5})", font=f_bold, align=center)
    put(hc, top + 6, 5, "Index = sum(score x weight) / (5 x total weight) x 100", font=f_sub)
    hc.merge_cells(start_row=top + 6, start_column=5, end_row=top + 6, end_column=9)
sec_title(hc, "A33", "5. Cut-offs, quotas and settings")
hdr_row(hc, 34, ["Code", "Setting", "Value", "Note"])
hc.merge_cells("D34:I34")
settings = [("H", "Hero cut-off (SSI >=)", H_CUT, "Checked first: a strong Staples shelf is a Hero whatever the market does."),
            ("P", "Probable Hero cut-off (MAI >=)", P_CUT, "Applied to shelves that are not Heroes."),
            ("PR", "Protect cut-off (Hero with C3 <=)", PROTECT_C3, "Hero shelves where design variants would substitute 1P: keep the marketplace out."),
            ("QH", "Top-20 quota: Hero", QUOTA["Hero"], "Brief: 20 shelves, at least 6 per segment."),
            ("QP", "Top-20 quota: Probable Hero", QUOTA["Probable Hero"], ""),
            ("QN", "Top-20 quota: Non-Hero", QUOTA["Non-Hero"], ""),
            ("RS", "Reserves per segment", RESERVE_N, "Swap-ins if a Top-20 shelf is dropped."),
            ("B5", "S1 band: items >= this scores 5", S1_BANDS[0], ""), ("B4", "S1 band: scores 4", S1_BANDS[1], ""),
            ("B3", "S1 band: scores 3", S1_BANDS[2], ""), ("B2", "S1 band: scores 2 (else 1)", S1_BANDS[3], ""),
            ("NY", "S2 score if the L1 is in the header nav", NAV_YES, ""), ("NN", "S2 score if not", NAV_NO, "")]
for i, (a, b, v, n) in enumerate(settings):
    rr = 35 + i
    put(hc, rr, 1, a, font=f_bold); put(hc, rr, 2, b)
    put(hc, rr, 3, v, font=f_input, fill=fill_input, align=center); put(hc, rr, 4, n)
    hc.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=9)
put(hc, 48, 2, "Top-20 check (live)", font=f_bold)
put(hc, 48, 3, "=SUM(C38:C40)", font=f_bold, align=center)
put(hc, 48, 4, '=IF(AND(MIN(C38:C40)>=6,SUM(C38:C40)=20),"OK: 20 shelves, at least 6 per segment","Check: the brief asks for 20 shelves with at least 6 per segment")', font=f_bold)
hc.merge_cells("D48:I48")
sec_title(hc, "A50", "6. Marketplace plays")
hdr_row(hc, 51, ["Segment", "Play", "When", "What Staples Marketplace does"])
hc.merge_cells("D51:I51")
for i, pl in enumerate(PLAYS):
    rr = 52 + i
    for j, v in enumerate(pl, 1):
        put(hc, rr, j, v, font=f_bold if j <= 2 else f_body)
    hc.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=9)
    hc.cell(row=rr, column=1).fill = PatternFill("solid", fgColor=SEG_FILL[pl[0]])
    hc.row_dimensions[rr].height = 36
HC_ = "'Hero Criteria'!"
SREF = [f"{HC_}$D${19 + i}" for i in range(4)]
MREF = [f"{HC_}$D${27 + i}" for i in range(4)]
H_REF, P_REF, PR_REF = f"{HC_}$C$35", f"{HC_}$C$36", f"{HC_}$C$37"
QH_REF, QP_REF, QN_REF, RS_REF = f"{HC_}$C$38", f"{HC_}$C$39", f"{HC_}$C$40", f"{HC_}$C$41"
B_REF = [f"{HC_}$C${42 + i}" for i in range(4)]
NY_REF, NN_REF = f"{HC_}$C$46", f"{HC_}$C$47"
SSI_TOT, MAI_TOT = f"{HC_}$D$23", f"{HC_}$D$31"

# --------------------------------------------------------------- Competitor Criteria
cc = sheet("Competitor Criteria", "Competitor criteria: 2 primary + 2 secondary competitors per shelf",
           "Primary = broad retailers that measure the gap. Secondary = enthusiast / design specialists outside Staples' scope that show the 'super-extension'. Yellow cells are editable.")
cc.column_dimensions["A"].width = 8
for col, w in zip("BCDEF", [34, 52, 40, 36, 10]):
    cc.column_dimensions[col].width = w
sec_title(cc, "A4", "1. Why four competitors, in two roles")
text_rows(cc, 5, [
    "PRIMARY (2 per shelf): broad retailers and marketplaces that sell to the same shopper at scale. They measure the gap (archetypes, price tiers, colourways Staples lacks) and feed the share-based gap maths. Chosen with K1-K5 from the primary pool.",
    "SECONDARY (2 per shelf): enthusiast or design specialists outside Staples' normal competitive set, like Scheels and Michaels in Ben's examples. They show the 'super-extension': how far the category stretches in depth, premium tier and aesthetics. Chosen with E1-E5 from the secondary pool.",
    "Secondaries are a style / upside lens, not a share baseline. Many are D2C brands (Erin Condren, Grovemade, Keychron, Bellroy, Felt Right, Fellow, Bentgo, YETI, noissue) that could list on Staples Marketplace themselves: a ready seller-recruitment list.",
], 6, 40, bullet="•")
for top, title, rows_ in [
    (9, "2. Primary competitor criteria (K1-K5)", [
        ("K1", "Assortment depth on this shelf", "Breadth of the competitor's range on the mapped node", "Several times Staples' range, many sub-types", "Similar or smaller than Staples"),
        ("K2", "Aesthetic & lifestyle range", "Variety of styles, colourways, materials and price tiers", "Destination for design-led choice", "Utility only"),
        ("K3", "Shopper overlap with Staples", "Same home-office, student, teacher or small-business shopper", "Same shopper and mission", "Different shopper"),
        ("K4", "Seller recruitability", "Can the brands / suppliers behind the range be invited onto Staples Marketplace?", "Multi-brand 3P or drop-ship suppliers", "Mostly own / private label"),
        ("K5", "Data accessibility", "Can the team collect listing-level data at PoC speed? (tested 29-Sep-2026)", "Already scraped or official API", "Bot-blocked, needs a paid proxy")]),
    (18, "3. Secondary competitor criteria (E1-E5): the 'super-extension' lens", [
        ("E1", "Enthusiast depth", "How far beyond Staples the specialist's range goes (technical, premium, hobby-grade sub-types)", "Enthusiast destination: many sub-types, premium tiers, accessories", "Similar to Staples"),
        ("E2", "Aesthetic authority", "Does it set the look: design-led, trend-setting, curated?", "A trend-setter the shopper follows", "Generic look"),
        ("E3", "Distance from Staples' scope", "Is it outside Staples' normal competitive set (not an office-supply or mass retailer)?", "Clearly outside (specialist, D2C, lifestyle)", "A direct Staples competitor"),
        ("E4", "Adjacency relevance", "Would the extension still make sense to a Staples shopper? (guards against off-brand picks)", "Obvious next step for the Staples shopper", "Unrelated hobby or luxury"),
        ("E5", "Seller / brand transferability", "Could its brands, or the brand itself if D2C, list on Staples Marketplace?", "D2C brand or multi-brand seller base", "Only own / private label")])]:
    sec_title(cc, f"A{top}", title)
    hdr_row(cc, top + 1, ["Code", "Criterion", "What it measures", "Scores 5 when", "Scores 1 when", "Weight"])
    wts = KW if top == 9 else ES
    for i, c in enumerate(rows_):
        rr = top + 2 + i
        for j, v in enumerate(c, 1):
            put(cc, rr, j, v, font=f_bold if j == 1 else f_body)
        put(cc, rr, 6, wts[i], font=f_input, fill=fill_input, align=center)
        cc.row_dimensions[rr].height = 34
    put(cc, top + 7, 5, "Total weight", font=f_bold)
    put(cc, top + 7, 6, f"=SUM(F{top + 2}:F{top + 6})", font=f_bold, align=center)
sec_title(cc, "A28", "4. Competitor pools")
hdr_row(cc, 29, ["Pool", "Members", "Rule"])
cc.merge_cells("C29:F29")
pools = [("Primary", ", ".join(CK), "Same shopper, broad assortment, data at scale. Office Depot is excluded: its range mirrors Staples', so it shows little gap."),
         ("Secondary", ", ".join(sorted(CE)), "Enthusiast, design or D2C specialists that are not direct Staples competitors. Dick's is primary (national mass sporting); Scheels and REI are secondary (enthusiast).")]
for i, (a, b, c_) in enumerate(pools):
    rr = 30 + i
    put(cc, rr, 1, a, font=f_bold); put(cc, rr, 2, b); put(cc, rr, 3, c_)
    cc.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
    cc.row_dimensions[rr].height = 70
sec_title(cc, "A33", "5. How the four are picked")
text_rows(cc, 34, [
    "For every Top-20 and reserve shelf, 3 plausible primaries and 2-4 plausible secondaries are scored 1-5 on the shelf-specific criteria (K1-K3; E1, E2, E4). The competitor-level criteria (K4, K5; E3, E5) are fixed per competitor ('Competitor Profiles').",
    "Primary 1 and 2 = the two highest primary fit scores ('Primary Competitor Fit'). Secondary 1 and 2 = the two highest secondary fit scores ('Secondary Competitor Fit'). Ties are broken by depth (K1 / E1).",
    "Everything is live: change a weight here or a score in the fit sheets and the picks update across the workbook.",
    "Scores are analyst judgement from desk research ('Evidence & Sources'). Confirm depth when listings are collected; secondaries can be sampled manually because they are a style lens, not a share baseline.",
], 6, 34, bullet="n")
CC_ = "'Competitor Criteria'!"
KREF = [f"{CC_}$F${11 + i}" for i in range(5)]
EREF = [f"{CC_}$F${20 + i}" for i in range(5)]
K_TOT, E_TOT = f"{CC_}$F$16", f"{CC_}$F$25"

# --------------------------------------------------------------- Primary / Secondary Competitor Fit
pid = {}
for i, r in enumerate(sorted([r for r in rows if r["gate"]], key=lambda r: -r["s"]), 1):
    pid[r["key"]] = f"P{i:02d}"
xi = 0
for r in rows:
    if not r["gate"]:
        xi += 1
        pid[r["key"]] = f"X{xi:02d}"
order = sorted(rows, key=lambda r: (not r["gate"], -(r["s"] or 0)))
SC_FIRST, SC_LAST = 5, 5 + len(order) - 1
HS_FIRST, HS_LAST = 5, 5 + len(order) - 1
HS_ = "'Hero Segmentation'!"
fit_paths = top20 + reserves
FIT = {}
for kind, title, pool, inputs, defaults, refs, tot, roles_lbl in [
    ("P", "Primary Competitor Fit", MP, ("K1 Depth", "K2 Aesthetic range", "K3 Shopper overlap"), CK, KREF, K_TOT, ("Primary 1", "Primary 2")),
    ("S", "Secondary Competitor Fit", MS, ("E1 Enthusiast depth", "E2 Aesthetic authority", "E4 Adjacency relevance"), CE, EREF, E_TOT, ("Secondary 1", "Secondary 2"))]:
    ws = sheet(title, f"{title}: every Top-20 and reserve shelf",
               "Blue = shelf-specific inputs; black = competitor-level defaults (from 'Competitor Profiles'). Rank 1-2 in each shelf = the two picks.")
    cols = (["Path ID", "Staples path (exact)", "Status (live)", "Competitor"] +
            ([inputs[0], inputs[1], inputs[2], "K4 Recruitable", "K5 Data access"] if kind == "P"
             else [inputs[0], inputs[1], "E3 Distance from Staples", inputs[2], "E5 Transferability"]) +
            ["Fit score", "Rank in shelf", "Role", "Lookup key", "What the secondaries show" if kind == "S" else "Note"])
    header(ws, 4, cols, [8, 56, 13, 22, 10, 10, 10, 10, 10, 9, 8, 13, 16, 60])
    rr = 5
    for r in fit_paths:
        for comp, a, b, c in pool[r["key"]]:
            d1, d2 = defaults[comp]
            vals = [a, b, d1, c, d2] if kind == "S" else [a, b, c, d1, d2]
            put(ws, rr, 1, pid[r["key"]], align=center); put(ws, rr, 2, r["path"])
            put(ws, rr, 3, f"=INDEX('Path Scorecard'!$AB${SC_FIRST}:$AB${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))", align=center)
            put(ws, rr, 4, comp, font=f_bold)
            inp_cols = (5, 6, 8) if kind == "S" else (5, 6, 7)
            for j, v in enumerate(vals):
                col = 5 + j
                put(ws, rr, col, v, font=f_input if col in inp_cols else f_body, fill=fill_input if col in inp_cols else None, align=center)
            q = "+".join(f"{get_column_letter(5 + j)}{rr}*{refs[j]}" for j in range(5))
            put(ws, rr, 10, f"=({q})/(5*{tot})*100+E{rr}/1000", align=center, fmt="0.0")
            rr += 1
    first, last = 5, rr - 1
    for i in range(first, last + 1):
        put(ws, i, 11, f"=COUNTIFS($A${first}:$A${last},A{i},$J${first}:$J${last},\">\"&J{i})+1", align=center)
        put(ws, i, 12, f'=IF(K{i}=1,"{roles_lbl[0]}",IF(K{i}=2,"{roles_lbl[1]}","Alternate"))', font=f_bold)
        put(ws, i, 13, f'=A{i}&"|"&L{i}')
    if kind == "S":
        prev = None
        for i in range(first, last + 1):
            pth = ws.cell(row=i, column=2).value
            if pth != prev:
                k = next(r["key"] for r in fit_paths if r["path"] == pth)
                put(ws, i, 14, SEC_LENS[k])
                prev = pth
    ws.freeze_panes = "E5"
    ws.auto_filter.ref = f"A4:N{last}"
    ws.conditional_formatting.add(f"A{first}:N{last}", FormulaRule(formula=[f'LEFT($L{first},7)="Primary"'], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(f"A{first}:N{last}", FormulaRule(formula=[f'LEFT($L{first},9)="Secondary"'], fill=PatternFill("solid", fgColor="FCE4D6")))
    dvf = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True)
    ws.add_data_validation(dvf)
    for col in ((5, 6, 8) if kind == "S" else (5, 6, 7)):
        dvf.add(f"{get_column_letter(col)}{first}:{get_column_letter(col)}{last}")
    FIT[kind] = (title, first, last)

def comp_lookup(p_cell, role, fallback):
    title, first, last = FIT["P" if role.startswith("Primary") else "S"]
    rng_c = f"'{title}'!$D${first}:$D${last}"
    rng_k = f"'{title}'!$M${first}:$M${last}"
    fb = f'"{fallback}"' if fallback else '"-"'
    return f'=IFERROR(INDEX({rng_c},MATCH({p_cell}&"|{role}",{rng_k},0)),{fb})'

# --------------------------------------------------------------- Path Scorecard
sc = sheet("Path Scorecard", "Path scorecard: every candidate Staples L3/L4 path",
           "Blue cells are analyst inputs. Gates, segment, score, rank in segment, status, play and competitors are formulas. Sort or filter freely; IDs are stable.")
cols = ["Path ID", "Staples path (exact, from nav tree)", "L1", "L2", "L3", "L4", "Level", "Category ID", "Staples items (site)",
        "Staples SKUs in hand (sample)", "Link to your list", "G1 Granularity", "G2 Brand authority", "G3 Baseline", "G4 No partner conflict",
        "All gates", "Hero segment", "C1 Design variance", "C2 Comp. depth", "C3 1P protection", "C4 Core-adjacency", "C5 Basket attach",
        "C6 Familiarity", "C7 Moment", "C8 Data readiness", "Path score (0-100)", "Rank in segment", "Status", "Marketplace play", "Fast-start",
        "Primary 1", "Primary 2", "Secondary 1", "Secondary 2", "Rationale / gate note", "Rank key (helper)"]
header(sc, 4, cols, [8, 60, 18, 20, 24, 24, 6, 11, 9, 10, 22, 9, 9, 9, 9, 8, 13] + [7] * 8 + [8, 7, 14, 11, 8, 18, 18, 20, 20, 60, 9])
C_SEG, C_SCORE, C_RANK, C_STATUS, C_PLAY, C_FAST = "Q", "Z", "AA", "AB", "AC", "AD"
C_P1, C_P2, C_S1, C_S2 = "AE", "AF", "AG", "AH"
C3_COL, C7_COL = "T", "X"
QUOTA_F = lambda rr: f'IF(Q{rr}="Hero",{QH_REF},IF(Q{rr}="Probable Hero",{QP_REF},{QN_REF}))'
for i, r in enumerate(order):
    rr = SC_FIRST + i
    n = r["n"]
    base = [pid[r["key"]], r["path"], n["L1"], n["L2"], n["L3"] if pd.notna(n["L3"]) else "", n["L4"] if pd.notna(n["L4"]) else "",
            r["level"], r["cid"], r["items"], r["hand"], r["link"]]
    for j, v in enumerate(base, 1):
        put(sc, rr, j, v, align=center if j in (1, 7, 8, 9, 10) else wrap, fmt="#,##0" if j in (9, 10) else None)
    put(sc, rr, 12, f'=IF(OR(G{rr}=3,G{rr}=4),"Pass","Fail")', align=center)
    put(sc, rr, 13, r["g2"], font=f_input, fill=fill_input, align=center)
    put(sc, rr, 14, f"=IF(I{rr}>={SC_}$C$25,\"Pass\",\"Fail\")", align=center)
    put(sc, rr, 15, r["g4"], font=f_input, fill=fill_input, align=center)
    put(sc, rr, 16, f'=IF(AND(L{rr}="Pass",M{rr}="Pass",N{rr}="Pass",O{rr}="Pass"),"Pass","Fail")', font=f_bold, align=center)
    put(sc, rr, 17, f"=INDEX({HS_}$Q${HS_FIRST}:$Q${HS_LAST},MATCH($A{rr},{HS_}$A${HS_FIRST}:$A${HS_LAST},0))", font=f_bold, align=center)
    for j in range(8):
        v = r["sc"][j] if r["sc"] else None
        put(sc, rr, 18 + j, v, font=f_input, fill=fill_input, align=center)
    q = "+".join(f"{get_column_letter(18 + j)}{rr}*{WREF[j]}" for j in range(8))
    put(sc, rr, 26, f'=IF(P{rr}="Pass",({q})/(5*{SC_}$F$21)*100,"-")', font=f_bold, align=center, fmt="0.0")
    put(sc, rr, 27, f'=IF(AJ{rr}<0,"-",COUNTIFS($Q${SC_FIRST}:$Q${SC_LAST},Q{rr},$AJ${SC_FIRST}:$AJ${SC_LAST},">"&AJ{rr})+1)', align=center)
    put(sc, rr, 28, (f'=IF(P{rr}<>"Pass","Parked (gate)",IF(AC{rr}="PROTECT","Protect (control)",'
                     f'IF(AA{rr}<={QUOTA_F(rr)},"Recommended",IF(AA{rr}<={QUOTA_F(rr)}+{RS_REF},"Reserve","Backlog"))))'), font=f_bold, align=center)
    put(sc, rr, 29, f"=INDEX({HS_}$R${HS_FIRST}:$R${HS_LAST},MATCH($A{rr},{HS_}$A${HS_FIRST}:$A${HS_LAST},0))", align=center)
    put(sc, rr, 30, f"=IF(AND(AB{rr}=\"Recommended\",J{rr}>={SC_}$C$26),\"Yes\",\"\")", align=center)
    t3 = r["t3"].split(" / ") if r["t3"] else ["", ""]
    put(sc, rr, 31, comp_lookup(f"$A{rr}", "Primary 1", t3[0] if r["gate"] else ""))
    put(sc, rr, 32, comp_lookup(f"$A{rr}", "Primary 2", (t3[1] if len(t3) > 1 else "") if r["gate"] else ""))
    put(sc, rr, 33, comp_lookup(f"$A{rr}", "Secondary 1", ""))
    put(sc, rr, 34, comp_lookup(f"$A{rr}", "Secondary 2", ""))
    put(sc, rr, 35, r["why"])
    c = sc.cell(row=rr, column=36, value=f'=IF(AND(P{rr}="Pass",AC{rr}<>"PROTECT"),Z{rr}+U{rr}/100+R{rr}/1000,-1)')
    c.font = Font(name=F, size=8, color="808080"); c.number_format = "0.000"
sc.freeze_panes = "C5"
sc.auto_filter.ref = f"A4:AJ{SC_LAST}"
status_cf(sc, f"AB{SC_FIRST}:AB{SC_LAST}", "AB", SC_FIRST)
seg_cf(sc, f"Q{SC_FIRST}:Q{SC_LAST}", "Q", SC_FIRST)
dv2 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True)
sc.add_data_validation(dv2); dv2.add(f"R{SC_FIRST}:Y{SC_LAST}")
dv3 = DataValidation(type="list", formula1='"Pass,Fail"', showErrorMessage=True)
sc.add_data_validation(dv3); dv3.add(f"M{SC_FIRST}:M{SC_LAST}"); dv3.add(f"O{SC_FIRST}:O{SC_LAST}")

def sc_lookup(col, cid_cell):
    return f"=INDEX('Path Scorecard'!${col}${SC_FIRST}:${col}${SC_LAST},MATCH({cid_cell},'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))"

# --------------------------------------------------------------- Hero Segmentation
hs = sheet("Hero Segmentation", "Hero segmentation: Staples strength vs market attractiveness, per L3/L4 path",
           "Rules and 1-5 anchors are in 'Hero Criteria'. S1-S2 and M4 are computed; S3, S4, M1-M3 are analyst inputs (blue).")
header(hs, 4, ["Path ID", "Staples path (exact)", "Link to your list", "L1 > L2", "Staples items", "L1 in header nav",
               "S1 Depth", "S2 Visibility", "S3 Merch. investment", "S4 Shopper association", "SSI Staples strength (0-100)",
               "M1 Demand momentum", "M2 Competitor destination", "M3 Size & frequency", "M4 Occasion (=C7)", "MAI Market attractiveness (0-100)",
               "Segment", "Marketplace play", "Status", "Evidence / note",
               "Chart helpers (#N/A = point hidden on purpose)", "chart: SSI Hero", "chart: MAI Hero", "chart: SSI Prob. Hero", "chart: MAI Prob. Hero",
               "chart: SSI Non-Hero", "chart: MAI Non-Hero"],
       [8, 56, 22, 34, 9, 9, 7, 8, 9, 9, 10, 9, 10, 9, 9, 10, 13, 11, 14, 60, 14, 9, 9, 9, 9, 9, 9])
for i, r in enumerate(order):
    rr = HS_FIRST + i
    n = r["n"]
    s3, s4, m1, m2, m3 = SM[r["key"]]
    put(hs, rr, 1, pid[r["key"]], align=center); put(hs, rr, 2, r["path"], font=f_bold); put(hs, rr, 3, r["link"])
    put(hs, rr, 4, f'{n["L1"]} > {n["L2"]}'); put(hs, rr, 5, r["items"], align=center, fmt="#,##0"); put(hs, rr, 6, r["nav"], align=center)
    put(hs, rr, 7, f"=IF(E{rr}>={B_REF[0]},5,IF(E{rr}>={B_REF[1]},4,IF(E{rr}>={B_REF[2]},3,IF(E{rr}>={B_REF[3]},2,1))))", align=center)
    put(hs, rr, 8, f'=IF(F{rr}="Yes",{NY_REF},{NN_REF})', align=center)
    put(hs, rr, 9, s3, font=f_input, fill=fill_input, align=center); put(hs, rr, 10, s4, font=f_input, fill=fill_input, align=center)
    put(hs, rr, 11, f"=(G{rr}*{SREF[0]}+H{rr}*{SREF[1]}+I{rr}*{SREF[2]}+J{rr}*{SREF[3]})/(5*{SSI_TOT})*100", font=f_bold, align=center, fmt="0")
    for j, v in enumerate((m1, m2, m3)):
        put(hs, rr, 12 + j, v, font=f_input, fill=fill_input, align=center)
    lk = f"INDEX('Path Scorecard'!${C7_COL}${SC_FIRST}:${C7_COL}${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))"
    put(hs, rr, 15, f'=IF({lk}="",3,{lk})', align=center)
    put(hs, rr, 16, f"=(L{rr}*{MREF[0]}+M{rr}*{MREF[1]}+N{rr}*{MREF[2]}+O{rr}*{MREF[3]})/(5*{MAI_TOT})*100", font=f_bold, align=center, fmt="0")
    put(hs, rr, 17, f'=IF(K{rr}>={H_REF},"Hero",IF(P{rr}>={P_REF},"Probable Hero","Non-Hero"))', font=f_bold, align=center)
    c3 = f"INDEX('Path Scorecard'!${C3_COL}${SC_FIRST}:${C3_COL}${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))"
    put(hs, rr, 18, f'=IF(Q{rr}="Hero",IF({c3}<={PR_REF},"PROTECT","EXTEND"),IF(Q{rr}="Probable Hero","BUILD","EXPLORE"))', font=f_bold, align=center)
    put(hs, rr, 19, f"=INDEX('Path Scorecard'!${C_STATUS}${SC_FIRST}:${C_STATUS}${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))", align=center)
    put(hs, rr, 20, SEG_NOTE.get(r["key"], ""))
    for j, (seg, src) in enumerate([("Hero", "K"), ("Hero", "P"), ("Probable Hero", "K"), ("Probable Hero", "P"), ("Non-Hero", "K"), ("Non-Hero", "P")]):
        c = hs.cell(row=rr, column=22 + j, value=f'=IF(AND($Q{rr}="{seg}",$S{rr}<>"Parked (gate)"),{src}{rr},NA())')
        c.font = Font(name=F, size=8, color="808080"); c.number_format = "0"
seg_cf(hs, f"Q{HS_FIRST}:Q{HS_LAST}", "Q", HS_FIRST)
status_cf(hs, f"S{HS_FIRST}:S{HS_LAST}", "S", HS_FIRST)
dv4 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True)
hs.add_data_validation(dv4); dv4.add(f"I{HS_FIRST}:J{HS_LAST}"); dv4.add(f"L{HS_FIRST}:N{HS_LAST}")
hs.freeze_panes = "C5"
hs.auto_filter.ref = f"A4:T{HS_LAST}"
R2 = HS_LAST + 3
hs.cell(row=R2 - 1, column=1, value="L2 roll-up: how many scored L3/L4 paths of each segment sit under every Staples L2 (live)").font = f_sec
for j, h in enumerate(["", "L1 > L2", "Hero", "Probable Hero", "Non-Hero", "Parked (gate)", "Reading"], 1):
    c = hs.cell(row=R2, column=j, value=h or None)
    if h:
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
l2s = sorted({f'{r["n"]["L1"]} > {r["n"]["L2"]}' for r in rows})
D_ = f"$D${HS_FIRST}:$D${HS_LAST}"; Q_ = f"$Q${HS_FIRST}:$Q${HS_LAST}"; S_ = f"$S${HS_FIRST}:$S${HS_LAST}"
for i, l2 in enumerate(l2s):
    rr = R2 + 1 + i
    put(hs, rr, 2, l2, font=f_bold)
    for j, seg in enumerate(["Hero", "Probable Hero", "Non-Hero"]):
        put(hs, rr, 3 + j, f'=COUNTIFS({D_},$B{rr},{Q_},"{seg}",{S_},"<>Parked (gate)")', align=center)
    put(hs, rr, 6, f'=COUNTIFS({D_},$B{rr},{S_},"Parked (gate)")', align=center)
    put(hs, rr, 7, f'=IF((C{rr}>0)+(D{rr}>0)+(E{rr}>0)>=2,"Mixed: one L2 holds several segments, so decide at L3/L4",IF(C{rr}+D{rr}+E{rr}=0,"Only parked paths","Single segment"))')

# --------------------------------------------------------------- Segment Map (PNG snapshot + live native chart)
pts = [r for r in rows if r["gate"]]
fig, ax = plt.subplots(figsize=(13, 9.5), dpi=150)
fig.patch.set_facecolor("#fcfcfb"); ax.set_facecolor("#fcfcfb")
YMIN = 22
ax.axvspan(H_CUT, 100, color="#2a78d6", alpha=0.05, lw=0)
ax.fill_between([0, H_CUT], P_CUT, 100, color="#eb6834", alpha=0.05, lw=0)
ax.fill_between([0, H_CUT], YMIN, P_CUT, color="#9e9d98", alpha=0.05, lw=0)
ax.axvline(H_CUT, color="#52514e", lw=1, ls=(0, (4, 3)))
ax.plot([0, H_CUT], [P_CUT, P_CUT], color="#52514e", lw=1, ls=(0, (4, 3)))
ax.text(99, 99.5, "HERO\nStaples' own strength: EXTEND (or PROTECT)", ha="right", va="top", fontsize=10, fontweight="bold", color="#1f3864")
ax.text(19, 99.5, "PROBABLE HERO\nMarket favourite, Staples not yet: BUILD", ha="left", va="top", fontsize=10, fontweight="bold", color="#8a3a12")
ax.text(19, YMIN + 1, "NON-HERO\nNeither: EXPLORE (low-risk test)", ha="left", va="bottom", fontsize=10, fontweight="bold", color="#52514e")
style = {"Recommended": dict(s=150, lw=1.6), "Reserve": dict(s=70, lw=1.2)}
for seg in ["Non-Hero", "Probable Hero", "Hero"]:
    for st in ["Backlog", "Protect (control)", "Reserve", "Recommended"]:
        sub = [r for r in pts if r["seg"] == seg and r["status"] == st]
        if not sub:
            continue
        filled = st in ("Recommended", "Reserve")
        ax.scatter([r["ssi"] for r in sub], [r["mai"] for r in sub], s=style.get(st, dict(s=42))["s"],
                   c=SEG_HEX[seg] if filled else "#fcfcfb", edgecolors="#fcfcfb" if filled else SEG_HEX[seg],
                   linewidths=style.get(st, dict(lw=1.2))["lw"], zorder=3)
NAME = lambda k: k[0].upper() + k[1:]
POS = {}   # hand-placed labels: key of the group's first shelf -> (x, y, ha)
POS.update({
 "Pencil cases": (59.2, 89.5, "right"), "Accent chairs": (54.2, 82.6, "right"), "Faux plants": (54.2, 80.6, "right"),
 "Desk lamps": (61.7, 82.3, "left"), "Office desks": (94.3, 77.7, "right"), "Calendars": (95, 73.6, "center"),
 "Outdoor games": (43.3, 58.9, "right"), "Work totes": (40, 63.5, "center"),
 "Keyboards": (76.3, 74.7, "right"), "Laptop sleeves": (63.5, 73.4, "left"), "Journals": (60.6, 86.3, "left"),
})
groups = {}
for r in pts:
    if r["status"] in ("Recommended", "Reserve"):
        groups.setdefault((round(r["ssi"], 1), round(r["mai"], 1)), []).append(r)
for (x, y), g in groups.items():
    rec = sorted([r for r in g if r["status"] == "Recommended"], key=lambda r: r["rank_seg"])
    res = sorted([r for r in g if r["status"] == "Reserve"], key=lambda r: r["rank_seg"])
    if rec:
        tx, ty, ha = POS.get(rec[0]["key"], (x + 0.5, y - 0.35, "left"))
        ax.annotate(" · ".join(NAME(r["key"]) for r in rec), (x, y), xytext=(tx, ty), textcoords="data", ha=ha, va="center",
                    fontsize=9, fontweight="bold", color="#0b0b0b", zorder=4)
    if res:
        dflt = (x + 0.5, y - 2.0, "left") if rec else (x + 0.5, y - 0.35, "left")
        tx, ty, ha = POS.get(res[0]["key"], dflt)
        ax.annotate(" · ".join(NAME(r["key"]) for r in res) + " (reserve)", (x, y), xytext=(tx, ty), textcoords="data", ha=ha, va="center",
                    fontsize=7.6, color="#52514e", zorder=4)
prot = [r for r in pts if r["play"] == "PROTECT"]
for (x, y) in sorted({(round(r["ssi"], 1), round(r["mai"], 1)) for r in prot}):
    names = " · ".join(NAME(r["key"]) for r in prot if (round(r["ssi"], 1), round(r["mai"], 1)) == (x, y))
    edge = x > 95
    ax.annotate(f"{names} (PROTECT)", (x, y), xytext=(x - 0.3 if edge else x + 0.6, y + 1.6 if edge else y - 0.4), textcoords="data",
                ha="right" if edge else "left", va="center", fontsize=7.4, fontstyle="italic", color="#52514e", zorder=4)
ax.set_xlim(18, 102); ax.set_ylim(YMIN, 100)
ax.set_xlabel("Staples Strength Index (SSI): depth, visibility, merchandising investment, shopper association", fontsize=10, color="#52514e")
ax.set_ylabel("Market Attractiveness Index (MAI): momentum, competitor destination, size, occasion", fontsize=10, color="#52514e")
ax.set_title("Top 20 by segment: 7 Hero · 7 Probable Hero · 6 Non-Hero  (bold = Top 20, grey = reserve, hollow = backlog / protect)",
             fontsize=12, fontweight="bold", color="#0b0b0b", loc="left")
for sp in ["top", "right"]:
    ax.spines[sp].set_visible(False)
for sp in ["left", "bottom"]:
    ax.spines[sp].set_color("#c3c2b7")
ax.tick_params(colors="#52514e", labelsize=9)
ax.grid(color="#e5e4df", lw=0.6, zorder=0)
from matplotlib.lines import Line2D
leg = [Line2D([0], [0], marker="o", ls="", markerfacecolor=SEG_HEX[k], markeredgecolor=SEG_HEX[k], markersize=9, label=k) for k in SEG_HEX]
leg += [Line2D([0], [0], marker="o", ls="", markerfacecolor="#52514e", markeredgecolor="#fcfcfb", markersize=11, label="Top 20 (large)"),
        Line2D([0], [0], marker="o", ls="", markerfacecolor="#52514e", markeredgecolor="#fcfcfb", markersize=7, label="Reserve"),
        Line2D([0], [0], marker="o", ls="", markerfacecolor="#fcfcfb", markeredgecolor="#52514e", markersize=6, label="Backlog / Protect (hollow)")]
ax.legend(handles=leg, loc="lower right", fontsize=8.5, frameon=False, ncol=2)
fig.tight_layout()
fig.savefig(FIG, facecolor=fig.get_facecolor())
plt.close(fig)

smap = sheet("Segment Map", "Segment map: Staples strength vs market attractiveness",
             "Picture = snapshot at default inputs (labelled). The Excel chart further down is live: it moves when you edit inputs or cut-offs.")
smap.column_dimensions["A"].width = 3
img = XLImage(str(FIG)); img.width, img.height = 1170, 855
smap.add_image(img, "B4")
notes = [
 "How to read it: right of the dashed line = Hero (Staples is already strong); top-left = Probable Hero (the market loves it, Staples is not yet a destination); bottom-left = Non-Hero.",
 "Plays: Hero = EXTEND (style extensions only) or PROTECT (keep the marketplace out, e.g. Office Chairs, Pens); Probable Hero = BUILD (fill the gap with sellers); Non-Hero = EXPLORE (low-risk test, sellers hold the stock).",
 "The Top 20 takes the best 7 Hero, 7 Probable Hero and 6 Non-Hero shelves by path score, so every segment is represented.",
]
for i, t in enumerate(notes):
    c = smap.cell(row=49 + i, column=2, value=t); c.font, c.alignment = f_body, wrap
    smap.merge_cells(start_row=49 + i, start_column=2, end_row=49 + i, end_column=16)
    smap.row_dimensions[49 + i].height = 30
ch = ScatterChart()
ch.title = "Live segment map (updates with inputs)"
ch.x_axis.title = "Staples Strength Index (SSI)"
ch.y_axis.title = "Market Attractiveness Index (MAI)"
ch.x_axis.scaling.min, ch.x_axis.scaling.max = 15, 100
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 20, 100
ch.x_axis.delete = False; ch.y_axis.delete = False
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
ch.x_axis.majorGridlines = None
ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E5E4DF"))
ch.legend.position = "r"
for t in (ch.title, ch.x_axis.title, ch.y_axis.title):
    t.overlay = False
for j, seg in enumerate(["Hero", "Probable Hero", "Non-Hero"]):
    xs = Reference(hs, min_col=22 + 2 * j, min_row=HS_FIRST, max_row=HS_LAST)
    ys = Reference(hs, min_col=23 + 2 * j, min_row=HS_FIRST, max_row=HS_LAST)
    se = Series(ys, xs, title=seg)
    se.marker.symbol = "circle"; se.marker.size = 8
    se.marker.graphicalProperties.solidFill = SEG_HEX[seg][1:]
    se.marker.graphicalProperties.line.solidFill = "FFFFFF"
    se.graphicalProperties.line.noFill = True
    ch.series.append(se)
ch.height, ch.width = 12, 24
smap.add_chart(ch, "B54")

# --------------------------------------------------------------- Top 20 Recommendations
tp = sheet("Top 20 Recommendations", "Top 20 L3/L4 recommendations: 7 Hero · 7 Probable Hero · 6 Non-Hero, each with 4 competitors",
           "Segment, rank, status, play, score and all four competitors are live lookups. Primary = broad retailer (measures the gap); Secondary = enthusiast / design specialist (the super-extension).")
TP_COLS = ["#", "Hero segment", "Rank in segment", "Status (live)", "Marketplace play", "Path score", "Staples path (exact)", "Level", "Category ID",
           "Staples items", "SKUs in hand", "Fast-start", "Primary 1", "Primary 2", "Secondary 1 (enthusiast)", "Secondary 2 (enthusiast)",
           "What the secondaries show (super-extension)", "Why this path", "Archetypes we expect to find (hypotheses)", "1P watch-out (cannibalisation)",
           "Commercial moment", "Why the room will relate", "Where to look on the primaries", "Staples URL"]
header(tp, 4, TP_COLS, [5, 13, 8, 13, 11, 8, 50, 6, 11, 9, 9, 8, 16, 16, 19, 19, 48, 46, 46, 34, 24, 34, 42, 40])

def top_row(ws, rr, i, r, cid_col="I"):
    put(ws, rr, 1, i, align=center)
    put(ws, rr, 2, sc_lookup(C_SEG, f"${cid_col}{rr}"), font=f_bold, align=center)
    put(ws, rr, 3, sc_lookup(C_RANK, f"${cid_col}{rr}"), align=center)
    put(ws, rr, 4, sc_lookup(C_STATUS, f"${cid_col}{rr}"), font=f_bold, align=center)
    put(ws, rr, 5, sc_lookup(C_PLAY, f"${cid_col}{rr}"), font=f_bold, align=center)
    put(ws, rr, 6, sc_lookup(C_SCORE, f"${cid_col}{rr}"), align=center, fmt="0.0")
    put(ws, rr, 7, r["path"], font=f_bold)
    put(ws, rr, 8, r["level"], align=center); put(ws, rr, 9, r["cid"], align=center)
    put(ws, rr, 10, r["items"], align=center, fmt="#,##0"); put(ws, rr, 11, r["hand"], align=center, fmt="#,##0")
    put(ws, rr, 12, sc_lookup(C_FAST, f"${cid_col}{rr}"), align=center)
    put(ws, rr, 13, sc_lookup(C_P1, f"${cid_col}{rr}"), font=f_bold); put(ws, rr, 14, sc_lookup(C_P2, f"${cid_col}{rr}"))
    put(ws, rr, 15, sc_lookup(C_S1, f"${cid_col}{rr}"), font=f_bold); put(ws, rr, 16, sc_lookup(C_S2, f"${cid_col}{rr}"))
    put(ws, rr, 17, SEC_LENS[r["key"]])
    for j, v in enumerate(D[r["key"]]):
        put(ws, rr, 18 + j, v)
    put(ws, rr, 24, r["url"])
    ws.row_dimensions[rr].height = 92

rr = 5
for i, r in enumerate(top20, 1):
    top_row(tp, rr, i, r); rr += 1
TP_LAST = rr - 1
rr += 1
tp.cell(row=rr, column=1, value="Reserves: the next 3 shelves per segment (swap-ins)").font = f_sec
rr += 1
RES_FIRST = rr
for i, r in enumerate(reserves, 1):
    top_row(tp, rr, f"R{i}", r); rr += 1
RES_LAST = rr - 1
tp.freeze_panes = "H5"
tp.auto_filter.ref = f"A4:X{TP_LAST}"
for a, b in [(5, TP_LAST), (RES_FIRST, RES_LAST)]:
    seg_cf(tp, f"B{a}:B{b}", "B", a)
    status_cf(tp, f"D{a}:D{b}", "D", a)

# --------------------------------------------------------------- Your List Review
yl = sheet("Your List Review", "Review of the initial 9-item list: verdicts, segment mix and exact L3/L4 drill-downs",
           "Table 1 = verdict per item; column I shows most L2s mix Hero, Probable Hero and Non-Hero shelves (why segmenting at L2 misleads). Table 2 = the drill-down paths with live segment, status and competitors.")
review = [
 ("1", "Chairs & Seating", "Wayfair", "Furniture > Chairs & Seating", "KEEP: strongest furniture item",
  "Core-adjacent furniture with 14 leaves. The pilot proved the method here (Accent & Waiting Room, Benches, Bar Stools). The 2026 'resimercial' return-to-office trend supports design-led seating.",
  "Wayfair is right as Primary 1, with Amazon as Primary 2 (seller pool, value tier). Secondaries: Article + West Elm (design-led accent chairs).",
  "P: Wayfair + Amazon | S: Article + West Elm"),
 ("2", "Desks", "Wayfair", "Furniture > Desks", "KEEP, narrow to Office Desks",
  "Office Desks is the real shelf (1,501 items), split by Staples' own type facet. Sit & Stand is spec-led, so design gaps are smaller and it is a PROTECT shelf.",
  "Wayfair + Amazon as primaries; Design Within Reach + West Elm as secondaries.",
  "P: Wayfair + Amazon | S: DWR + West Elm"),
 ("3", "Lamps & Lighting", "Wayfair", "Furniture > Lamps & Lighting", "KEEP (Desk Lamps)",
  "The purest aesthetic category in the core catalogue; Staples SKUs already in hand.",
  "Wayfair + Amazon as primaries; Lamps Plus (specialist depth) + Urban Outfitters (trend lamps) as secondaries.",
  "P: Wayfair + Amazon | S: Lamps Plus + Urban Outfitters"),
 ("4", "Rugs", "Wayfair", "Decor > Rugs (L2 leaf, no L3/L4 under it)", "PARK: fails the granularity gate",
  "There is nothing to drill into: 'Decor > Rugs' is a single L2 leaf of 804 items. Rugs are also residential decor, the weakest link to Staples' work and school mission.",
  "For a decor story at L3/L4, use Storage Baskets (Top 20) or Faux Plants (reserve).", "n/a (parked)"),
 ("5", "Fitness Equipment", "Amazon", "Fitness > Fitness Equipment > Fitness Machines > Fitness Machines (the only leaf)", "REFOCUS as 'active workstation' (backlog)",
  "Fitness is not a Staples destination: one leaf, 43 items, not in the header nav. Defensible only as under-desk walking pads and desk bikes next to sit-stand desks. A Probable Hero, but a low path score.",
  "If run: Amazon + Dick's as primaries; Scheels as the enthusiast secondary (recovery and training gear).", "Backlog"),
 ("6", "Backpacks & Bags", "Scheels", "Bags, Backpacks & Luggage > Backpacks & Laptop Bags (+ Totes & Handbags, Briefcases & Padfolios)", "KEEP, re-cast Scheels as a secondary",
  "Strong BTS anchor (85% of parents buy a backpack, Circana). Scheels is exactly the enthusiast lens, but its depth is technical / hunting packs, so for backpacks Bellroy and Urban Outfitters score higher on adjacency.",
  "Primaries: Amazon + Target. Secondaries: Bellroy (premium commuter) + Urban Outfitters (fashion). Scheels and REI stay as technical-pack alternates.",
  "P: Amazon + Target | S: Bellroy + Urban Outfitters"),
 ("7", "Hydration & Drinkware", "Scheels", "Coffee, Water & Snacks > Water & Beverages > Water Bottles, Tumblers & Travel Mugs", "KEEP: #1 Probable Hero",
  "The Staples node is an L3 under Coffee, Water & Snacks. 463 items, already in the sample. The kitchen 'Drinkware & Glassware' leaf has only 15 items and fails the baseline gate.",
  "Primaries: Dick's + Target. Scheels is kept, now as Secondary 1 (wall-to-wall colourways), with YETI (customisation).",
  "P: Dick's + Target | S: Scheels + YETI"),
 ("8", "Labels & Packaging Accessories", "Michaels", "Office Supplies > Labels (L2 leaf) · Party Supplies > Wrapping Supplies · Retail Store Supplies > Retail Packaging", "REFRAME",
  "'Labels' is an L2 leaf of 11,493 commodity labels (a Hero to protect). Gift wrap and gift bags overlap the Party City partner assortment on Staples.com (Apr-2026). Label Makers and Mailers are Hero reserves.",
  "Mailers: P Amazon + Uline, S noissue + Packhelp (custom branded packaging). Label Makers: P Amazon + Walmart, S Michaels (Cricut) + Container Store.",
  "Reserves: see Table 2"),
 ("9", "Planning & Craft Organization", "Michaels", "Office Supplies > Calendars & Planners · Arts & Crafts > Scrapbooking / Crafting · Office Supplies > Storage & Organization", "KEEP, split in two",
  "Two shopper missions in one bucket. Planning is the #1 Hero (Michaels' Happy Planner ecosystem vs Staples' dated planners). Organization works better as Desk Organizers (Hero) and Storage Baskets (Probable Hero).",
  "Planners: P Amazon + Target, S Erin Condren + Michaels (your pick, kept). Organization: P Target + Amazon / Wayfair, S Grovemade + Container Store.",
  "P: Amazon + Target | S: Erin Condren + Michaels"),
]
header(yl, 4, ["#", "Your item", "Your competitor", "Exact Staples node(s)", "Verdict", "What the research says", "Competitor view (v3: 2 primary + 2 secondary)",
               "Headline competitors", "Segment mix of its L3/L4 paths (live)"], [4, 22, 13, 46, 30, 66, 56, 28, 24])
for i, rv in enumerate(review):
    rr = 5 + i
    for j, v in enumerate(rv, 1):
        put(yl, rr, j, v, font=f_bold if j in (2, 5) else f_body, align=center if j == 1 else wrap)
    link = f"{rv[0]}. {rv[1]}"
    C_ = f"{HS_}$C${HS_FIRST}:$C${HS_LAST}"; Q2 = f"{HS_}$Q${HS_FIRST}:$Q${HS_LAST}"; S2 = f"{HS_}$S${HS_FIRST}:$S${HS_LAST}"
    cnt = lambda seg: f'COUNTIFS({C_},"{link}",{Q2},"{seg}",{S2},"<>Parked (gate)")'
    put(yl, rr, 9, f'="Hero "&{cnt("Hero")}&CHAR(10)&"Probable Hero "&{cnt("Probable Hero")}&CHAR(10)&"Non-Hero "&{cnt("Non-Hero")}'
                   f'&CHAR(10)&"Parked "&COUNTIFS({C_},"{link}",{S2},"Parked (gate)")')
    yl.row_dimensions[rr].height = 96
    color = "C6EFCE" if rv[4].startswith("KEEP") else "FFEB9C" if rv[4].startswith(("REFOCUS", "REFRAME")) else "F8CBAD"
    yl.cell(row=rr, column=5).fill = PatternFill("solid", fgColor=color)
drill = [
 ("1", ["Accent chairs", "Benches", "Bar stools", "Breakroom chairs", "Ottomans", "Office chairs", "Kids seating", "Gaming chairs"]),
 ("2", ["Office desks", "Sit & stand", "Hutches"]),
 ("3", ["Desk lamps", "Table lamps", "Floor lamps"]),
 ("4", ["Rugs"]),
 ("5", ["Fitness machines"]),
 ("6", ["Backpacks", "Lunch", "Laptop sleeves", "Work totes", "Luggage"]),
 ("7", ["Water bottles", "Kitchen drinkware"]),
 ("8", ["Label makers", "Mailers", "Retail boxes", "Labels", "Gift boxes", "Gift wrap"]),
 ("9", ["Planners", "Journals", "Calendars", "Stickers", "Washi tape", "Calendar boards", "Craft storage", "Storage bins"]),
]
start = 5 + len(review) + 2
yl.cell(row=start - 1, column=1, value="Table 2. Drill-down L3/L4 paths for each item (exact nav-tree paths; segment, status and competitors are live)").font = f_sec
hdr = ["#", "Path ID", "Staples path (exact)", "Level", "Category ID", "Staples items", "Hero segment", "Marketplace play", "Status", "Path score",
       "Primary 1 / Primary 2", "Secondary 1 / Secondary 2", "Role in PoC"]
for j, h in enumerate(hdr, 1):
    c = yl.cell(row=start, column=j, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
for col, w in zip("JKLM", [9, 28, 30, 40]):
    yl.column_dimensions[col].width = max(yl.column_dimensions[col].width or 0, w)
role_note = {"Office chairs": "Control shelf (Hero, PROTECT): shows the safety gate rejecting core-1P look-alikes",
             "Rugs": "Parked: no L3/L4 exists", "Luggage": "Parked: leisure travel, off-mission",
             "Kitchen drinkware": "Parked: 15 items, no baseline", "Labels": "Parked: L2 leaf, commodity",
             "Gift boxes": "Parked: Party City partner overlap", "Gift wrap": "Parked: Party City partner overlap"}
STATUS_NOTE = {"Recommended": "Top 20", "Reserve": "Reserve (swap-in)", "Backlog": "Backlog", "Protect (control)": "Control shelf (PROTECT)", "Parked (gate)": "Parked"}
LK2 = lambda col, rr_: f"INDEX('Path Scorecard'!${col}${SC_FIRST}:${col}${SC_LAST},MATCH($E{rr_},'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))"
rr = start + 1
for num, keys in drill:
    for k in keys:
        r = by_key[k]
        put(yl, rr, 1, num, align=center); put(yl, rr, 2, pid[k], align=center)
        put(yl, rr, 3, r["path"], font=f_bold); put(yl, rr, 4, r["level"], align=center)
        put(yl, rr, 5, r["cid"], align=center); put(yl, rr, 6, r["items"], align=center, fmt="#,##0")
        put(yl, rr, 7, sc_lookup(C_SEG, f"$E{rr}"), font=f_bold, align=center)
        put(yl, rr, 8, sc_lookup(C_PLAY, f"$E{rr}"), align=center)
        put(yl, rr, 9, sc_lookup(C_STATUS, f"$E{rr}"), font=f_bold, align=center)
        put(yl, rr, 10, sc_lookup(C_SCORE, f"$E{rr}"), align=center, fmt="0.0")
        put(yl, rr, 11, f'={LK2(C_P1, rr)}&" / "&{LK2(C_P2, rr)}')
        put(yl, rr, 12, f'={LK2(C_S1, rr)}&" / "&{LK2(C_S2, rr)}')
        put(yl, rr, 13, role_note.get(k) or STATUS_NOTE[r["status"]])
        rr += 1
seg_cf(yl, f"G{start + 1}:G{rr - 1}", "G", start + 1)
status_cf(yl, f"I{start + 1}:I{rr - 1}", "I", start + 1)
yl.freeze_panes = "C5"

# --------------------------------------------------------------- Competitor Profiles
cp = sheet("Competitor Profiles", "Competitor profiles: primary pool and secondary (enthusiast / design) pool",
           "Competitor-level scores feed the fit sheets (primary: K4 recruitable, K5 data access; secondary: E3 distance from Staples, E5 transferability). 'Used as' = default picks.")
PROF = {
 # primary pool
 "Wayfair": ("Primary", "Home & commercial furniture / decor pure-play", "Drop-ship model with 11,000+ suppliers and a 14M+ product catalogue (secondary sources); 21.4M active customers in Q1-2026.",
             "Very wide style range at every price tier.", "Already scraped for the pilot. Today: HTTP 429 (rate-limited).", "House brands hide the supplier; recruit the drop-ship suppliers behind them.", "Yes"),
 "Amazon": ("Primary", "Everything marketplace; the B2B rival Staples meets daily", "3P sellers = ~60-61% of paid units (Q1-Q2 2026). Amazon Business: $60B+ annualised gross sales, 11M+ organisations.",
            "Huge variety, noisy quality.", "Bot-protected; mature paid APIs exist (e.g. Rainforest).", "De-duplicate near-identical 3P listings; sample by facet.", "Yes"),
 "Target": ("Primary", "Mass lifestyle retailer; invite-only curated marketplace (Target Plus)", "Target Plus: invite-only, 'curation at scale', dozens of hand-picked brands added summer 2026; JanSport named.",
            "Design-led owned brands (Brightroom, Threshold, Studio McGee) + exclusive colourways.", "JavaScript-rendered; JSON endpoints behind the site.", "Owned brands are not recruitable; Target Plus brands are. Closest peer to Pat's curated model.", "Yes"),
 "Walmart": ("Primary", "Mass retailer with an open marketplace", "Very large third-party marketplace (general knowledge; not sized in this study).",
             "Value-led, limited design range.", "HTTP 200 but 'Robot or human?' bot check.", "Value tier benchmark; weak on aesthetics.", "No"),
 "Best Buy": ("Primary", "Consumer electronics specialist", "Examples doc: keyboards, audio, gaming at enthusiast depth; Mirakl marketplace.",
              "Colour and style options within known tech brands.", "Official developer Products API (key required).", "Higher cannibalisation risk against Staples' tech range.", "Yes"),
 "Dick's Sporting Goods": ("Primary", "National mass sporting goods", "Stocks YETI, HydroJug, Owala, Stanley and CamelBak with a BTS hydration page.",
                           "Full colourways of trending bottle brands.", "HTTP 403.", "Primary pool because it is national mass; Scheels / REI are the enthusiast secondaries.", "Yes"),
 "Uline": ("Primary", "B2B packaging & facilities catalogue", "45,000+ products in stock.", "Utility-first, many colours / sizes.", "HTTP 200: server-rendered HTML.", "Mostly Uline-branded: weak seller pool.", "Yes"),
 # secondary pool
 "Michaels": ("Secondary", "Arts, crafts, planning & framing specialist", "~200k online SKUs + ~1.3M via its Mirakl marketplace; Happy Planner stickers & companion accessories.", "Enthusiast planning / craft depth.", "HTML 200; grids JavaScript-rendered.", "Hobby-pro items can pull off-mission: keep to planning & display.", "Yes"),
 "Scheels": ("Secondary", "Enthusiast sporting destination (regional)", "Wall-to-wall hydration (Examples doc); Osprey & Mystery Ranch packs; premium cornhole & yard games.", "Technical and colourway depth.", "HTTP 403.", "Technical skew: strongest for hydration and yard games.", "Yes"),
 "REI": ("Secondary", "Outdoor co-op specialist", "Water bottles from YETI, Nalgene, Hydro Flask, Klean Kanteen, BruMate, HydraPak.", "Technical outdoor.", "Not tested.", "Outdoor skew.", "Yes"),
 "JetPens": ("Secondary", "Online Japanese / European stationery specialist", "10,000+ products.", "Design-led stationery benchmark.", "HTTP 403 (Cloudflare).", "Enthusiast skew; sample manually.", "Yes"),
 "The Container Store": ("Secondary", "Organization specialist", "Coordinated acrylic / marble desktop collections; 2026 reset liquidating ~30% of select SKUs.", "Organization style benchmark.", "HTTP 307 to a bot check.", "Shrinking range in 2026; much own brand (Elfa).", "Yes"),
 "Erin Condren": ("Secondary", "D2C planner brand (also at Target, Barnes & Noble)", "LifePlanner collection; personalisation; own stores in Austin and Irvine.", "Aesthetic, customisable planners.", "Not tested.", "D2C brand = direct recruitment target.", "Yes"),
 "Grovemade": ("Secondary", "D2C premium desk-accessory maker (USA)", "Walnut / cork desk shelves, organizers, vegetable-tanned leather desk pads (10 colours, 6 sizes).", "Premium natural-material desk systems.", "Not tested.", "Premium price; D2C = recruitment target.", "Yes"),
 "Orbitkey": ("Secondary", "D2C accessories brand", "Desk mat in vegan leather / recycled PET felt with magnetic cable holder and hidden compartment.", "Functional premium desk mats.", "Not tested.", "D2C = recruitment target.", "Yes"),
 "Pottery Barn Teen": ("Secondary", "Williams-Sonoma Inc. teen furniture & decor", "Pinboards (scalloped, fabric, statement), calendars, desk accessories.", "Teen / dorm aesthetic.", "Not tested.", "Own brand only: style lens, not sellers.", "Yes"),
 "Pottery Barn Kids": ("Secondary", "Williams-Sonoma Inc. kids", "Mackenzie lunch boxes, bento bundles, personalisation.", "Personalised, coordinated kids' gear.", "Not tested.", "Own brand only.", "Yes"),
 "Pottery Barn": ("Secondary", "Williams-Sonoma Inc. home", "Woven baskets in rattan, seagrass and natural fibres; basket collections.", "Premium natural storage.", "Not tested.", "Own brand only.", "Yes"),
 "West Elm": ("Secondary", "Williams-Sonoma Inc. modern home", "Mid-century accent chairs ($300-500 band), wall clocks, home-office desks.", "Mid-century / contemporary design.", "Not tested.", "Own brand only.", "Yes"),
 "CB2": ("Secondary", "Crate & Barrel's modern line", "Modern furniture and decor (general knowledge).", "Modern, sculptural.", "Not tested.", "Own brand only.", "No"),
 "Article": ("Secondary", "D2C modern furniture", "Mid-century accent chairs comparable to West Elm at $300-500.", "Mid-century modern.", "Not tested.", "Sells own designs only.", "Yes"),
 "Design Within Reach": ("Secondary", "Design retailer (MillerKnoll subsidiary)", "Herman Miller desks (Renew, Mode), Eames and Nelson designs.", "Iconic modern design.", "Not tested.", "Group brands: benchmark, unlikely sellers.", "Yes"),
 "Anthropologie": ("Secondary", "Lifestyle retailer (URBN)", "Woven decorative baskets; carries third-party brands such as Vitruvi.", "Bohemian / artisanal decor.", "Not tested.", "Mix of own and third-party brands.", "Yes"),
 "Lamps Plus": ("Secondary", "Largest US specialty lighting retailer", "90,000+ lighting designs; family-owned for 40+ years.", "Specialist lighting depth.", "Not tested.", "Carries many lighting brands (transferable).", "Yes"),
 "Schoolhouse": ("Secondary", "D2C heirloom home brand", "Wall and table clocks collection.", "Heirloom, craft design.", "Not tested.", "Own brand; premium.", "Yes"),
 "Keychron": ("Secondary", "D2C mechanical-keyboard brand", "40+ keyboard models incl. retro designs; widely reviewed as a top enthusiast brand.", "Retro, customisable keyboards.", "Not tested.", "D2C = recruitment target.", "Yes"),
 "Drop": ("Secondary", "Enthusiast keyboard community & store", "Custom keyboards incl. licensed collabs; also sold at Micro Center.", "Custom / collectible keyboards.", "Not tested.", "Enthusiast pricing.", "Yes"),
 "Micro Center": ("Secondary", "Enthusiast computer retailer", "Carries Drop and other enthusiast keyboards.", "PC-builder culture.", "Not tested.", "Alternate only.", "Yes"),
 "Kinokuniya": ("Secondary", "Japanese bookstore & stationery (US stores + web)", "102 results in Pen Cases; PuniLabo, Lihit Lab, character collabs.", "Japanese stationery culture.", "Not tested.", "Imported brands via distributors.", "Yes"),
 "Oriental Trading": ("Secondary", "Party, craft & classroom catalogue", "A major classroom-decor and teacher catalogue with full themed kits.", "Themed classroom sets.", "HTTP 403.", "Mostly own-sourced.", "Yes"),
 "Really Good Stuff": ("Secondary", "Teacher-designed classroom supplies", "More than a dozen classroom themes (Boho, Calm, Modern Farmhouse).", "Teacher-first themes.", "Not tested.", "Alternate.", "Yes"),
 "Etsy": ("Secondary", "Handmade / independent marketplace", "Letter boards, felt panels, custom cornhole, classroom theme sets (general knowledge).", "Handmade, personalised.", "Not tested.", "Small makers may not scale: curate carefully.", "No"),
 "Bellroy": ("Secondary", "D2C carry brand (B Corp, Australia)", "Laptop sleeves with magnetic closure; work bags; also on Amazon.", "Premium minimal carry.", "Not tested.", "D2C = recruitment target.", "Yes"),
 "Urban Outfitters": ("Secondary", "Youth lifestyle retailer (URBN)", "Fashion backpacks, lamps, tech accessories (general knowledge).", "Trend-led, Gen-Z.", "Not tested.", "Mix of own and third-party brands.", "No"),
 "Felt Right": ("Secondary", "D2C acoustic felt brand", "PET-felt desk dividers (NRC .35, ~50% recycled), room dividers and tiles.", "Acoustics as a design object.", "Not tested.", "D2C = recruitment target.", "Yes"),
 "Williams Sonoma": ("Secondary", "Premium kitchen retailer", "Coffee-bar accessories, Hold Everything pod drawers and canisters.", "Premium kitchen / coffee culture.", "Not tested.", "Mostly own / exclusive brands.", "Yes"),
 "Crate & Barrel": ("Secondary", "Home retailer", "Coffee storage and kitchen organisation (general knowledge).", "Modern home.", "Not tested.", "Own brand mostly.", "No"),
 "Fellow": ("Secondary", "D2C coffee-gear brand", "Stagg EKG gooseneck kettle: a design benchmark; also sold at Costco and Amazon.", "Coffee-enthusiast design.", "Not tested.", "D2C = recruitment target.", "Yes"),
 "Vitruvi": ("Secondary", "D2C aromatherapy brand", "Stone diffusers; stocked at Nordstrom, Anthropologie, Madewell (300+ stores).", "Diffusers as decor.", "Not tested.", "D2C = recruitment target.", "Yes"),
 "Bentgo": ("Secondary", "D2C bento lunch-box brand", "Bento boxes for kids and adults; BTS range.", "Bento system.", "Not tested.", "D2C = recruitment target (also at Target / Amazon).", "Yes"),
 "YETI": ("Secondary", "Premium drinkware / outdoor brand", "Rambler bottles at REI; D2C customisation and colour drops.", "Premium colour drops.", "Not tested.", "Brand = recruitment target.", "Yes"),
 "Nordstrom": ("Secondary", "Premium department store", "Carries Vitruvi and premium work bags.", "Premium fashion / lifestyle.", "Not tested.", "Multi-brand.", "Yes"),
 "Paper Source": ("Secondary", "Stationery & gift retailer (Barnes & Noble)", "Planners, cards, gift wrap (general knowledge).", "Design stationery.", "Plain request failed (000).", "Alternate.", "No"),
 "noissue": ("Secondary", "Custom sustainable packaging platform", "Custom mailers, boxes, tissue; low minimums; compostable / recycled.", "Branded, eco unboxing.", "Not tested.", "Platform = recruitment target.", "Yes"),
 "Packhelp": ("Secondary", "Custom packaging platform", "Custom mailer boxes from 30 pieces; 3D designer; free US delivery.", "Branded packaging.", "Not tested.", "Platform = recruitment target.", "Yes"),
}
assert set(CK) | set(CE) <= set(PROF), (set(CK) | set(CE)) - set(PROF)
used = {}
for r in top20 + reserves:
    k = r["key"]
    for role, comp in zip(["P1", "P2"], roles_p[k]):
        used.setdefault(comp, []).append(f"{role} {NAME(k)}")
    for role, comp in zip(["S1", "S2"], roles_s[k]):
        used.setdefault(comp, []).append(f"{role} {NAME(k)}")
header(cp, 4, ["Competitor", "Pool", "Type", "Scale / depth evidence", "Aesthetic lens", "Recruitability (K4 or E5)", "Data access K5 / Distance E3",
               "Access test / note", "Used as (default picks: Top 20 + reserves)", "Cautions", "Verified in this study's searches?"],
       [24, 10, 30, 52, 30, 12, 12, 32, 48, 40, 12])
for i, (name, (pool, typ, ev_, lens, acc, caut, ver)) in enumerate(sorted(PROF.items(), key=lambda kv: (kv[1][0] != "Primary", kv[0]))):
    rr = 5 + i
    a, b = (CK[name][0], CK[name][1]) if pool == "Primary" else (CE[name][1], CE[name][0])
    vals = [name, pool, typ, ev_, lens, a, b, acc, "; ".join(used.get(name, [])) or "Alternate only", caut, ver]
    for j, v in enumerate(vals, 1):
        put(cp, rr, j, v, font=f_bold if j == 1 else f_body, align=center if j in (2, 6, 7, 11) else wrap)
    cp.cell(row=rr, column=2).fill = PatternFill("solid", fgColor="C6EFCE" if pool == "Primary" else "FCE4D6")
    cp.row_dimensions[rr].height = 60
rr = 5 + len(PROF)
for j, v in enumerate(["Office Depot / ODP", "Excluded", "Direct like-for-like competitor", "Assortment mirrors Staples' utility range.", "Utility-first",
                       "-", "-", "Not tested.", "Not used", "Little gap signal: like-for-like comparison finds few design-led archetypes.", "-"], 1):
    put(cp, rr, j, v, font=f_bold if j == 1 else f_body, align=center if j in (2, 6, 7, 11) else wrap)
cp.freeze_panes = "B5"
cp.auto_filter.ref = f"A4:K{rr}"

# --------------------------------------------------------------- Parked & Excluded
pk = sheet("Parked & Excluded", "Paths considered and not in the Top 20 or reserves (with the reason)",
           "Parked = failed a gate. Protect = Hero control shelves (marketplace kept out). Backlog = scored, but below the segment quota + reserves.")
header(pk, 4, ["Path ID", "Staples path (exact)", "Category ID", "Staples items", "Hero segment", "Marketplace play", "Status", "Path score",
               "Reason", "Revisit when"], [8, 64, 11, 10, 13, 11, 16, 8, 64, 40])
revisit = {"Rugs": "A pseudo-L3 split by rug type is accepted as an exception", "Chair mats": "Same: only with a pseudo-L3 split",
           "Labels": "Never as a design play; commodity", "Gift boxes": "The Party City partnership scope changes",
           "Gift wrap": "The Party City partnership scope changes", "Kitchen drinkware": "Staples adds a baseline (mugs)",
           "Gift stationery": "Pseudo-L3 split", "Patio furniture": "Never; mis-shelved node", "Luggage": "Travel becomes a marketplace theme",
           "Office chairs": "Keep as the control shelf in the PoC", "Pens": "Only for a separate premium-pen (JetPens) study",
           "Whiteboards": "Glass boards only, if 1P risk is managed", "Sit & stand": "Paired with walking pads as a 'wellness desk' story",
           "Headphones": "BTS tech attach is prioritised", "Fitness machines": "Paired with Sit & Stand Desks",
           "Massage tools": "A wellness-at-work theme is added", "Drink mixes": "Food compliance for 3P sellers is solved"}
rr = 5
for grp in (["Parked (gate)"], ["Protect (control)"], ["Backlog"]):
    for r in [r for r in order if r["status"] in grp]:
        put(pk, rr, 1, pid[r["key"]], align=center); put(pk, rr, 2, r["path"]); put(pk, rr, 3, r["cid"], align=center)
        put(pk, rr, 4, r["items"], align=center, fmt="#,##0")
        put(pk, rr, 5, sc_lookup(C_SEG, f"$C{rr}"), font=f_bold, align=center)
        put(pk, rr, 6, sc_lookup(C_PLAY, f"$C{rr}"), align=center)
        put(pk, rr, 7, sc_lookup(C_STATUS, f"$C{rr}"), font=f_bold, align=center)
        put(pk, rr, 8, sc_lookup(C_SCORE, f"$C{rr}"), align=center, fmt="0.0")
        put(pk, rr, 9, r["why"]); put(pk, rr, 10, revisit.get(r["key"], "A Top-20 shelf is dropped, or weights / quotas change"))
        rr += 1
put(pk, rr, 1, "-", align=center); put(pk, rr, 2, "Furniture > Decor > Wall Art > Wall Art/Decor (and sister Wall Art leaves)")
put(pk, rr, 3, "CL140795", align=center); put(pk, rr, 4, int(tree.loc["CL140795", "Count"]), align=center, fmt="#,##0")
for j in (5, 6, 8):
    put(pk, rr, j, "-", align=center)
put(pk, rr, 7, "Not scored", font=f_bold, align=center)
put(pk, rr, 9, "Staples already lists 10,000+ wall-art items (drop-ship catalogue): the gap is already closed."); put(pk, rr, 10, "Never")
status_cf(pk, f"G5:G{rr}", "G", 5)
seg_cf(pk, f"E5:E{rr}", "E", 5)
pk.freeze_panes = "C5"

# --------------------------------------------------------------- Evidence & Sources
ev = sheet("Evidence & Sources", "Evidence behind the scores",
           "Reliability: High = company / official release; Medium = trade press or analyst; Low = vendor blog or market-research PR. 'Opened' = page read directly, not only the search summary.")
evidence = [
 ("Staples BTS 2026: Recess Club & Jellygram exclusive pouches / notebooks; Sincerely Jules planners; Vera Bradley; 750+ items under $10; Party City in-store", "Staples press release", "https://www.businesswire.com/news/home/20260706408220/en/Staples-Holds-Last-Years-Prices-on-Back-to-School-Essentials-as-The-Most-Wonderful-Time-of-the-Year-Returns", "2026-07-06", "High", "No (403); search summary", "Pencil cases, planners, familiarity"),
 ("Blue Sky x Sincerely Jules: 27-piece academic collection sold at Staples", "GlobeNewswire", "https://www.globenewswire.com/news-release/2026/06/01/3304284/0/en/Blue-Sky-and-Sincerely-Jules-Bring-Vintage-Style-to-Academic-Planners-at-Staples.html", "2026-06-01", "High", "No; search summary", "Planners"),
 ("Staples x Party City: shop-in-shops in 700+ stores and on Staples.com (décor, tableware, gift bags, favors)", "Business Wire", "https://www.businesswire.com/news/home/20260421927178/en/Staples-and-Party-City-Announce-Strategic-Partnership-to-Make-Celebrations-Easy", "2026-04-21", "High", "No; search summary", "Gate G4 (gift wrap / gift bags)"),
 ("85% of parents bought a backpack; lifestyle backpacks +3% $ (2025); stereo headphones +6% $; writing instruments & art paper +2% $", "Circana", "https://www.circana.com/post/best-in-class-what-circana-s-insights-reveal-about-2026-back-to-school-trends", "2026", "Medium-High", "Yes", "Backpacks, headphones"),
 ("Teachers spent $895 out of pocket on average (2024-25); median school supply budget $200", "AdoptAClassroom.org", "https://www.adoptaclassroom.org/2025/06/09/2025-teacher-survey-spending-stats-classroom-needs/", "2025-06-09", "Medium", "No; search summary", "Classroom decor"),
 ("Target Plus: invite-only, 'curation at scale', dozens of new hand-picked brands; JanSport named", "Target corporate", "https://corporate.target.com/news-features/article/2026/07/target-plus-growth", "2026-07-01", "High", "Yes", "Competitor choice; curated-marketplace peer"),
 ("3P sellers = 60-61% of Amazon paid units (Q1-Q2 2026)", "Marketplace Pulse", "https://www.marketplacepulse.com/stats/amazon-percent-of-units-by-third-party-sellers", "2026", "Medium-High", "No; search summary", "Amazon K4"),
 ("Amazon Business: $60B+ annualised gross sales; 11M+ organisations", "Amazon", "https://www.aboutamazon.com/news/company-news/what-is-amazon-business", "2026", "High", "No; search summary", "Amazon profile"),
 ("Wayfair: 11,000+ suppliers, 14M+ products, drop-ship model", "Secondary (dropship guides); sell.wayfair.com", "https://sell.wayfair.com/", "2025-26", "Medium", "No; search summary", "Wayfair K4"),
 ("Wayfair Q1-2026: 21.4M active customers, return to customer growth", "Wayfair IR", "https://investor.wayfair.com/news/news-details/2026/Wayfair-Announces-First-Quarter-2026-Results-Reports-Strong-Share-Capture-and-a-Return-to-Active-Customer-Growth/default.aspx", "2026", "High", "No; search summary", "Wayfair profile"),
 ("Michaels: ~200k online SKUs, +1.3M via marketplace", "eMarketer", "https://www.emarketer.com/content/lessons-macy-s-michaels-h-m-launching-third-party-marketplace", "2023-24", "Medium", "No; search summary", "Michaels K1"),
 ("Container Store 2026 reset: liquidating ~30% of select categories / SKUs; BB&B co-branded stores", "Beyond / Business Wire; Chain Store Age", "https://www.businesswire.com/news/home/20260423281194/en/The-Container-Store-Launches-Nationwide-Overhaul-Across-98-Stores", "2026-04-23", "High", "No; search summary", "Container Store caution"),
 ("Uline stocks 45,000+ products", "uline.com", "https://www.uline.com/", "2026", "High", "No; search summary", "Uline K1"),
 ("JetPens carries 10,000+ products", "jetpens.com", "https://www.jetpens.com/AboutUs", "2026", "High", "No; search summary", "JetPens profile"),
 ("Dick's hydration brands: YETI, HydroJug, Owala, Stanley, CamelBak; dedicated BTS hydration page", "dickssportinggoods.com", "https://www.dickssportinggoods.com/f/drinkware", "2026", "Medium", "No; search summary", "Water bottles"),
 ("Owala FreeSip is the top design in 2026; HydroJug Traveler reached #1 in Amazon Tumblers early 2026", "The Kitchn / Ad Age (search summary)", "https://adage.com/article/marketing-news-strategy/how-stanley-owala-and-rivals-are-battling-water-bottle-sales/2542271/", "2026", "Medium", "No; search summary", "Water bottles"),
 ("'Resimercial' is the default office design expectation in 2026 (soft seating, warm wood, rugs)", "AFR Furniture Rental blog; Conklin", "https://www.rentfurniture.com/blog/2026/03/12/why-resimercial-office-design-still-wins-the-modern-workplace/", "2026-03-12", "Low-Medium", "No; search summary", "Accent chairs, desks, plants"),
 ("2026 home-office trends: warm wood, pastels / sage, better lighting; plants the most popular personalisation", "Decorilla / Homedit", "https://www.decorilla.com/online-decorating/home-office-trends-2026", "2026", "Low", "No; search summary", "Desk lamps, faux plants, desk organizers"),
 ("Retro / typewriter and pastel keyboards mainstream; Logitech POP ICON combo sold out at $40 (Feb-2026)", "9to5Toys and review sites", "https://9to5toys.com/2026/02/04/logitech-pop-icon-keyboard-mouse-combo-down-to-40/", "2026-02-04", "Low-Medium", "No; search summary", "Keyboards"),
 ("2026 classroom decor: boho, pastel, calm / SEL themes", "Teacher blogs (Differentiation Corner)", "https://www.differentiationcorner.com/2026/03/13/10-calming-classroom-themes-decor-teachers-are-loving-right-now/", "2026-03-13", "Low", "No; search summary", "Classroom decor"),
 ("Bento-style lunch boxes: search interest peaks Jul-Aug; adult work lunches a major use", "Blog + market sizing", "https://www.mombloglife.com/bento-style-lunchboxes-back-to-school-2026/", "2026", "Low", "No; search summary", "Lunch"),
 ("Mini Bluetooth thermal label printers viral for pantry, classroom and TikTok-Shop sellers", "Label vendor blog / TikTok", "https://mcauleylabels.com/blogs/articles/label-printer-for-tiktok-shop-sellers", "2026", "Low", "No; search summary", "Label makers"),
 ("Custom / recycled poly mailers and 'unboxing' packaging trend for small sellers", "Packaging vendor blogs", "https://theboxology.us/blog/top-mailer-box-packaging-trends-2026/", "2026", "Low", "No; search summary", "Retail boxes, mailers"),
 ("Walking-pad market: North America ~56% share (2025); hybrid work and corporate wellness drivers", "Straits Research", "https://straitsresearch.com/report/walking-pad-market", "2025", "Low", "No; search summary", "Fitness machines"),
 ("Acoustic felt panels used as a design element in 2026 offices", "Acoustics vendor blog", "https://www.easyfelt.co.uk/blogs/interior-trends-2026-acoustics-as-a-design-feature", "2026", "Low", "No; search summary", "Partitions, privacy panels"),
 ("Oversized, minimalist and sculptural wall clocks trend in 2026 home-office decor", "Clock vendor blogs", "https://www.sangnihome.com/2026/06/05/unique-wall-clocks-for-modern-homes-2026-design-trends/", "2026", "Low", "No; search summary", "Clocks (Non-Hero)"),
 ("Room dividers: 'adjustable room divider' and 'folding screen' searches peaked Feb-2026; folding screens back on trend", "Homes & Gardens; market blogs", "https://www.homesandgardens.com/news/folding-screen-room-divider-trend", "2026", "Low-Medium", "No; search summary", "Partitions (Non-Hero)"),
 ("Coffee-station organisation: carousels, drawer organisers and styled coffee-bar stations", "Review blog; Staples node", "https://clutterscience.com/blog/best-coffee-station-organizers/", "2026", "Low", "No; search summary", "Coffee organizers (Non-Hero)"),
 ("Grovemade leather desk pads: vegetable-tanned US leather on cork; 10 colours, 6 sizes; walnut desk shelf systems", "grovemade.com; Gear Patrol", "https://grovemade.com/organizers/", "2026", "High (product facts)", "No; search summary", "Desk pads, desk organizers (secondary)"),
 ("Orbitkey desk mat: vegan leather + recycled PET felt, magnetic cable holder, hidden compartment", "orbitkey.com", "https://www.orbitkey.com/products/desk-mat", "2026", "High (product facts)", "No; search summary", "Desk pads (secondary)"),
 ("Erin Condren planners sold at most Target stores and Barnes & Noble; own stores in Austin and Irvine", "erincondren.com / target.com", "https://www.erincondren.com/planners", "2026", "High", "No; search summary", "Planners, calendars (secondary)"),
 ("Kinokuniya US webstore: 102 results in Pen Cases (PuniLabo, Lihit Lab, character collabs)", "kinokuniya.com", "https://united-states.kinokuniya.com/t/hobby-and-goods/stationery/pencil-cases", "2026", "High", "No; search summary", "Pencil cases, journals (secondary)"),
 ("Pottery Barn Teen pinboards: scalloped, fabric, oversized, no-nail; study-wall boards", "pbteen.com", "https://www.pbteen.com/shop/accessories/calendars-pinboards/", "2026", "High", "No; search summary", "Bulletin boards (secondary)"),
 ("Keychron: 40+ keyboard models, retro designs; widely rated a top enthusiast brand", "keychron.com", "https://www.keychron.com/pages/about-us", "2026", "Medium-High", "No; search summary", "Keyboards (secondary)"),
 ("Drop enthusiast keyboards, also sold at Micro Center", "microcenter.com / drop.com", "https://drop.com/mechanical-keyboards/s", "2026", "Medium", "No; search summary", "Keyboards (secondary)"),
 ("REI water bottles: YETI, Nalgene, Hydro Flask, Klean Kanteen, BruMate, HydraPak", "rei.com", "https://www.rei.com/c/water-bottles", "2026", "High", "No; search summary", "Water bottles (alternate)"),
 ("Scheels backpacks (Osprey, Mystery Ranch, The North Face) and yard games (premium cornhole, Spikeball, giant games)", "scheels.com", "https://www.scheels.com/c/yard-games", "2026", "High", "No; search summary", "Backpacks, outdoor games, water bottles (secondary)"),
 ("Really Good Stuff: a dozen-plus classroom themes incl. Boho and Cool & Calm", "reallygoodstuff.com", "https://www.reallygoodstuff.com/classroom-themes", "2026", "High", "No; search summary", "Classroom decor (alternate)"),
 ("Lamps Plus: nation's largest specialty lighting retailer; 90,000+ lighting designs", "lampsplus.com (search summary)", "https://www.lampsplus.com/lamps/", "2026", "Medium", "No; search summary", "Desk lamps (secondary)"),
 ("Schoolhouse and West Elm run dedicated wall-clock collections", "schoolhouse.com; westelm.com", "https://schoolhouse.com/collections/clocks", "2026", "High", "No; search summary", "Clocks (secondary)"),
 ("Felt Right acoustic desk dividers: PET felt, NRC .35, ~50% recycled content, pinnable", "feltright.com", "https://feltright.com/collections/dividers", "2026", "High (product facts)", "No; search summary", "Privacy panels, partitions (secondary)"),
 ("Pottery Barn Kids Mackenzie lunch boxes with bento bundles and personalisation", "potterybarnkids.com", "https://www.potterybarnkids.com/shopping/kids-bento-lunch-box/", "2026", "High", "No; search summary", "Lunch (secondary)"),
 ("Bellroy: B Corp carry brand; magnetic-closure laptop sleeve; sold D2C and on Amazon", "bellroy.com", "https://bellroy.com/products/laptop-sleeve", "2026", "High", "No; search summary", "Laptop sleeves, backpacks, work totes (secondary)"),
 ("Pottery Barn and Anthropologie woven basket collections (rattan, seagrass, banana bark)", "potterybarn.com; anthropologie.com", "https://www.potterybarn.com/shop/organization/baskets-organization/", "2026", "High", "No; search summary", "Storage baskets (secondary)"),
 ("Williams Sonoma coffee-bar accessories: ash-wood coffee station with pod drawer, Hold Everything canisters", "williams-sonoma.com", "https://www.williams-sonoma.com/shopping/coffee-bar-accessories/", "2026", "High", "No; search summary", "Coffee organizers, kettles (secondary)"),
 ("Fellow Stagg EKG: design-benchmark gooseneck kettle (also at Costco / Amazon)", "fellowproducts.com; reviews", "https://fellowproducts.com/collections/electric-kettles", "2026", "Medium-High", "No; search summary", "Kettles (secondary)"),
 ("Vitruvi diffusers stocked at Nordstrom, Anthropologie, Madewell (300+ stores)", "Forbes; anthropologie.com", "https://www.anthropologie.com/brands/vitruvi", "2018-2026", "Medium", "No; search summary", "Diffusers (secondary)"),
 ("noissue and Packhelp: custom branded, sustainable mailers with low minimums (Packhelp from 30 pieces)", "noissue.co; packhelp.com", "https://packhelp.com/p/custom-mailer-box/custom/", "2026", "High (product facts)", "No; search summary", "Mailers (secondary)"),
 ("Design Within Reach (MillerKnoll subsidiary) sells Herman Miller home-office desks (Renew, Mode)", "dwr.com", "https://www.dwr.com/furniture-office?lang=en_US", "2026", "High", "No; search summary", "Office desks (secondary)"),
 ("Mid-century accent chairs cluster at $300-500 at West Elm and Article", "Mid-century blog comparison", "https://midinmod.com/blogs/mid-century-modern-design-blog/best-online-furniture-stores-in-2026-honest-brand-comparison", "2026", "Low", "No; search summary", "Accent chairs (secondary)"),
 ("Staples navigation tree: 1,877 nodes, 1,240 terminal, item counts per leaf (21-Sep-2026 snapshot)", "Internal", "Excels/1. Staples_Navigation_Tree_repaired.xlsx", "2026-09-21", "High", "Yes", "Every path, G1, G3"),
 ("Staples product sample: 26,861 SKUs (leaf coverage = 'SKUs in hand')", "Internal", "data/interim/sku_staples.parquet", "2026-09-27", "High", "Yes", "C8 data readiness"),
 ("Pilot results per shelf (recommendations, Strong tier, mapping confidence, Spearman)", "Internal PoC run", "outputs/tables/shelf_selection.csv", "2026-09-28", "High", "Yes", "Chairs & desks scores"),
 ("Plain-HTTP access test of 15 competitor sites (status codes in 'Competitor Profiles')", "Own test", "curl, 29-Sep-2026", "2026-09-29", "High", "Yes", "K5 data access"),
]
header(ev, 4, ["#", "Claim / figure", "Source", "URL / location", "Date", "Reliability", "Opened?", "Used for"], [4, 70, 24, 60, 11, 11, 18, 30])
for i, e in enumerate(evidence):
    rr = 5 + i
    put(ev, rr, 1, i + 1, align=center)
    for j, v in enumerate(e, 2):
        put(ev, rr, j, v, align=center if j in (5, 6) else wrap)
    ev.row_dimensions[rr].height = 44
ev.freeze_panes = "C5"

# --------------------------------------------------------------- Staples Tree Extract
te = sheet("Staples Tree Extract", "Staples navigation tree: all terminal nodes (reference)",
           "Copied from Excels/1. Staples_Navigation_Tree_repaired.xlsx (snapshot 21-Sep-2026). Filter L1/L2 to verify any path.")
term = tree[tree["Is Terminal"] == "Yes"].sort_values(["L1", "L2", "L3", "L4"], na_position="first")
header(te, 4, ["L1", "L2", "L3", "L4", "Level", "Category ID", "Items", "SKUs in hand", "Path ID (this study)", "Full path", "URL"],
       [26, 30, 30, 34, 6, 11, 8, 9, 10, 70, 60])
ids = {r["cid"]: pid[r["key"]] for r in rows}
for i, (_, n) in enumerate(term.iterrows()):
    rr = 5 + i
    vals = [n["L1"], n["L2"] if pd.notna(n["L2"]) else "", n["L3"] if pd.notna(n["L3"]) else "", n["L4"] if pd.notna(n["L4"]) else "",
            int(n["Level"]), n["Category ID"], int(n["Count"]) if pd.notna(n["Count"]) else None,
            int(in_hand.get(n["path"], 0)), ids.get(n["Category ID"], ""), n["path"], n["URL"]]
    for j, v in enumerate(vals, 1):
        c = te.cell(row=rr, column=j, value=v)
        c.font = f_bold if (j == 9 and v) else f_body
te.freeze_panes = "A5"
te.auto_filter.ref = f"A4:K{4 + len(term)}"

# --------------------------------------------------------------- Executive Summary
ex = wb.create_sheet("Executive Summary", 0)
ex.sheet_view.showGridLines = False
ex["A1"] = "Staples Marketplace PoC: the Top 20 L3/L4 shelves, and who to benchmark them against"
ex["A1"].font = f_title
ex["A2"] = "v3, prepared 29-Sep-2026 for Sai (LatentView). Paths are exact nodes from the Staples navigation tree (snapshot 21-Sep-2026)."
ex["A2"].font = f_sub
for col, w in zip("ABCDEFGHIJK", [5, 20, 52, 8, 15, 15, 18, 18, 44, 26, 9]):
    ex.column_dimensions[col].width = w
fast = [r for r in top20 if r["hand"] >= FAST_MIN]
d2c = sorted({c for r in top20 for c in roles_s[r["key"]] if CE[c][1] == 5})
msgs = [
 f"Recommendation: 20 L3/L4 shelves: {QUOTA['Hero']} Hero, {QUOTA['Probable Hero']} Probable Hero and {QUOTA['Non-Hero']} Non-Hero. Each is benchmarked against 4 competitors: 2 primaries (broad retailers that measure the gap) and 2 secondaries (enthusiast / design specialists outside Staples' scope that show the 'super-extension', like Scheels and Michaels in Ben's examples).",
 "Hero = Staples' own top categories (EXTEND with style variants; PROTECT where variants would copy 1P). Probable Hero = the market's top categories where Staples is not yet a destination (BUILD). Non-Hero = neither (EXPLORE: a low-risk test where sellers hold the stock). Full rules and 1-5 anchors are in 'Hero Criteria'.",
 "How the 20 are chosen: 4 hard gates, then the Hero segment, then an 8-criterion path score. The best shelves inside each segment fill the quota. Quotas, cut-offs and weights are editable, and everything recalculates.",
 f"Secondaries double as a recruitment list: {len(d2c)} of the Top-20 secondary picks are D2C brands that could list on Staples Marketplace themselves ({', '.join(d2c)}).",
 f"{len(fast)} Top-20 shelves already have Staples SKUs in hand ('Fast-start'); the rest need a Staples leaf scrape first.",
 "Caveats: no internal Staples sales or margin data, so 'Hero' uses external proxies (swap in Staples' own sales rank when Pat shares it). Competitor depth is scored from desk research (retailer sites blocked live counts); confirm it when listings are collected.",
]
text_rows(ex, 4, msgs, 11, 36, bullet="•")
r0 = 4 + len(msgs) + 1
ex.cell(row=r0, column=1, value="Top 20 (grouped by segment; everything in this table is live)").font = f_sec
hdr_row(ex, r0 + 1, ["#", "Segment · play", "Staples path (exact L3/L4)", "Score", "Primary 1", "Primary 2", "Secondary 1", "Secondary 2",
                     "The story for the room", "Commercial moment", "Fast-start"])
LK = lambda col, cid: f"INDEX('Path Scorecard'!${col}${SC_FIRST}:${col}${SC_LAST},MATCH(\"{cid}\",'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))"
for i, r in enumerate(top20):
    rr = r0 + 2 + i
    put(ex, rr, 1, i + 1, align=center)
    put(ex, rr, 2, f'={LK(C_SEG, r["cid"])}&" · "&{LK(C_PLAY, r["cid"])}', font=f_bold, align=center)
    put(ex, rr, 3, r["path"], font=f_bold)
    put(ex, rr, 4, f"={LK(C_SCORE, r['cid'])}", align=center, fmt="0.0")
    put(ex, rr, 5, f"={LK(C_P1, r['cid'])}", font=f_bold); put(ex, rr, 6, f"={LK(C_P2, r['cid'])}")
    put(ex, rr, 7, f"={LK(C_S1, r['cid'])}", font=f_bold); put(ex, rr, 8, f"={LK(C_S2, r['cid'])}")
    put(ex, rr, 9, D[r["key"]][4]); put(ex, rr, 10, D[r["key"]][3])
    put(ex, rr, 11, f"={LK(C_FAST, r['cid'])}", align=center)
    ex.cell(row=rr, column=2).fill = PatternFill("solid", fgColor=SEG_FILL[r["seg"]])
    ex.row_dimensions[rr].height = 44
rr = r0 + 2 + len(top20)
ST = f"'Path Scorecard'!${C_STATUS}${SC_FIRST}:${C_STATUS}${SC_LAST}"; SG = f"'Path Scorecard'!${C_SEG}${SC_FIRST}:${C_SEG}${SC_LAST}"
put(ex, rr, 2, "Top-20 mix (live)", font=f_bold)
c = ex.cell(row=rr, column=3, value=(f'=COUNTIFS({ST},"Recommended",{SG},"Hero")&" Hero  ·  "&COUNTIFS({ST},"Recommended",{SG},"Probable Hero")&" Probable Hero  ·  "'
                                     f'&COUNTIFS({ST},"Recommended",{SG},"Non-Hero")&" Non-Hero  =  "&COUNTIFS({ST},"Recommended")&" shelves"'))
c.font = f_bold
ex.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=10)
r1 = rr + 2
ex.cell(row=r1, column=1, value="What changed").font = f_sec
changes = [
 "v3 vs v2: the recommendation is now a Top 20 with every segment represented (7 / 7 / 6). Non-Hero shelves are no longer capped; their play is EXPLORE (a low-risk marketplace test). Six Non-Hero shelves enter: Clocks, Desk Pads, Coffee Organizers, Desktop Privacy Panels, Office Partitions and Kettles.",
 "v3 vs v2: four competitors per shelf (2 primary + 2 secondary) with separate, documented criteria ('Competitor Criteria'). Scheels and Michaels, your original picks, are kept as secondaries where they score best (Water Bottles; Planners). For Backpacks, Bellroy and Urban Outfitters edge Scheels on adjacency, so Scheels stays an alternate.",
 "v3 vs v2: dedicated 'Hero Criteria' sheet with plain-English definitions, the decision rule, 1-5 scoring anchors for every factor, quotas and worked plays.",
 "v3 vs v2: 9 new candidate shelves researched (incl. Ben's Scheels examples: yard games, recovery tools, hydration mixes). Outdoor Games is a Non-Hero reserve; the rest are backlog.",
 "Versus your initial 9-item list: kept and drilled down Chairs & Seating, Desks, Lamps, Hydration and Planning. Parked Rugs and Labels (no L3/L4 below them), reframed Labels & Packaging to small-business packaging, and refocused Fitness to 'active workstation'. See 'Your List Review'.",
]
text_rows(ex, r1 + 1, changes, 11, 36, bullet="•")
r2 = r1 + 1 + len(changes) + 1
ex.cell(row=r2, column=1, value="Suggested sequencing").font = f_sec
seq = [
 "Wave 1a (Staples data in hand): " + ", ".join(NAME(r["key"]) for r in fast) + ".",
 "Wave 1b (needs a Staples leaf scrape): " + ", ".join(NAME(r["key"]) for r in top20 if r not in fast) + ".",
 "Keep Office Chairs as a control shelf (Hero, PROTECT): it shows the safety gate correctly rejecting look-alikes of core 1P chairs.",
 "Collect primaries at scale for the gap maths; sample secondaries by hand as a style / upside lens. Ask Pat for a Staples sales / margin rank to replace the Hero proxies.",
]
text_rows(ex, r2 + 1, seq, 11, 32, bullet="n")
r3 = r2 + 1 + len(seq) + 1
ex.cell(row=r3, column=1, value="How to read this workbook").font = f_sec
guide = [
 ("Top 20 Recommendations", "The 20 shelves (+ 9 reserves) with segment, play, 4 competitors, the super-extension, archetypes, 1P watch-outs and moments."),
 ("Segment Map", "One picture: every shelf placed as Hero, Probable Hero or Non-Hero."),
 ("Hero Criteria", "Segment definitions, the decision rule, 1-5 anchors, cut-offs and quotas (editable)."),
 ("Competitor Criteria", "Primary (K1-K5) and secondary (E1-E5) competitor criteria, pools and weights (editable)."),
 ("Your List Review", "Verdict on each of the 9 initial items, their segment mix, and exact L3/L4 drill-downs."),
 ("Selection Criteria", "Gates and the 8 path-score criteria (editable)."),
 ("Path Scorecard", f"All {len(rows)} candidate shelves with gates, segment, score, rank, status, play and 4 competitors (formulas)."),
 ("Hero Segmentation", "The two indices behind each segment, plus an L2 roll-up."),
 ("Primary / Secondary Competitor Fit", "Competitor scores per shelf; ranks 1-2 are the picks."),
 ("Competitor Profiles", "Every competitor: pool, evidence, scores, where it is used, cautions."),
 ("Parked & Excluded / Evidence & Sources / Staples Tree Extract", "What was dropped and why; every fact with its source; all terminal Staples nodes."),
]
for i, (a, b) in enumerate(guide):
    rr = r3 + 1 + i
    put(ex, rr, 2, a, font=f_bold)
    c = ex.cell(row=rr, column=3, value=b); c.font = f_body
    ex.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=10)
lg = r3 + 2 + len(guide)
put(ex, lg, 2, "Legend", font=f_bold)
c = ex.cell(row=lg, column=3, value=("Blue text on yellow = editable input. Segments: blue Hero, orange Probable Hero, grey Non-Hero. "
                                     "Status: green Recommended, amber Reserve, grey Backlog, light blue Protect, red Parked."))
c.font = f_body
ex.merge_cells(start_row=lg, start_column=3, end_row=lg, end_column=10)

del wb["Sheet"]
order_names = ["Executive Summary", "Top 20 Recommendations", "Segment Map", "Hero Criteria", "Competitor Criteria", "Your List Review",
               "Selection Criteria", "Path Scorecard", "Hero Segmentation", "Primary Competitor Fit", "Secondary Competitor Fit",
               "Competitor Profiles", "Parked & Excluded", "Evidence & Sources", "Staples Tree Extract"]
wb._sheets = [wb[n] for n in order_names]
TAB = {"Executive Summary": NAVY, "Top 20 Recommendations": "70AD47", "Segment Map": "2A78D6", "Hero Criteria": "2A78D6",
       "Competitor Criteria": "EB6834", "Your List Review": TEAL}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = TAB.get(ws.title, "A6A6A6")
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)

print("saved", OUT)
for r in top20 + reserves:
    k = r["key"]
    print(f"{r['status']:11} {r['seg']:13} #{r['rank_seg']:<2} {r['s']:5.1f} {r['play']:7} P:{roles_p[k][0]}/{roles_p[k][1]}  S:{roles_s[k][0]}/{roles_s[k][1]}  {r['path']}")
