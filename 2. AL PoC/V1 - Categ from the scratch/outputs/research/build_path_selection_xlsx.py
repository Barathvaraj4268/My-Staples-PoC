"""Build the L3/L4 path-selection workbook (research deliverable, 2026-09-29).

Every Staples path is looked up from the navigation tree by Category ID, so the
paths in the workbook match `Excels/1. Staples_Navigation_Tree_repaired.xlsx`
exactly (asserted below). Scores are analyst inputs (1-5) backed by the desk
research in the 'Evidence & Sources' sheet; weights, thresholds, weighted scores,
ranks, tiers and competitor roles are live Excel formulas.

v2 (2026-09-29) adds the Hero / Probable Hero / Non-Hero lens: a Staples
Strength Index (SSI) and a Market Attractiveness Index (MAI) per path, a
segment + marketplace play per path, a Non-Hero cap on tiers, and a segment map.
Earlier outputs (v1) are left untouched.

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
OUT = ROOT / "outputs" / "research" / "Staples_L3L4_Path_Selection_v2.xlsx"
FIG = ROOT / "outputs" / "research" / "segment_map_v2.png"

# --------------------------------------------------------------------------- data
tree = pd.read_excel(TREE, "Navigation Tree")
tree["path"] = tree[["L1", "L2", "L3", "L4"]].fillna("").apply(
    lambda r: " > ".join(x for x in r if x), axis=1)
tree = tree.set_index("Category ID", drop=False)
in_hand = pd.read_parquet(SKU, columns=["leaf_path"])["leaf_path"].value_counts()
_nav = pd.read_excel(TREE, "Summary by L1", header=3)
NAV = dict(zip(_nav.iloc[:, 0], _nav.iloc[:, 1]))   # L1 -> "Yes"/"No" (in Staples header nav)

W = [20, 15, 15, 15, 10, 10, 5, 10]            # C1..C8 default weights
KW = [30, 25, 15, 15, 15]                      # K1..K5 default weights
T1, T2, MIN_ITEMS, FAST_MIN = 84, 72, 20, 100
SW = [30, 15, 25, 30]                          # S1..S4 Staples Strength weights
MW = [35, 25, 25, 15]                          # M1..M4 Market Attractiveness weights
H_CUT, P_CUT, PROTECT_C3, NH_CAP = 65, 65, 2, "Yes"
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
 "Desk pads": "Small, flat category; the desk-mat trend is captured by Mouse Pads & Wrist Rests.",
}
PLAYS = [
 ("Hero", "EXTEND", "Staples is strong and 1P risk is manageable (C3 above the Protect cut-off).",
  "Add style / lifestyle extensions of the hero only (the 'White Chair' move); strict cannibalisation gate; invite brands Staples already buys first."),
 ("Hero", "PROTECT", "Staples is strong and design variants would substitute 1P (C3 at or below the Protect cut-off).",
  "Keep the marketplace out; use as a control shelf that shows the safety gate working."),
 ("Probable Hero", "BUILD", "The market treats it as a top category but Staples is not (yet) a destination.",
  "Fill the depth gap quickly with marketplace sellers recruited from the primary competitor's brands: the biggest commission upside."),
 ("Non-Hero", "HOLD", "Neither a Staples strength nor a market favourite.",
  "Backlog, capped at Tier 3. Revisit only if strategy or data changes."),
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
}

# competitor defaults: K4 seller recruitability, K5 data access (tested 2026-09-29)
CK = {"Wayfair": (3, 4), "Amazon": (5, 3), "Target": (3, 3), "Michaels": (3, 3),
      "The Container Store": (2, 2), "Best Buy": (3, 4), "Dick's Sporting Goods": (4, 2),
      "Scheels": (4, 2), "Uline": (2, 4), "JetPens": (3, 2), "Oriental Trading": (2, 2)}
# per path: (competitor, K1 depth, K2 aesthetic range, K3 shopper overlap)
M = {
 "Planners": [("Michaels",5,5,4),("Amazon",5,4,4),("Target",3,5,5),("JetPens",3,4,3)],
 "Water bottles": [("Dick's Sporting Goods",5,5,3),("Target",4,5,5),("Amazon",5,3,4),("Scheels",4,4,3)],
 "Desk organizers": [("Target",4,5,5),("Amazon",5,3,4),("The Container Store",5,5,4),("Wayfair",3,4,3)],
 "Accent chairs": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",3,5,4)],
 "Pencil cases": [("Amazon",5,4,4),("Target",4,5,5),("JetPens",4,5,3)],
 "Lunch": [("Amazon",5,4,4),("Target",4,5,5),("The Container Store",3,4,3)],
 "Desk lamps": [("Wayfair",5,5,4),("Amazon",5,3,4),("Target",3,5,5)],
 "Classroom decor": [("Amazon",5,4,4),("Oriental Trading",5,4,5),("Michaels",3,4,4),("Target",2,5,4)],
 "Backpacks": [("Amazon",5,4,4),("Target",4,5,5),("Dick's Sporting Goods",4,4,3),("Scheels",3,4,2)],
 "Storage baskets": [("Wayfair",5,5,4),("Target",4,5,5),("The Container Store",4,5,4)],
 "Laptop sleeves": [("Amazon",5,4,4),("Target",3,4,5),("Best Buy",3,3,4)],
 "Work totes": [("Amazon",5,4,4),("Target",3,4,4)],
 "Office desks": [("Wayfair",5,5,5),("Amazon",5,3,4)],
 "Journals": [("Amazon",5,4,4),("Michaels",4,5,4),("JetPens",4,5,3),("Target",3,5,4)],
 "Faux plants": [("Wayfair",5,5,4),("Michaels",5,4,3),("Target",3,5,4)],
 "Stickers": [("Michaels",5,5,4),("Amazon",5,4,3)],
 "Desk mats": [("Amazon",5,4,4),("Wayfair",3,4,3),("Best Buy",3,3,4)],
 "Washi tape": [("Michaels",5,5,4),("Amazon",5,4,3),("JetPens",4,5,3)],
 "Retail boxes": [("Amazon",5,4,4),("Uline",5,3,5)],
 "Frames": [("Michaels",5,5,4),("Wayfair",5,4,3),("Target",3,5,4)],
 "Clocks": [("Wayfair",5,5,4),("Amazon",5,3,3),("Target",3,5,4)],
 "Bulletin boards": [("Amazon",5,4,4),("Wayfair",4,4,4),("Target",2,5,4)],
 "Laptop stands": [("Amazon",5,4,4),("Best Buy",3,3,4),("Target",2,4,4)],
 "Desk pads": [("Amazon",5,4,4),("Wayfair",4,4,4)],
 "Table lamps": [("Wayfair",5,5,3),("Target",4,5,4)],
 "Keyboards": [("Amazon",5,4,4),("Best Buy",5,4,4)],
 "Benches": [("Wayfair",5,5,4),("Amazon",4,3,3)],
 "Calendars": [("Amazon",5,4,4),("Target",3,5,4),("Michaels",3,4,4)],
 "Breakroom chairs": [("Wayfair",5,5,4),("Amazon",4,3,3)],
 "Mailers": [("Amazon",5,4,4),("Uline",5,3,5)],
 "Label makers": [("Amazon",5,4,4),("Michaels",3,4,3),("Best Buy",2,3,3)],
 "Floor lamps": [("Wayfair",5,5,3),("Target",4,5,4)],
 "Bar stools": [("Wayfair",5,5,3),("Amazon",5,3,3)],
 "Storage bins": [("Target",4,5,5),("Amazon",5,3,4),("Wayfair",5,4,3),("The Container Store",4,4,4)],
 "Calendar boards": [("Amazon",5,4,4),("Wayfair",3,4,3)],
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
    r["play"] = ("PROTECT" if c3 <= PROTECT_C3 else "EXTEND") if r["seg"] == "Hero" else "BUILD" if r["seg"] == "Probable Hero" else "HOLD"
    r["tier"] = ("Parked (gate)" if not gates else "Tier 3" if (NH_CAP == "Yes" and r["seg"] == "Non-Hero")
                 else "Tier 1" if r["s"] >= T1 else "Tier 2" if r["s"] >= T2 else "Tier 3")
    if gates and r["seg"] == "Non-Hero" and r["s"] >= T2:
        r["why"] += " Capped at Tier 3 because it is Non-Hero."
by_key = {r["key"]: r for r in rows}
short = [r for r in rows if r["tier"] in ("Tier 1", "Tier 2")]
short.sort(key=lambda r: -r["s"])
assert {r["key"] for r in short} <= set(D) and set(D) == set(M), ({r["key"] for r in short} - set(D))
assert set(SM) == {r["key"] for r in rows}
roles = {}
for k, cands in M.items():
    ranked = sorted(cands, key=lambda c: -kscore(c[1], c[2], c[3], c[0]))
    roles[k] = (ranked[0][0], ranked[1][0])
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

def tier_cf(ws, rng, col_letter, first):
    for t, color in TIER_FILL.items():
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'LEFT(${col_letter}{first},{len(t)})="{t}"'], fill=PatternFill("solid", fgColor=color)))

# --------------------------------------------------------------- Selection Criteria
cr = sheet("Selection Criteria", "Selection criteria, weights and thresholds",
           "Yellow cells with blue text are editable. Every score, rank, tier and competitor role in the workbook recalculates from them.")
cr.column_dimensions["A"].width = 8
for col, w in zip("BCDEF", [34, 48, 38, 38, 10]):
    cr.column_dimensions[col].width = w
cr["A4"] = "A. Hard gates: a path must pass all four to be scored"; cr["A4"].font = f_sec
for i, h in enumerate(["Code", "Gate", "Rule", "Why it matters"], 1):
    c = cr.cell(row=5, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
gates_txt = [
 ("G1", "Granularity", "The node is an L3 or L4 in the Staples navigation tree (Level 3 or 4).", "Your feedback: L2 is too coarse. The gap must be readable on one shelf."),
 ("G2", "Brand authority", "The node sits on a Staples work / school / small-business mission (not mis-shelved or leisure).", "Pat: build momentum 'without straying too far from Staples' brand authority'."),
 ("G3", "Staples baseline", "Staples lists at least the minimum item count (threshold in section C).", "Needs a nearest-Staples-family comparison and a share baseline."),
 ("G4", "No partner conflict", "No Staples strategic partner already owns the assortment on Staples.com.", "Party City (Apr-2026) now supplies party décor, gift bags & wrap: marketplace adds would compete with a partner."),
]
for i, g in enumerate(gates_txt):
    for j, v in enumerate(g, 1):
        put(cr, 6 + i, j, v, font=f_bold if j == 1 else f_body)
cr["A11"] = "B. Weighted path criteria: each scored 1-5, weighted score scaled to 0-100"; cr["A11"].font = f_sec
for i, h in enumerate(["Code", "Criterion", "What it measures", "Scores 5 when", "Scores 1 when", "Weight"], 1):
    c = cr.cell(row=12, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
crit = [
 ("C1", "Design & lifestyle variance ('White Chair' test)", "Do shoppers choose on colour, material, style or 'vibe' rather than spec?", "Style drives the choice; many aesthetic archetypes exist", "Pure spec / commodity"),
 ("C2", "Competitor depth advantage", "How much wider the best competitor's assortment is on this shelf", "Best competitor is several times wider, with archetypes Staples lacks", "Parity with Staples"),
 ("C3", "1P protection (low cannibalisation)", "Can design variants be added without substituting Staples' core best-sellers?", "Design variants clearly distinct from the 1P core", "Direct substitute of core 1P"),
 ("C4", "Core-adjacency & shopper mission", "Same work, school or small-business mission as Staples' core", "Sits on the core Staples mission", "Residential or leisure mission"),
 ("C5", "Basket attach & quick-win economics", "Parcel-shippable, lower ticket, attaches to core baskets (Pat: +1-2 items per checkout)", "Under ~$50, parcel, frequent attach", "Oversize freight, considered purchase"),
 ("C6", "Business familiarity & story", "Would the Staples room recognise it instantly (header nav, Staples' own recent moves, visible trend)?", "Header-nav shelf + Staples' own recent move or a famous trend", "Obscure shelf"),
 ("C7", "Commercial moment", "Peaks in the Q1-2027 window or BTS-2027 (Pat's seasonal anchor)", "Peaks in Jan-Mar or Jul-Sep 2027", "No clear moment"),
 ("C8", "Data readiness", "Staples SKUs already in hand + competitor already scraped or reachable", "Both sides in hand / reachable", "Neither side available"),
]
for i, c in enumerate(crit):
    for j, v in enumerate(c, 1):
        put(cr, 13 + i, j, v, font=f_bold if j == 1 else f_body)
    put(cr, 13 + i, 6, W[i], font=f_input, fill=fill_input, align=center)
put(cr, 21, 5, "Total weight", font=f_bold)
put(cr, 21, 6, "=SUM(F13:F20)", font=f_bold, align=center)
cr["A23"] = "C. Thresholds"; cr["A23"].font = f_sec
for i, h in enumerate(["Code", "Parameter", "Value", "Note"], 1):
    c = cr.cell(row=24, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
thr = [("T1", "Tier 1 cut-off (score >=)", T1, "Tier 1 = PoC wave 1 (10 paths at default weights)."),
       ("T2", "Tier 2 cut-off (score >=)", T2, "Tier 2 = wave 2 / swap-ins (25 paths at default weights)."),
       ("G3", "Minimum Staples items on the leaf", MIN_ITEMS, "Feeds gate G3."),
       ("FS", "Fast-start: min Staples SKUs already in the sample", FAST_MIN, "Tier 1 paths above this can start without a new Staples scrape.")]
for i, (a, b, v, n) in enumerate(thr):
    put(cr, 25 + i, 1, a, font=f_bold); put(cr, 25 + i, 2, b)
    put(cr, 25 + i, 3, v, font=f_input, fill=fill_input, align=center); put(cr, 25 + i, 4, n)
cr["A30"] = "D. Competitor selection criteria: each competitor scored 1-5 per path"; cr["A30"].font = f_sec
for i, h in enumerate(["Code", "Criterion", "What it measures", "Scores 5 when", "Scores 1 when", "Weight"], 1):
    c = cr.cell(row=31, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
kcrit = [
 ("K1", "Assortment depth on this shelf", "Breadth of the competitor's range on the mapped node", "Several times Staples' range, many sub-types", "Similar or smaller than Staples"),
 ("K2", "Aesthetic & lifestyle range", "Variety of styles, colourways, materials and price tiers", "Destination for design-led choice", "Utility only"),
 ("K3", "Shopper overlap with Staples", "Same home-office, student, teacher or small-business shopper", "Same shopper and mission", "Different shopper (e.g. hunters, hobby pros)"),
 ("K4", "Seller recruitability", "Can the brands / suppliers behind the range be invited onto Staples Marketplace?", "Multi-brand 3P or drop-ship suppliers", "Mostly own / private label"),
 ("K5", "Data accessibility", "Can the team collect listing-level data at PoC speed? (tested 29-Sep-2026)", "Already scraped or official API", "Bot-blocked, needs a paid proxy"),
]
for i, c in enumerate(kcrit):
    for j, v in enumerate(c, 1):
        put(cr, 32 + i, j, v, font=f_bold if j == 1 else f_body)
    put(cr, 32 + i, 6, KW[i], font=f_input, fill=fill_input, align=center)
put(cr, 37, 5, "Total weight", font=f_bold)
put(cr, 37, 6, "=SUM(F32:F36)", font=f_bold, align=center)
cr["A39"] = "E. How primary and secondary competitors are picked"; cr["A39"].font = f_sec
for i, t in enumerate([
    "For each Tier 1 / Tier 2 path, 2-4 plausible competitors are scored K1-K5 in 'Competitor Fit'. Rank 1 = Primary, rank 2 = Secondary; ties are broken by K1 (depth).",
    "K4 and K5 are competitor-level (see 'Competitor Profiles'); K1-K3 are path-specific.",
    "Office Depot is deliberately NOT used as a primary: its assortment mirrors Staples' utility range, so it produces little gap signal. Use it only as a sanity check.",
    "All scores are analyst judgement from the desk research in 'Evidence & Sources'. Live SKU counts could not be pulled (retailer sites bot-protected); confirm depth with the S0 scrape before the analysis run."]):
    cr.cell(row=40 + i, column=1, value=f"{i + 1}.").font = f_body
    c = cr.cell(row=40 + i, column=2, value=t); c.font, c.alignment = f_body, wrap
    cr.merge_cells(start_row=40 + i, start_column=2, end_row=40 + i, end_column=6)
    cr.row_dimensions[40 + i].height = 28
for r in range(6, 21):
    cr.row_dimensions[r].height = 42
WREF = [f"'Selection Criteria'!$F${13 + i}" for i in range(8)]
KREF = [f"'Selection Criteria'!$F${32 + i}" for i in range(5)]

cr["A46"] = "F. Hero segmentation: the second lens (Hero / Probable Hero / Non-Hero)"; cr["A46"].font = f_sec
c = cr.cell(row=47, column=1, value=(
    "Hero = Staples' own top categories. Probable Hero = the market's top categories where Staples is not (yet) a destination: room for improvement. "
    "Non-Hero = neither. There is no internal sales or margin data, so both axes use external proxies; when Staples shares sales / margin rank, "
    "replace S1-S4 with it and everything recalculates."))
c.font, c.alignment = f_body, wrap
cr.merge_cells("A47:F47"); cr.row_dimensions[47].height = 44
seg_parts = [
 (48, "Staples Strength Index (SSI): is it a Staples hero?", [
  ("S1", "Assortment depth", "Staples items on the leaf, banded (computed from the tree)", f">= {S1_BANDS[0]:,} items", f"< {S1_BANDS[3]} items"),
  ("S2", "Shelf visibility", "Is the L1 in Staples' header navigation? (computed from the tree)", "In the header nav", "Not in the header nav"),
  ("S3", "Merchandising investment", "Exclusives, collaborations, own brands, BTS / seasonal feature slots", "Staples exclusives or own brand + feature slots", "No visible investment"),
  ("S4", "Shopper association", "Would a shopper think of Staples first for this?", "Staples is a go-to destination", "Nobody thinks of Staples for this")], SW),
 (54, "Market Attractiveness Index (MAI): is it a market hero?", [
  ("M1", "Demand momentum", "Growth and trend signals (Circana, trade press, social)", "Clear growth or a viral trend", "Flat or declining"),
  ("M2", "Competitor destination strength", "Do competitors run it as a department with specialist depth?", "Several retailers treat it as a destination", "Few competitors care"),
  ("M3", "Category size & purchase frequency", "How big the category is and how often it is bought", "Large and frequently bought", "Small, rare purchase"),
  ("M4", "Occasion strength", "Reuses C7 (Q1-2027 / BTS-2027 moment) from the scorecard", "Strong Q1 or BTS peak", "No moment")], MW),
]
for top, title, items, wts in seg_parts:
    for i, h in enumerate(["Code", title, "What it measures", "Scores 5 when", "Scores 1 when", "Weight"], 1):
        c = cr.cell(row=top, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    for i, it in enumerate(items):
        for j, v in enumerate(it, 1):
            put(cr, top + 1 + i, j, v, font=f_bold if j == 1 else f_body)
        put(cr, top + 1 + i, 6, wts[i], font=f_input, fill=fill_input, align=center)
        cr.row_dimensions[top + 1 + i].height = 30
    put(cr, top + 5, 5, "Total weight", font=f_bold)
    put(cr, top + 5, 6, f"=SUM(F{top + 1}:F{top + 4})", font=f_bold, align=center)
cr.row_dimensions[48].height = cr.row_dimensions[54].height = 30
for i, h in enumerate(["Code", "Segmentation parameter", "Value", "Note"], 1):
    c = cr.cell(row=61, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
seg_thr = [("H", "Hero cut-off (SSI >=)", H_CUT, "Checked first: a strong Staples shelf is a Hero whatever the market does."),
           ("P", "Probable Hero cut-off (MAI >=)", P_CUT, "Applied to shelves that are not Heroes."),
           ("PR", "Protect cut-off (Hero with C3 1P-protection <=)", PROTECT_C3, "Hero shelves where design variants would substitute 1P: keep the marketplace out."),
           ("NC", "Cap Non-Hero shelves at Tier 3? (Yes / No)", NH_CAP, "PoC time goes to Heroes and Probable Heroes."),
           ("B5", "S1 band: items >= this scores 5", S1_BANDS[0], ""), ("B4", "S1 band: scores 4", S1_BANDS[1], ""),
           ("B3", "S1 band: scores 3", S1_BANDS[2], ""), ("B2", "S1 band: scores 2 (else 1)", S1_BANDS[3], ""),
           ("NY", "S2 score if the L1 is in the header nav", NAV_YES, ""), ("NN", "S2 score if not", NAV_NO, "")]
for i, (a, b, v, n) in enumerate(seg_thr):
    put(cr, 62 + i, 1, a, font=f_bold); put(cr, 62 + i, 2, b)
    put(cr, 62 + i, 3, v, font=f_input, fill=fill_input, align=center); put(cr, 62 + i, 4, n)
cr["A73"] = "G. Segments and marketplace plays"; cr["A73"].font = f_sec
for i, h in enumerate(["Segment", "Play", "When", "What Staples Marketplace does"], 1):
    c = cr.cell(row=74, column=i, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
for i, pl in enumerate(PLAYS):
    for j, v in enumerate(pl, 1):
        put(cr, 75 + i, j, v, font=f_bold if j <= 2 else f_body)
    cr.cell(row=75 + i, column=1).fill = PatternFill("solid", fgColor=SEG_FILL[pl[0]])
    cr.row_dimensions[75 + i].height = 40
cr.merge_cells("D75:F75"); cr.merge_cells("D76:F76"); cr.merge_cells("D77:F77"); cr.merge_cells("D78:F78")
SC_ = "'Selection Criteria'!"
SREF = [f"{SC_}$F${49 + i}" for i in range(4)]
MREF = [f"{SC_}$F${55 + i}" for i in range(4)]

# --------------------------------------------------------------- Competitor Fit
cf = sheet("Competitor Fit", "Competitor fit per shortlisted path (Tier 1 + Tier 2 at v1, incl. two now capped at Tier 3)",
           "K1-K3 are path-specific inputs (blue). K4-K5 come from 'Competitor Profiles'. Role = rank within the path: 1 Primary, 2 Secondary.")
header(cf, 4, ["Path ID", "Staples path (exact)", "Competitor", "K1 Depth", "K2 Aesthetic", "K3 Shopper overlap",
               "K4 Recruitable", "K5 Data access", "Fit score", "Rank in path", "Role", "Lookup key"],
       [8, 62, 22, 9, 10, 10, 11, 10, 9, 9, 22, 18])
pid = {}
for i, r in enumerate(sorted([r for r in rows if r["gate"]], key=lambda r: -r["s"]), 1):
    pid[r["key"]] = f"P{i:02d}"
xi = 0
for r in rows:
    if not r["gate"]:
        xi += 1
        pid[r["key"]] = f"X{xi:02d}"
row = 5
CF_FIRST = row
for r in sorted([by_key[k] for k in M], key=lambda r: -r["s"]):
    for comp, k1, k2, k3 in M[r["key"]]:
        k4, k5 = CK[comp]
        vals = [pid[r["key"]], r["path"], comp, k1, k2, k3, k4, k5]
        for j, v in enumerate(vals, 1):
            put(cf, row, j, v, font=f_input if j in (4, 5, 6) else f_body,
                align=center if j >= 4 else wrap, fill=fill_input if j in (4, 5, 6) else None)
        put(cf, row, 9, "=(" + "+".join(f"{c}{row}*{KREF[i]}" for i, c in enumerate("DEFGH"))
            + f")/(5*'Selection Criteria'!$F$37)*100+D{row}/1000", align=center, fmt="0.0")
        row += 1
CF_LAST = row - 1
for rr in range(CF_FIRST, CF_LAST + 1):
    put(cf, rr, 10, f"=COUNTIFS($A${CF_FIRST}:$A${CF_LAST},A{rr},$I${CF_FIRST}:$I${CF_LAST},\">\"&I{rr})+1", align=center)
    put(cf, rr, 11, f'=IF(J{rr}=1,"Primary",IF(J{rr}=2,"Secondary","Alternate / style reference"))')
    put(cf, rr, 12, f'=A{rr}&"|"&K{rr}')
cf.freeze_panes = "D5"
cf.auto_filter.ref = f"A4:L{CF_LAST}"
cf.conditional_formatting.add(f"A{CF_FIRST}:L{CF_LAST}", FormulaRule(formula=[f'$K{CF_FIRST}="Primary"'], fill=PatternFill("solid", fgColor="C6EFCE")))
cf.conditional_formatting.add(f"A{CF_FIRST}:L{CF_LAST}", FormulaRule(formula=[f'$K{CF_FIRST}="Secondary"'], fill=PatternFill("solid", fgColor="DDEBF7")))
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True)
cf.add_data_validation(dv); dv.add(f"D{CF_FIRST}:F{CF_LAST}")

def comp_lookup(p_cell, role, fallback):
    rng_c = f"'Competitor Fit'!$C${CF_FIRST}:$C${CF_LAST}"
    rng_k = f"'Competitor Fit'!$L${CF_FIRST}:$L${CF_LAST}"
    fb = f'"{fallback}"' if fallback else '"-"'
    return f'=IFERROR(INDEX({rng_c},MATCH({p_cell}&"|{role}",{rng_k},0)),{fb})'

# --------------------------------------------------------------- Path Scorecard
sc = sheet("Path Scorecard", "Path scorecard: every candidate Staples L3/L4 path",
           "Blue cells are analyst inputs. Gates, segment, score, rank, tier, play and competitors are formulas. Sort or filter freely; IDs are stable.")
cols = ["Path ID", "Staples path (exact, from nav tree)", "L1", "L2", "L3", "L4", "Level", "Category ID", "Staples items (site)",
        "Staples SKUs in hand (sample)", "Link to your list", "G1 Granularity", "G2 Brand authority", "G3 Baseline", "G4 No partner conflict",
        "All gates", "Hero segment", "C1 Design variance", "C2 Comp. depth", "C3 1P protection", "C4 Core-adjacency", "C5 Basket attach",
        "C6 Familiarity", "C7 Moment", "C8 Data readiness", "Score (0-100)", "Rank", "Tier", "Marketplace play", "Fast-start",
        "Primary competitor", "Secondary competitor", "Rationale / gate note"]
header(sc, 4, cols, [8, 60, 18, 20, 24, 24, 6, 11, 9, 10, 22, 9, 9, 9, 9, 8, 13] + [7] * 8 + [8, 6, 9, 11, 8, 20, 20, 60])
C_SEG, C_SCORE, C_RANK, C_TIER, C_PLAY, C_FAST, C_PRI, C_SEC = "Q", "Z", "AA", "AB", "AC", "AD", "AE", "AF"
C3_COL, C7_COL = "T", "X"
SC_FIRST = 5
order = sorted(rows, key=lambda r: (not r["gate"], -(r["s"] or 0)))
SC_LAST = SC_FIRST + len(order) - 1
HS_FIRST, HS_LAST = 5, 5 + len(order) - 1
HS_ = "'Hero Segmentation'!"
for i, r in enumerate(order):
    rr = SC_FIRST + i
    n = r["n"]
    base = [pid[r["key"]], r["path"], n["L1"], n["L2"], n["L3"] if pd.notna(n["L3"]) else "", n["L4"] if pd.notna(n["L4"]) else "",
            r["level"], r["cid"], r["items"], r["hand"], r["link"]]
    for j, v in enumerate(base, 1):
        put(sc, rr, j, v, align=center if j in (1, 7, 8, 9, 10) else wrap, fmt="#,##0" if j in (9, 10) else None)
    put(sc, rr, 12, f'=IF(OR(G{rr}=3,G{rr}=4),"Pass","Fail")', align=center)
    put(sc, rr, 13, r["g2"], font=f_input, fill=fill_input, align=center)
    put(sc, rr, 14, f"=IF(I{rr}>={SC_}$C$27,\"Pass\",\"Fail\")", align=center)
    put(sc, rr, 15, r["g4"], font=f_input, fill=fill_input, align=center)
    put(sc, rr, 16, f'=IF(AND(L{rr}="Pass",M{rr}="Pass",N{rr}="Pass",O{rr}="Pass"),"Pass","Fail")', font=f_bold, align=center)
    put(sc, rr, 17, f"=INDEX({HS_}$Q${HS_FIRST}:$Q${HS_LAST},MATCH($A{rr},{HS_}$A${HS_FIRST}:$A${HS_LAST},0))", font=f_bold, align=center)
    for j in range(8):
        v = r["sc"][j] if r["sc"] else None
        put(sc, rr, 18 + j, v, font=f_input, fill=fill_input, align=center)
    q = "+".join(f"{get_column_letter(18 + j)}{rr}*{WREF[j]}" for j in range(8))
    put(sc, rr, 26, f'=IF(P{rr}="Pass",({q})/(5*{SC_}$F$21)*100,"-")', font=f_bold, align=center, fmt="0.0")
    put(sc, rr, 27, f'=IF(P{rr}="Pass",COUNTIFS($P${SC_FIRST}:$P${SC_LAST},"Pass",$Z${SC_FIRST}:$Z${SC_LAST},">"&Z{rr})+1,"-")', align=center)
    put(sc, rr, 28, (f'=IF(P{rr}<>"Pass","Parked (gate)",IF(AND({SC_}$C$65="Yes",Q{rr}="Non-Hero"),"Tier 3",'
                     f'IF(Z{rr}>={SC_}$C$25,"Tier 1",IF(Z{rr}>={SC_}$C$26,"Tier 2","Tier 3"))))'), font=f_bold, align=center)
    put(sc, rr, 29, f"=INDEX({HS_}$R${HS_FIRST}:$R${HS_LAST},MATCH($A{rr},{HS_}$A${HS_FIRST}:$A${HS_LAST},0))", align=center)
    put(sc, rr, 30, f"=IF(AND(AB{rr}=\"Tier 1\",J{rr}>={SC_}$C$28),\"Yes\",\"\")", align=center)
    t3 = r["t3"].split(" / ") if r["t3"] else ["", ""]
    put(sc, rr, 31, comp_lookup(f"$A{rr}", "Primary", t3[0] if r["gate"] else ""))
    put(sc, rr, 32, comp_lookup(f"$A{rr}", "Secondary", (t3[1] if len(t3) > 1 else "") if r["gate"] else ""))
    put(sc, rr, 33, r["why"])
sc.freeze_panes = "C5"
sc.auto_filter.ref = f"A4:AG{SC_LAST}"
tier_cf(sc, f"AB{SC_FIRST}:AB{SC_LAST}", "AB", SC_FIRST)
seg_cf(sc, f"Q{SC_FIRST}:Q{SC_LAST}", "Q", SC_FIRST)
dv2 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True)
sc.add_data_validation(dv2); dv2.add(f"R{SC_FIRST}:Y{SC_LAST}")
dv3 = DataValidation(type="list", formula1='"Pass,Fail"', showErrorMessage=True)
sc.add_data_validation(dv3); dv3.add(f"M{SC_FIRST}:M{SC_LAST}"); dv3.add(f"O{SC_FIRST}:O{SC_LAST}")

def sc_lookup(col, cid_cell):
    return f"=INDEX('Path Scorecard'!${col}${SC_FIRST}:${col}${SC_LAST},MATCH({cid_cell},'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))"

# --------------------------------------------------------------- Hero Segmentation
hs = sheet("Hero Segmentation", "Hero segmentation: Staples strength vs market attractiveness, per L3/L4 path",
           "S1-S2 and M4 are computed; S3, S4, M1-M3 are analyst inputs (blue). Segment: Hero if SSI >= cut-off, else Probable Hero if MAI >= cut-off, else Non-Hero.")
header(hs, 4, ["Path ID", "Staples path (exact)", "Link to your list", "L1 > L2", "Staples items", "L1 in header nav",
               "S1 Depth", "S2 Visibility", "S3 Merch. investment", "S4 Shopper association", "SSI Staples strength (0-100)",
               "M1 Demand momentum", "M2 Competitor destination", "M3 Size & frequency", "M4 Occasion (=C7)", "MAI Market attractiveness (0-100)",
               "Segment", "Marketplace play", "Tier", "Evidence / note",
               "Chart helpers (#N/A = point hidden on purpose)", "chart: SSI Hero", "chart: MAI Hero", "chart: SSI Prob. Hero", "chart: MAI Prob. Hero", "chart: SSI Non-Hero", "chart: MAI Non-Hero"],
       [8, 56, 22, 34, 9, 9, 7, 8, 9, 9, 10, 9, 10, 9, 9, 10, 13, 11, 12, 60, 14, 9, 9, 9, 9, 9, 9])
for i, r in enumerate(order):
    rr = HS_FIRST + i
    n = r["n"]
    s3, s4, m1, m2, m3 = SM[r["key"]]
    put(hs, rr, 1, pid[r["key"]], align=center); put(hs, rr, 2, r["path"], font=f_bold); put(hs, rr, 3, r["link"])
    put(hs, rr, 4, f'{n["L1"]} > {n["L2"]}'); put(hs, rr, 5, r["items"], align=center, fmt="#,##0"); put(hs, rr, 6, r["nav"], align=center)
    put(hs, rr, 7, f"=IF(E{rr}>={SC_}$C$66,5,IF(E{rr}>={SC_}$C$67,4,IF(E{rr}>={SC_}$C$68,3,IF(E{rr}>={SC_}$C$69,2,1))))", align=center)
    put(hs, rr, 8, f'=IF(F{rr}="Yes",{SC_}$C$70,{SC_}$C$71)', align=center)
    put(hs, rr, 9, s3, font=f_input, fill=fill_input, align=center); put(hs, rr, 10, s4, font=f_input, fill=fill_input, align=center)
    put(hs, rr, 11, f"=(G{rr}*{SREF[0]}+H{rr}*{SREF[1]}+I{rr}*{SREF[2]}+J{rr}*{SREF[3]})/(5*{SC_}$F$53)*100", font=f_bold, align=center, fmt="0")
    for j, v in enumerate((m1, m2, m3)):
        put(hs, rr, 12 + j, v, font=f_input, fill=fill_input, align=center)
    lk = f"INDEX('Path Scorecard'!${C7_COL}${SC_FIRST}:${C7_COL}${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))"
    put(hs, rr, 15, f'=IF({lk}="",3,{lk})', align=center)
    put(hs, rr, 16, f"=(L{rr}*{MREF[0]}+M{rr}*{MREF[1]}+N{rr}*{MREF[2]}+O{rr}*{MREF[3]})/(5*{SC_}$F$59)*100", font=f_bold, align=center, fmt="0")
    put(hs, rr, 17, f'=IF(K{rr}>={SC_}$C$62,"Hero",IF(P{rr}>={SC_}$C$63,"Probable Hero","Non-Hero"))', font=f_bold, align=center)
    c3 = f"INDEX('Path Scorecard'!${C3_COL}${SC_FIRST}:${C3_COL}${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))"
    put(hs, rr, 18, f'=IF(Q{rr}="Hero",IF({c3}<={SC_}$C$64,"PROTECT","EXTEND"),IF(Q{rr}="Probable Hero","BUILD","HOLD"))', font=f_bold, align=center)
    put(hs, rr, 19, f"=INDEX('Path Scorecard'!${C_TIER}${SC_FIRST}:${C_TIER}${SC_LAST},MATCH($A{rr},'Path Scorecard'!$A${SC_FIRST}:$A${SC_LAST},0))", align=center)
    put(hs, rr, 20, SEG_NOTE.get(r["key"], ""))
    for j, (seg, src) in enumerate([("Hero", "K"), ("Hero", "P"), ("Probable Hero", "K"), ("Probable Hero", "P"), ("Non-Hero", "K"), ("Non-Hero", "P")]):
        c = hs.cell(row=rr, column=22 + j, value=f'=IF(AND($Q{rr}="{seg}",$S{rr}<>"Parked (gate)"),{src}{rr},NA())')
        c.font = Font(name=F, size=8, color="808080"); c.number_format = "0"
seg_cf(hs, f"Q{HS_FIRST}:Q{HS_LAST}", "Q", HS_FIRST)
tier_cf(hs, f"S{HS_FIRST}:S{HS_LAST}", "S", HS_FIRST)
dv4 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True)
hs.add_data_validation(dv4); dv4.add(f"I{HS_FIRST}:J{HS_LAST}"); dv4.add(f"L{HS_FIRST}:N{HS_LAST}")
hs.freeze_panes = "C5"
hs.auto_filter.ref = f"A4:T{HS_LAST}"
# L2 roll-up: shows why segmenting at L2 hides the answer
R2 = HS_LAST + 3
hs.cell(row=R2 - 1, column=1, value="L2 roll-up: how many scored L3/L4 paths of each segment sit under every Staples L2 (live)").font = f_sec
for j, h in enumerate(["", "L1 > L2", "Hero", "Probable Hero", "Non-Hero", "Parked (gate)", "Reading"], 1):
    c = hs.cell(row=R2, column=j, value=h or None)
    if h:
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
l2s = sorted({f'{r["n"]["L1"]} > {r["n"]["L2"]}' for r in rows})
for i, l2 in enumerate(l2s):
    rr = R2 + 1 + i
    put(hs, rr, 2, l2, font=f_bold)
    D_ = f"$D${HS_FIRST}:$D${HS_LAST}"; Q_ = f"$Q${HS_FIRST}:$Q${HS_LAST}"; S_ = f"$S${HS_FIRST}:$S${HS_LAST}"
    for j, seg in enumerate(["Hero", "Probable Hero", "Non-Hero"]):
        put(hs, rr, 3 + j, f'=COUNTIFS({D_},$B{rr},{Q_},"{seg}",{S_},"<>Parked (gate)")', align=center)
    put(hs, rr, 6, f'=COUNTIFS({D_},$B{rr},{S_},"Parked (gate)")', align=center)
    put(hs, rr, 7, f'=IF((C{rr}>0)+(D{rr}>0)+(E{rr}>0)>=2,"Mixed: one L2 holds several segments, so decide at L3/L4",IF(C{rr}+D{rr}+E{rr}=0,"Only parked paths","Single segment"))')
HS_R2_FIRST, HS_R2_LAST = R2 + 1, R2 + len(l2s)

# --------------------------------------------------------------- Segment Map (PNG snapshot + live native chart)
pts = [r for r in rows if r["gate"]]
fig, ax = plt.subplots(figsize=(13, 9), dpi=150)
fig.patch.set_facecolor("#fcfcfb"); ax.set_facecolor("#fcfcfb")
ax.axvspan(H_CUT, 100, color="#2a78d6", alpha=0.05, lw=0)
ax.fill_between([0, H_CUT], P_CUT, 100, color="#eb6834", alpha=0.05, lw=0)
ax.axvline(H_CUT, color="#52514e", lw=1, ls=(0, (4, 3)))
ax.plot([0, H_CUT], [P_CUT, P_CUT], color="#52514e", lw=1, ls=(0, (4, 3)))
ax.text(98, 99, "HERO\nStaples' own strength: EXTEND (or PROTECT)", ha="right", va="top", fontsize=10, fontweight="bold", color="#1f3864")
ax.text(22, 99, "PROBABLE HERO\nMarket favourite, Staples not yet: BUILD", ha="left", va="top", fontsize=10, fontweight="bold", color="#8a3a12")
ax.text(22, 45, "NON-HERO\nNeither: HOLD", ha="left", va="bottom", fontsize=10, fontweight="bold", color="#52514e")
style = {"Tier 1": dict(s=150, lw=1.6), "Tier 2": dict(s=70, lw=1.2), "Tier 3": dict(s=45, lw=1.2)}
for seg in ["Non-Hero", "Probable Hero", "Hero"]:
    for tier in ["Tier 3", "Tier 2", "Tier 1"]:
        sub = [r for r in pts if r["seg"] == seg and r["tier"] == tier]
        if not sub:
            continue
        face = SEG_HEX[seg] if tier != "Tier 3" else "#fcfcfb"
        ax.scatter([r["ssi"] for r in sub], [r["mai"] for r in sub], s=style[tier]["s"], c=face,
                   edgecolors=SEG_HEX[seg] if tier == "Tier 3" else "#fcfcfb", linewidths=style[tier]["lw"], zorder=3)
NAME = lambda k: k[0].upper() + k[1:]
# hand-placed labels for crowded spots: text position (data coords) and alignment; others sit just right of the point
POS = {
 "Pencil cases": (59.2, 89.5, "right"), "Accent chairs": (54.2, 81.5, "right"), "Desk lamps": (61.7, 82.3, "left"),
 "Faux plants": (61.7, 80.6, "left"), "Breakroom chairs": (54.3, 71.5, "right"), "Desk mats": (61.3, 70.0, "left"),
 "Floor lamps": (37.3, 67.4, "right"), "Benches": (44.7, 67.4, "left"), "Work totes": (40, 63.4, "center"),
 "Retail boxes": (46, 63.4, "center"), "Calendar boards": (60, 63.4, "center"), "Office desks": (94.3, 77.7, "right"),
 "Calendars": (94.3, 75.4, "right"), "Keyboards": (76.3, 74.7, "right"), "Laptop sleeves": (63.5, 73.4, "left"),
}
groups = {}
for r in pts:
    if r["tier"] in ("Tier 1", "Tier 2"):
        groups.setdefault((round(r["ssi"], 1), round(r["mai"], 1), r["tier"]), []).append(r)
for (x, y, tier), g in groups.items():
    g.sort(key=lambda r: -r["s"])
    t1 = tier == "Tier 1"
    label = " · ".join(NAME(r["key"]) for r in g)
    tx, ty, ha = POS.get(g[0]["key"], (x + 0.45, y - 0.35, "left"))
    ax.annotate(label, (x, y), xytext=(tx, ty), textcoords="data", ha=ha, va="center",
                fontsize=9 if t1 else 7.6, fontweight="bold" if t1 else "normal",
                color="#0b0b0b" if t1 else "#52514e", zorder=4)
prot = [r for r in pts if r["play"] == "PROTECT"]
for (x, y) in sorted({(round(r["ssi"], 1), round(r["mai"], 1)) for r in prot}):
    names = " · ".join(NAME(r["key"]) for r in prot if (round(r["ssi"], 1), round(r["mai"], 1)) == (x, y))
    edge = x > 95
    ax.annotate(f"{names} (PROTECT)", (x, y), xytext=(x - 0.3 if edge else x + 0.6, y + 1.6 if edge else y - 0.4), textcoords="data",
                ha="right" if edge else "left", va="center",
                fontsize=7.4, fontstyle="italic", color="#52514e", zorder=4)
ax.set_xlim(20, 102); ax.set_ylim(44, 100)
ax.set_xlabel("Staples Strength Index (SSI): depth, visibility, merchandising investment, shopper association", fontsize=10, color="#52514e")
ax.set_ylabel("Market Attractiveness Index (MAI): momentum, competitor destination, size, occasion", fontsize=10, color="#52514e")
ax.set_title("Where each Staples L3/L4 shelf sits: Hero, Probable Hero, Non-Hero  (labels: Tier 1 bold, Tier 2 grey)",
             fontsize=12.5, fontweight="bold", color="#0b0b0b", loc="left")
for sp in ["top", "right"]:
    ax.spines[sp].set_visible(False)
for sp in ["left", "bottom"]:
    ax.spines[sp].set_color("#c3c2b7")
ax.tick_params(colors="#52514e", labelsize=9)
ax.grid(color="#e5e4df", lw=0.6, zorder=0)
from matplotlib.lines import Line2D
leg = [Line2D([0], [0], marker="o", ls="", markerfacecolor=SEG_HEX[k], markeredgecolor=SEG_HEX[k], markersize=9, label=k) for k in SEG_HEX]
leg += [Line2D([0], [0], marker="o", ls="", markerfacecolor="#52514e", markeredgecolor="#fcfcfb", markersize=11, label="Tier 1 (large)"),
        Line2D([0], [0], marker="o", ls="", markerfacecolor="#52514e", markeredgecolor="#fcfcfb", markersize=7, label="Tier 2"),
        Line2D([0], [0], marker="o", ls="", markerfacecolor="#fcfcfb", markeredgecolor="#52514e", markersize=6, label="Tier 3 (hollow)")]
ax.legend(handles=leg, loc="lower right", fontsize=8.5, frameon=False, ncol=2)
fig.tight_layout()
fig.savefig(FIG, facecolor=fig.get_facecolor())
plt.close(fig)

smap = sheet("Segment Map", "Segment map: Staples strength vs market attractiveness",
             "Picture = snapshot at default inputs (labelled). The Excel chart further down is live: it moves when you edit inputs or cut-offs.")
smap.column_dimensions["A"].width = 3
img = XLImage(str(FIG)); img.width, img.height = 1170, 810
smap.add_image(img, "B4")
notes = [
 "How to read it: right of the dashed line = Hero (Staples is already strong); top-left = Probable Hero (the market loves it, Staples is not yet a destination); bottom-left = Non-Hero.",
 "Plays: Hero shelves get EXTEND (add style extensions only) or PROTECT (keep the marketplace out, e.g. Office Chairs, Pens); Probable Heroes get BUILD (fill the gap with marketplace sellers); Non-Heroes are held.",
 "Tier 1 mixes both growth stories: Heroes to extend (Planners, Desk Organizers, Backpacks, Classroom Decor) and Probable Heroes to build (Water Bottles, Accent Chairs, Desk Lamps, Pencil Cases, Lunch, Storage Baskets).",
]
for i, t in enumerate(notes):
    c = smap.cell(row=47 + i, column=2, value=t); c.font, c.alignment = f_body, wrap
    smap.merge_cells(start_row=47 + i, start_column=2, end_row=47 + i, end_column=16)
    smap.row_dimensions[47 + i].height = 30
ch = ScatterChart()
ch.title = "Live segment map (updates with inputs)"
ch.x_axis.title = "Staples Strength Index (SSI)"
ch.y_axis.title = "Market Attractiveness Index (MAI)"
ch.x_axis.scaling.min, ch.x_axis.scaling.max = 20, 100
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 40, 100
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
smap.add_chart(ch, "B52")

# --------------------------------------------------------------- Recommended Paths
rp = sheet("Recommended Paths", "Recommended L3/L4 paths: Tier 1 (PoC wave 1) and Tier 2 (wave 2 / swap-ins)",
           "Rows are in default-score order. Segment, play, score, tier, rank and competitors are live lookups; the narrative reflects the research at default inputs.")
header(rp, 4, ["Rank", "Tier", "Hero segment", "Marketplace play", "Score", "Staples path (exact)", "Level", "Category ID", "Staples items",
               "SKUs in hand", "Fast-start", "Primary competitor", "Secondary competitor", "Why this path",
               "Archetypes we expect to find (hypotheses)", "1P watch-out (cannibalisation)", "Commercial moment",
               "Why the room will relate", "Where to look on the competitors", "Staples URL"],
       [6, 8, 13, 11, 7, 52, 6, 11, 9, 9, 8, 18, 18, 48, 48, 36, 26, 36, 44, 40])
for i, r in enumerate(short):
    rr = 5 + i
    put(rp, rr, 1, sc_lookup(C_RANK, f"$H{rr}"), align=center)
    put(rp, rr, 2, sc_lookup(C_TIER, f"$H{rr}"), font=f_bold, align=center)
    put(rp, rr, 3, sc_lookup(C_SEG, f"$H{rr}"), font=f_bold, align=center)
    put(rp, rr, 4, sc_lookup(C_PLAY, f"$H{rr}"), font=f_bold, align=center)
    put(rp, rr, 5, sc_lookup(C_SCORE, f"$H{rr}"), align=center, fmt="0.0")
    put(rp, rr, 6, r["path"], font=f_bold)
    put(rp, rr, 7, r["level"], align=center)
    put(rp, rr, 8, r["cid"], align=center)
    put(rp, rr, 9, r["items"], align=center, fmt="#,##0")
    put(rp, rr, 10, r["hand"], align=center, fmt="#,##0")
    put(rp, rr, 11, sc_lookup(C_FAST, f"$H{rr}"), align=center)
    put(rp, rr, 12, sc_lookup(C_PRI, f"$H{rr}"), font=f_bold)
    put(rp, rr, 13, sc_lookup(C_SEC, f"$H{rr}"))
    for j, v in enumerate(D[r["key"]]):
        put(rp, rr, 14 + j, v)
    put(rp, rr, 20, r["url"])
    rp.row_dimensions[rr].height = 90
RP_LAST = 4 + len(short)
rp.freeze_panes = "G5"
rp.auto_filter.ref = f"A4:T{RP_LAST}"
tier_cf(rp, f"B5:B{RP_LAST}", "B", 5)
seg_cf(rp, f"C5:C{RP_LAST}", "C", 5)

# --------------------------------------------------------------- Your List Review
yl = sheet("Your List Review", "Review of the initial 9-item list, with exact L3/L4 drill-downs and their Hero segments",
           "Table 1 = verdict per item; column I shows that most L2s mix Hero, Probable Hero and Non-Hero shelves, which is why segmenting at L2 misleads. Table 2 = the drill-down paths.")
review = [
 ("1", "Chairs & Seating", "Wayfair", "Furniture > Chairs & Seating", "KEEP: strongest item on the list",
  "Core-adjacent furniture with 14 leaves. The pilot already proved the method here (Accent & Waiting Room, Benches, Bar Stools). The 2026 'resimercial' return-to-office trend supports design-led seating.",
  "Wayfair is the right primary (depth + aesthetic, already scraped). Add Amazon as secondary for the recruitable seller pool and the value tier.",
  "Wayfair / Amazon"),
 ("2", "Desks", "Wayfair", "Furniture > Desks", "KEEP, narrow to Office Desks",
  "Office Desks is the real shelf (1,501 items) and splits cleanly by Staples' own type facet. Sit & Stand is spec-led (motors, frames), so design gaps are smaller and cannibalisation risk is higher.",
  "Wayfair primary for Office Desks. If Sit & Stand is kept, Amazon should lead it (FlexiSpot / VIVO-type sellers).",
  "Wayfair / Amazon"),
 ("3", "Lamps & Lighting", "Wayfair", "Furniture > Lamps & Lighting", "KEEP",
  "The purest aesthetic category in the core catalogue, and Staples SKUs are already in hand for desk, table and floor lamps.",
  "Wayfair is right. Amazon as secondary for Desk Lamps; Target for Table and Floor Lamps.",
  "Wayfair / Amazon or Target"),
 ("4", "Rugs", "Wayfair", "Decor > Rugs (L2 leaf, no L3/L4 under it)", "PARK: fails the granularity gate",
  "There is nothing to drill into: 'Decor > Rugs' is a single L2 leaf of 804 items. Rugs are also residential decor, the weakest link to Staples' work and school mission.",
  "Wayfair would be the right competitor if a rug story is ever needed. For a decor story at L3/L4, use Storage Baskets (Tier 1) or Faux Plants / Frames (Tier 2).",
  "n/a"),
 ("5", "Fitness Equipment", "Amazon", "Fitness > Fitness Equipment > Fitness Machines > Fitness Machines (the only leaf)", "REFOCUS as 'active workstation': Tier 3",
  "Fitness is not a Staples destination: one leaf, 43 items, and not in the header nav. The only defensible angle is under-desk walking pads and desk bikes sold next to sit-stand desks. Bulky, low attach.",
  "Amazon is the right primary; Dick's Sporting Goods as secondary.",
  "Amazon / Dick's"),
 ("6", "Backpacks & Bags", "Scheels", "Bags, Backpacks & Luggage > Backpacks & Laptop Bags (+ Totes & Handbags, Briefcases & Padfolios)", "KEEP the category, CHANGE the competitor",
  "A strong BTS anchor (85% of parents buy a backpack, Circana). But Scheels' depth is technical, hunting and outdoor packs, far from Staples' student and commuter shopper, so the method would mostly return OFF-BRAND archetypes. Scheels is also regional and blocks scraping.",
  "Amazon primary (breadth, recruitable sellers), Target secondary (lifestyle styles; JanSport sells via Target Plus). Keep Scheels only as an outdoor style reference.",
  "Amazon / Target"),
 ("7", "Hydration & Drinkware", "Scheels", "Coffee, Water & Snacks > Water & Beverages > Water Bottles, Tumblers & Travel Mugs", "KEEP: a top-3 path",
  "The Staples node is an L3 under Coffee, Water & Snacks (there is no 'Hydration' node). 463 items, already in the sample. The kitchen 'Drinkware & Glassware' leaf has only 15 items and fails the baseline gate.",
  "Dick's primary (same brand depth as Scheels, but national); Target secondary (exclusive colourways, closest shopper). Scheels is an acceptable alternate.",
  "Dick's / Target"),
 ("8", "Labels & Packaging Accessories", "Michaels", "Office Supplies > Labels (L2 leaf) · Party Supplies > Wrapping Supplies · Retail Store Supplies > Retail Packaging", "REFRAME",
  "'Labels' is an L2 leaf of 11,493 Avery-type labels: no L3/L4, commodity, core 1P. Gift wrap and gift bags now overlap the Party City partner assortment on Staples.com (Apr-2026). Michaels has little depth in business labels or small-business packaging.",
  "Reframe as small-business packaging & labelling: Amazon primary, Uline secondary. For Label Makers: Amazon + Michaels (Cricut).",
  "Amazon / Uline"),
 ("9", "Planning & Craft Organization", "Michaels", "Office Supplies > Calendars & Planners · Arts & Crafts > Scrapbooking / Crafting · Office Supplies > Storage & Organization", "KEEP, split in two",
  "Two different shopper missions in one bucket. Planning is the #1 path overall (Michaels' Happy Planner ecosystem vs Staples' dated planners). Craft organization is thin at Staples (66 items) and works better as general organization.",
  "Planning: Amazon and Michaels are a near-tie (Amazon edges it on seller recruitability): run both, with Michaels as the enthusiast-format benchmark. Organization: Target + Amazon, with Container Store as the style benchmark.",
  "Amazon / Michaels; Target / Amazon"),
]
header(yl, 4, ["#", "Your item", "Your competitor", "Exact Staples node(s)", "Verdict", "What the research says", "Competitor view",
               "Recommended primary / secondary", "Segment mix of its L3/L4 paths (live)"],
       [4, 22, 13, 46, 30, 70, 56, 24, 30])
for i, rv in enumerate(review):
    rr = 5 + i
    for j, v in enumerate(rv, 1):
        put(yl, rr, j, v, font=f_bold if j in (2, 5) else f_body, align=center if j == 1 else wrap)
    link = f"{rv[0]}. {rv[1]}"
    C_ = f"{HS_}$C${HS_FIRST}:$C${HS_LAST}"; Q_ = f"{HS_}$Q${HS_FIRST}:$Q${HS_LAST}"; S_ = f"{HS_}$S${HS_FIRST}:$S${HS_LAST}"
    cnt = lambda seg: f'COUNTIFS({C_},"{link}",{Q_},"{seg}",{S_},"<>Parked (gate)")'
    put(yl, rr, 9, f'="Hero "&{cnt("Hero")}&CHAR(10)&"Probable Hero "&{cnt("Probable Hero")}&CHAR(10)&"Non-Hero "&{cnt("Non-Hero")}'
                   f'&CHAR(10)&"Parked "&COUNTIFS({C_},"{link}",{S_},"Parked (gate)")')
    yl.row_dimensions[rr].height = 92
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
 ("8", ["Retail boxes", "Label makers", "Mailers", "Labels", "Gift boxes", "Gift wrap"]),
 ("9", ["Planners", "Journals", "Stickers", "Washi tape", "Calendars", "Calendar boards", "Craft storage", "Storage bins"]),
]
start = 5 + len(review) + 2
yl.cell(row=start - 1, column=1, value="Table 2. Drill-down L3/L4 paths for each item (exact nav-tree paths; tier, score and competitors are live)").font = f_sec
hdr = ["#", "Path ID", "Staples path (exact)", "Level", "Category ID", "Staples items", "Hero segment", "Marketplace play", "Tier", "Score",
       "Primary / Secondary competitor", "Role in PoC"]
for j, h in enumerate(hdr, 1):
    c = yl.cell(row=start, column=j, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
role_note = {"Office chairs": "Control shelf (Hero, PROTECT): shows the safety gate rejecting core-1P look-alikes",
             "Clocks": "Backlog: Non-Hero, capped at Tier 3", "Desk pads": "Backlog: Non-Hero, capped at Tier 3",
             "Rugs": "Parked: no L3/L4 exists", "Luggage": "Parked: leisure travel, off-mission",
             "Kitchen drinkware": "Parked: 15 items, no baseline", "Labels": "Parked: L2 leaf, commodity",
             "Gift boxes": "Parked: Party City partner overlap", "Gift wrap": "Parked: Party City partner overlap"}
rr = start + 1
for num, keys in drill:
    for k in keys:
        r = by_key[k]
        put(yl, rr, 1, num, align=center); put(yl, rr, 2, pid[k], align=center)
        put(yl, rr, 3, r["path"], font=f_bold); put(yl, rr, 4, r["level"], align=center)
        put(yl, rr, 5, r["cid"], align=center); put(yl, rr, 6, r["items"], align=center, fmt="#,##0")
        put(yl, rr, 7, sc_lookup(C_SEG, f"$E{rr}"), font=f_bold, align=center)
        put(yl, rr, 8, sc_lookup(C_PLAY, f"$E{rr}"), align=center)
        put(yl, rr, 9, sc_lookup(C_TIER, f"$E{rr}"), font=f_bold, align=center)
        put(yl, rr, 10, sc_lookup(C_SCORE, f"$E{rr}"), align=center, fmt="0.0")
        put(yl, rr, 11, f"=INDEX('Path Scorecard'!${C_PRI}${SC_FIRST}:${C_PRI}${SC_LAST},MATCH($E{rr},'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))&\" / \"&INDEX('Path Scorecard'!${C_SEC}${SC_FIRST}:${C_SEC}${SC_LAST},MATCH($E{rr},'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))")
        note = role_note.get(k) or {"Tier 1": "PoC wave 1", "Tier 2": "Wave 2 / swap-in", "Tier 3": "Backlog"}[r["tier"]]
        put(yl, rr, 12, note)
        rr += 1
seg_cf(yl, f"G{start + 1}:G{rr - 1}", "G", start + 1)
tier_cf(yl, f"I{start + 1}:I{rr - 1}", "I", start + 1)
yl.freeze_panes = "C5"

# --------------------------------------------------------------- Competitor Profiles
cp = sheet("Competitor Profiles", "Competitor profiles: why each one is (or is not) a benchmark",
           "K4 / K5 are the competitor-level scores used in 'Competitor Fit'. Access = what a plain HTTP request returned on 29-Sep-2026.")
profiles = [
 ("Wayfair (+ Wayfair Professional)", "Home & commercial furniture / decor pure-play", "Drop-ship model with 11,000+ suppliers and a 14M+ product catalogue (secondary sources); 21.4M active customers in Q1-2026.",
  "Very wide style range (boucle, rattan, mid-century, glam) at every price tier.", 3, "House-brand labels (Latitude Run, Ebern Designs...) hide the supplier, but the drop-ship suppliers behind them can be recruited.",
  4, "Already scraped for the pilot (chairs, desks, mats). Today: HTTP 429 (rate-limited).", "Seating, desks, lamps, storage baskets, clocks, faux plants, frames, desk pads, memo boards",
  "Convenience sample so far: compare shares, never counts. House brands are not sellers."),
 ("Amazon (+ Amazon Business)", "Everything marketplace; the B2B rival Staples meets daily", "3P sellers = ~60-61% of paid units (Q1-Q2 2026). Amazon Business: $60B+ annualised gross sales, 11M+ organisations (Amazon, per search summary).",
  "Huge variety, noisy quality; strongest for small low-ticket items.", 5, "The richest pool of sellers who already run multi-channel and could list on Staples Marketplace.",
  3, "Bot-protected to plain requests; mature paid APIs exist (e.g. the Rainforest API the Track-B plan named).", "Pencil cases, lunch, backpacks, desk mats, laptop stands, label makers, packaging, classroom decor",
  "Needs de-duplication of near-identical 3P listings; sample by facet, not bestseller."),
 ("Target (+ Target Plus)", "Mass lifestyle retailer; invite-only curated marketplace", "Target Plus: invite-only, 'curation at scale', adding dozens of hand-picked brands in summer 2026; JanSport is a named partner.",
  "Design-led owned brands (Brightroom, Threshold, Studio McGee) + exclusive colourways.", 3, "Owned brands are not recruitable; Target Plus brands are.",
  3, "Pages are JavaScript-rendered; data loads from the site's JSON endpoints (feasible with a headless scraper).", "Desk organizers, storage bins, lunch, pencil cases, backpacks, table & floor lamps",
  "Target Plus is the closest peer to Pat's curated-marketplace model: a strong talking point."),
 ("Michaels", "Arts, crafts, planning & framing specialist", "~200k online SKUs plus ~1.3M added via its Mirakl marketplace (eMarketer).",
  "Enthusiast depth: Happy Planner ecosystem, washi, stickers, frames, floral.", 3, "Marketplace sellers are recruitable; some brands are Michaels-exclusive.",
  3, "HTML returns 200 but product grids are JavaScript-rendered.", "Planners, stickers, washi tape, frames, faux plants, journals",
  "Hobby-pro items (e.g. Copic markers) can pull the analysis off-mission: keep to planning & display."),
 ("The Container Store", "Organization specialist", "Reset in 2026: liquidating ~30% of select categories / SKUs; co-branded stores with Bed Bath & Beyond (Beyond Inc).",
  "Coordinated acrylic, bamboo and marble collections; the organization style benchmark.", 2, "Much is own brand (Elfa, TCS); some third-party brands (Yamazaki, Russell+Hazel).",
  2, "HTTP 307 redirect to a bot check.", "Desk organizers, storage baskets & bins",
  "Shrinking assortment in 2026: use as a style reference rather than the primary depth source."),
 ("Best Buy", "Consumer electronics specialist", "Examples doc: keyboards, audio, gaming at enthusiast depth; also runs a Mirakl marketplace.",
  "Colour and style options within well-known tech brands.", 3, "National tech brands; marketplace sellers.",
  4, "Official developer Products API (key required); the website itself did not answer a plain request.", "Keyboards, laptop stands, desk mats, headphones",
  "Tech items carry higher cannibalisation risk against Staples' own tech range."),
 ("Dick's Sporting Goods", "National sporting goods", "Stocks YETI, HydroJug, Owala, Stanley and CamelBak with a dedicated BTS hydration page.",
  "Full colourways and sizes of the trending bottle brands.", 4, "National brands that already sell wholesale are recruitable.",
  2, "HTTP 403 ('Site Unavailable').", "Water bottles & tumblers; secondary for backpacks and fitness",
  "Sports-family shopper; overlap with Staples' office shopper is partial."),
 ("Scheels", "Regional sporting goods specialist", "Examples doc: 'wall-to-wall' hydration and technical packs (Osprey, Mystery Ranch, North Face).",
  "Deep in outdoor and technical.", 4, "Same national brands as Dick's.",
  2, "HTTP 403.", "Alternate for water bottles; style reference for outdoor packs",
  "Regional footprint and technical skew make it a weaker primary than Dick's for Staples' shopper."),
 ("Uline", "B2B packaging & facilities catalogue", "45,000+ products in stock (uline.com).",
  "Utility-first, but many colour and size options in mailers, retail bags and gift boxes.", 2, "Largely Uline-branded: few recruitable sellers.",
  4, "HTTP 200: server-rendered HTML.", "Retail boxes & bags, mailers",
  "Good depth benchmark, weak seller pool: pair with Amazon."),
 ("JetPens", "Online stationery specialist (Japanese & European)", "10,000+ products (jetpens.com, About Us).",
  "Design-led stationery; the benchmark for 'aesthetic' pens and pouches.", 3, "Brands sold via US distributors (Pilot, Uni, Kokuyo...).",
  2, "HTTP 403 (Cloudflare).", "Style reference for pencil cases, washi, journals",
  "Enthusiast skew: use as a style reference, not the primary."),
 ("Oriental Trading", "Party, craft & classroom catalogue", "A major classroom-decor and teacher catalogue.",
  "Themed classroom sets (boho, calm, pastel).", 2, "Mostly own-sourced.",
  2, "HTTP 403 (security verification).", "Classroom decor (secondary)", "Weak seller pool; style and theme reference."),
 ("Office Depot / ODP (not used as primary)", "Direct like-for-like competitor", "Assortment mirrors Staples' utility range, with a marketplace on officedepot.com.",
  "Utility-first.", None, "", None, "Not tested.", "Sanity check only", "Little gap signal: comparing like with like finds few design-led archetypes."),
]
header(cp, 4, ["Competitor", "Type", "Scale / depth evidence", "Aesthetic lens", "K4 Recruitable", "Seller notes", "K5 Data access",
               "Access test (29-Sep-2026)", "Best-fit Staples paths", "Cautions"], [26, 24, 48, 36, 10, 36, 10, 34, 36, 40])
for i, p in enumerate(profiles):
    rr = 5 + i
    for j, v in enumerate(p, 1):
        put(cp, rr, j, v, font=f_bold if j == 1 else f_body, align=center if j in (5, 7) else wrap)
    cp.row_dimensions[rr].height = 78
cp.freeze_panes = "B5"

# --------------------------------------------------------------- Parked & Excluded
pk = sheet("Parked & Excluded", "Paths considered and not recommended (with the reason)",
           "Gate failures are also in the scorecard. Tier 3 paths stay in the backlog and can be promoted by changing weights or cut-offs.")
header(pk, 4, ["Path ID", "Staples path (exact)", "Category ID", "Staples items", "Hero segment", "Marketplace play", "Status", "Reason",
               "Revisit when"], [8, 64, 11, 10, 13, 11, 16, 70, 40])
revisit = {"Rugs": "A pseudo-L3 split by rug type is accepted as an exception", "Chair mats": "Same: only with a pseudo-L3 split",
           "Labels": "Never as a design play; commodity", "Gift boxes": "The Party City partnership scope changes",
           "Gift wrap": "The Party City partnership scope changes", "Kitchen drinkware": "Staples adds a baseline (mugs)",
           "Gift stationery": "Pseudo-L3 split", "Patio furniture": "Never; mis-shelved node", "Luggage": "Travel becomes a marketplace theme",
           "Office chairs": "Keep as the control shelf in the PoC", "Pens": "Only for a separate premium-pen (JetPens) study",
           "Headphones": "BTS tech attach is prioritised", "Sit & stand": "Paired with walking pads as a 'wellness desk' story",
           "Fitness machines": "Paired with Sit & Stand Desks", "Clocks": "Market signal for decor clocks strengthens",
           "Desk pads": "Merged with desk mats (Mouse Pads & Wrist Rests)"}
rr = 5
for r in [r for r in order if not r["gate"]] + [r for r in order if r["tier"] == "Tier 3"]:
    put(pk, rr, 1, pid[r["key"]], align=center); put(pk, rr, 2, r["path"]); put(pk, rr, 3, r["cid"], align=center)
    put(pk, rr, 4, r["items"], align=center, fmt="#,##0")
    put(pk, rr, 5, sc_lookup(C_SEG, f"$C{rr}"), font=f_bold, align=center)
    put(pk, rr, 6, sc_lookup(C_PLAY, f"$C{rr}"), align=center)
    put(pk, rr, 7, sc_lookup(C_TIER, f"$C{rr}"), font=f_bold, align=center)
    put(pk, rr, 8, r["why"]); put(pk, rr, 9, revisit.get(r["key"], "Weights shift toward its strengths"))
    rr += 1
put(pk, rr, 1, "-", align=center); put(pk, rr, 2, "Furniture > Decor > Wall Art > Wall Art/Decor (and sister Wall Art leaves)")
put(pk, rr, 3, "CL140795", align=center); put(pk, rr, 4, int(tree.loc["CL140795", "Count"]), align=center, fmt="#,##0")
put(pk, rr, 5, "-", align=center); put(pk, rr, 6, "-", align=center)
put(pk, rr, 7, "Not scored", font=f_bold, align=center)
put(pk, rr, 8, "Staples already lists 10,000+ wall-art items (drop-ship catalogue): the gap is already closed."); put(pk, rr, 9, "Never")
tier_cf(pk, f"G5:G{rr}", "G", 5)
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
ex["A1"] = "Staples Marketplace PoC: which L3/L4 shelves to analyse, and against whom"
ex["A1"].font = f_title
ex["A2"] = "v2, prepared 29-Sep-2026 for Sai (LatentView). Paths are exact nodes from the Staples navigation tree (snapshot 21-Sep-2026)."
ex["A2"].font = f_sub
for col, w in zip("ABCDEFGHIJ", [5, 56, 8, 7, 22, 18, 18, 42, 28, 9]):
    ex.column_dimensions[col].width = w
n_t1 = sum(r["tier"] == "Tier 1" for r in rows); n_t2 = sum(r["tier"] == "Tier 2" for r in rows)
t1 = [r for r in short if r["tier"] == "Tier 1"]
n_h = sum(r["seg"] == "Hero" for r in t1); n_ph = sum(r["seg"] == "Probable Hero" for r in t1)
msgs = [
 f"Recommendation: run the PoC on {n_t1} Tier-1 L3/L4 shelves across 5 Staples L1s, benchmarked against 6 competitors. A further {n_t2} Tier-2 shelves are ready as swap-ins.",
 "Two lenses, one answer. Lens 1, Hero segmentation: is the shelf a Staples Hero (its own strength), a Probable Hero (a market favourite where Staples is not yet a destination) or a Non-Hero? Lens 2, a path score: 8 criteria led by Pat's 'White Chair' test, competitor depth and 1P protection.",
 f"Tier 1 mixes both growth stories: {n_h} Heroes to EXTEND with style variants and {n_ph} Probable Heroes to BUILD with marketplace sellers. Non-Heroes are capped at Tier 3, and Heroes where variants would substitute 1P (Office Chairs, Pens) are marked PROTECT.",
 "Every shelf first passed four hard gates: a true L3/L4 in the Staples tree, on Staples' work / school / small-business mission, 20+ Staples items as a baseline, and no clash with a Staples partner (Party City now supplies party décor and gift wrap on Staples.com).",
 "Six Tier-1 shelves already have Staples SKUs in hand ('Fast-start'). Competitors are picked per shelf; Amazon wins most secondaries because its sellers are the ones Staples can actually recruit.",
 "Caveats: no internal Staples sales or margin data exists, so 'Hero' uses external proxies (depth, header-nav visibility, merchandising investment, shopper association); swap in Staples' own sales rank when Pat shares it. Competitor depth is scored from desk research; confirm it when the listings are collected.",
]
for i, m in enumerate(msgs):
    c = ex.cell(row=4 + i, column=1, value="•"); c.font = f_bold
    c = ex.cell(row=4 + i, column=2, value=m); c.font, c.alignment = f_body, wrap
    ex.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=9)
    ex.row_dimensions[4 + i].height = 34
r0 = 4 + len(msgs) + 1
ex.cell(row=r0, column=1, value="Tier 1: PoC wave 1").font = f_sec
LK = lambda col, cid: f"=INDEX('Path Scorecard'!${col}${SC_FIRST}:${col}${SC_LAST},MATCH(\"{cid}\",'Path Scorecard'!$H${SC_FIRST}:$H${SC_LAST},0))"
for j, h in enumerate(["#", "Staples path (exact L3/L4)", "Score", "Level", "Hero segment · play", "Primary competitor", "Secondary competitor",
                       "The story for the room", "Commercial moment", "Fast-start"], 1):
    c = ex.cell(row=r0 + 1, column=j, value=h); c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
for i, r in enumerate(t1):
    rr = r0 + 2 + i
    put(ex, rr, 1, i + 1, align=center)
    put(ex, rr, 2, r["path"], font=f_bold)
    put(ex, rr, 3, LK(C_SCORE, r["cid"]), align=center, fmt="0.0")
    put(ex, rr, 4, f"L{r['level']}", align=center)
    put(ex, rr, 5, LK(C_SEG, r["cid"]) + "&\" · \"&" + LK(C_PLAY, r["cid"])[1:], font=f_bold, align=center)
    put(ex, rr, 6, LK(C_PRI, r["cid"]), font=f_bold)
    put(ex, rr, 7, LK(C_SEC, r["cid"]))
    put(ex, rr, 8, D[r["key"]][4])
    put(ex, rr, 9, D[r["key"]][3])
    put(ex, rr, 10, LK(C_FAST, r["cid"]), align=center)
    ex.row_dimensions[rr].height = 44
    ex.cell(row=rr, column=5).fill = PatternFill("solid", fgColor=SEG_FILL[r["seg"]])
TR = f"'Path Scorecard'!${C_TIER}${SC_FIRST}:${C_TIER}${SC_LAST}"; SR = f"'Path Scorecard'!${C_SEG}${SC_FIRST}:${C_SEG}${SC_LAST}"
rr = r0 + 2 + len(t1)
put(ex, rr, 2, "Tier-1 mix (live)", font=f_bold)
c = ex.cell(row=rr, column=3, value=(f'=COUNTIFS({TR},"Tier 1",{SR},"Hero")&" Hero  ·  "&COUNTIFS({TR},"Tier 1",{SR},"Probable Hero")&" Probable Hero  ·  "'
                                     f'&COUNTIFS({TR},"Tier 1",{SR},"Non-Hero")&" Non-Hero"'))
c.font = f_bold
ex.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=9)
r1 = rr + 2
ex.cell(row=r1, column=1, value="What changed versus the initial 9-item list").font = f_sec
changes = [
 "Kept and drilled down: Chairs & Seating, Desks, Lamps & Lighting (Wayfair stays primary), Hydration (now the exact L3 'Water Bottles, Tumblers & Travel Mugs') and Planning (Planners is the #1 shelf overall).",
 "Hero segmentation moved from L2 to L3/L4. Most of the original L2s hold more than one segment: Chairs & Seating has a Hero to protect (Office Chairs), Probable Heroes to build (Accent, Bar Stools, Benches, Breakroom, Ottomans) and Non-Heroes (Kids, Gaming). See 'Your List Review' column I and the roll-up in 'Hero Segmentation'.",
 "Competitor changed: Backpacks & Bags moves from Scheels to Amazon + Target (Scheels skews technical / outdoor, away from Staples' student and commuter shopper); Hydration moves to Dick's + Target (same brands as Scheels, national reach).",
 "Parked: Rugs (no L3/L4 under 'Decor > Rugs') and Labels (an L2 leaf of commodity labels, and a Staples Hero to protect). Reframed: 'Labels & Packaging' becomes small-business packaging (Retail Boxes & Bags, Mailers, Label Makers vs Amazon + Uline); gift wrap and gift bags are gated out because of the Party City partnership.",
 "Refocused: Fitness becomes an 'active workstation' idea (walking pads next to sit-stand desks); a Probable Hero, but Tier 3 on path score.",
 "New shelves the list missed: Desk Organizers, Pencil Cases & Pouches, Classroom Decor, Storage Baskets (all Tier 1), plus desk-setup items and office decor in Tier 2.",
 "v2 vs v1: added the Hero lens. Clocks and Desk Pads drop from Tier 2 to Tier 3 because they are Non-Hero; Tier 2 now holds 23 shelves.",
]
for i, m in enumerate(changes):
    rr = r1 + 1 + i
    ex.cell(row=rr, column=1, value="•").font = f_bold
    c = ex.cell(row=rr, column=2, value=m); c.font, c.alignment = f_body, wrap
    ex.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=9)
    ex.row_dimensions[rr].height = 34
r2 = r1 + 1 + len(changes) + 1
ex.cell(row=r2, column=1, value="Suggested sequencing").font = f_sec
seq = [
 "Wave 1a (Staples data already in hand): Accent & Waiting Room Chairs, Desk Lamps, Planners, Water Bottles, Backpacks, Lunch Bags & Boxes: 2 Heroes (extend) and 4 Probable Heroes (build).",
 "Wave 1b (needs a Staples leaf scrape first): Desk Organizers, Pencil Cases & Pouches, Classroom Decor, Storage Baskets.",
 "Keep Office Chairs as a control shelf (Hero, PROTECT): it shows the safety gate correctly rejecting look-alikes of core 1P chairs.",
 "Before the run: collect competitor listings per 'Competitor Fit' (primary first), confirm the depth scores, and ask Pat for a Staples sales / margin rank to replace the Hero proxies.",
]
for i, m in enumerate(seq):
    rr = r2 + 1 + i
    ex.cell(row=rr, column=1, value=f"{i + 1}.").font = f_bold
    c = ex.cell(row=rr, column=2, value=m); c.font, c.alignment = f_body, wrap
    ex.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=9)
    ex.row_dimensions[rr].height = 30
r3 = r2 + 1 + len(seq) + 1
ex.cell(row=r3, column=1, value="How to read this workbook").font = f_sec
guide = [
 ("Segment Map", "One picture: every shelf placed as Hero, Probable Hero or Non-Hero."),
 ("Your List Review", "Verdict on each of the 9 initial items, their segment mix, and exact L3/L4 drill-downs."),
 ("Recommended Paths", "Tier 1 + Tier 2 shelves with segment, play, archetype hypotheses, 1P watch-outs, moments and where to look."),
 ("Selection Criteria", "Gates, weights, segmentation proxies, cut-offs and plays (editable yellow cells)."),
 ("Path Scorecard", f"All {len(rows)} candidate shelves with gates, segment, scores, rank, tier and play (formulas)."),
 ("Hero Segmentation", "The two indices behind each segment, plus an L2 roll-up."),
 ("Competitor Fit / Competitor Profiles", "Why each primary and secondary competitor was chosen."),
 ("Parked & Excluded", "What was considered and dropped, and why."),
 ("Evidence & Sources", "Every fact with its source, date and reliability."),
 ("Staples Tree Extract", "All terminal Staples nodes, for checking any path."),
]
for i, (a, b) in enumerate(guide):
    rr = r3 + 1 + i
    put(ex, rr, 2, a, font=f_bold)
    c = ex.cell(row=rr, column=3, value=b); c.font = f_body
    ex.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=9)
lg = r3 + 2 + len(guide)
put(ex, lg, 2, "Legend", font=f_bold)
c = ex.cell(row=lg, column=3, value=("Blue text on yellow = editable input. Tiers: green Tier 1, amber Tier 2, grey Tier 3, red Parked. "
                                     "Segments: blue Hero, orange Probable Hero, grey Non-Hero."))
c.font = f_body
ex.merge_cells(start_row=lg, start_column=3, end_row=lg, end_column=9)

del wb["Sheet"]
order_names = ["Executive Summary", "Segment Map", "Your List Review", "Recommended Paths", "Selection Criteria", "Path Scorecard",
               "Hero Segmentation", "Competitor Fit", "Competitor Profiles", "Parked & Excluded", "Evidence & Sources", "Staples Tree Extract"]
wb._sheets = [wb[n] for n in order_names]
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {"Executive Summary": NAVY, "Segment Map": "2A78D6", "Your List Review": TEAL,
                                    "Recommended Paths": "70AD47", "Hero Segmentation": "2A78D6"}.get(ws.title, "A6A6A6")
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)

print("saved", OUT)
print("tiers:", pd.Series([r["tier"] for r in rows]).value_counts().to_dict())
for r in short:
    print(f"{r['tier']:6} {r['s']:5.1f} {r['seg']:13} {r['play']:7} {roles[r['key']][0]:22} {roles[r['key']][1]:18} {r['path']}")
