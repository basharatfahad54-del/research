import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.Workbook()

NAVY = "1F3864"
BLUE = "2E75B6"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
AMBER = "FFF2CC"
RED = "F8CBAD"
GREEN = "C6E0B4"

hdr_font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
hdr_fill = PatternFill("solid", fgColor=NAVY)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_sheet(ws, widths, freeze="A2", height=42, wrap_cols=None):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = border
    ws.row_dimensions[1].height = 45
    ws.freeze_panes = freeze
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(widths)):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = border
            c.font = Font(size=9, name="Calibri")
        ws.row_dimensions[row[0].row].height = height
    ws.auto_filter.ref = ws.dimensions

# ---------------- Sheet 0: README ----------------
ws = wb.active
ws.title = "README"
ws["A1"] = "LITERATURE MATRIX — PhD: Developing an Innovative Systemic Living Lab for Just Energy Transition in a South African Informal Settlement"
ws["A1"].font = Font(bold=True, size=13, color=NAVY)
ws["A2"] = "Compiled 12 November 2026 | Phase 1 deliverable | University of Johannesburg (institution to be confirmed)"
ws["A2"].font = Font(italic=True, size=10)

rows = [
    ("", ""),
    ("PURPOSE", "Records, for every source, what it says, what it is good for, where it is weak, which research question it serves, and whether it has been verified. It is a working instrument, not a bibliography."),
    ("", ""),
    ("SHEETS", ""),
    ("Sources", "The matrix itself: 46 sources spanning theory, energy justice, Living Labs, systems thinking and the South African context."),
    ("Gap map", "Each identified gap, its type, the RQ it serves, and the chapter where it is addressed."),
    ("RQ-theory map", "Which concepts from which sources do analytical work for each research question."),
    ("Verification queue", "What must be checked before a source may be cited, in priority order."),
    ("", ""),
    ("HONESTY RULES APPLIED", ""),
    ("Rule 1", "No title, journal, volume, page range or DOI is asserted unless it was seen in a document in the workspace or is a work whose details the compiler is confident of. Where uncertain, the field reads 'TBC' (to be confirmed)."),
    ("Rule 2", "'Read depth' distinguishes: Full = read end to end in this workspace; Partial = read in substantial part; Secondary = known only through another document's citation; Canon = foundational work known to the compiler but not yet re-read against the original text in this project."),
    ("Rule 3", "A source marked 'Secondary' or 'Canon' may inform planning but must NOT be quoted or closely paraphrased until it has been obtained and read. This is recorded in the Verification queue."),
    ("Rule 4", "No reference has been entered into EndNote. No PDF has been downloaded. No DOI has been resolved. Those are Phase 6 tasks."),
    ("", ""),
    ("COLOUR KEY", ""),
    ("Green", "Verified and read in this project"),
    ("Amber", "Identified and relevant, but not yet verified/read in full"),
    ("Red", "Cited in the existing draft with insufficient or corrupted bibliographic detail — high risk"),
]
r = 3
for k, v in rows:
    ws.cell(row=r, column=1, value=k).font = Font(bold=True, size=10, color=NAVY)
    ws.cell(row=r, column=2, value=v).font = Font(size=10)
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    if k in ("Green", "Amber", "Red"):
        fill = {"Green": GREEN, "Amber": AMBER, "Red": RED}[k]
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=fill)
    r += 1
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 150
for i in range(1, r):
    ws.row_dimensions[i].height = 32 if i > 2 else 22

# ---------------- Sheet 1: Sources ----------------
ws = wb.create_sheet("Sources")
headers = [
    "ID", "Author(s)", "Year", "Title", "Type", "Outlet / Publisher", "DOI / identifier",
    "Read depth", "Verification", "Context / geography", "Theory / concept", "Method",
    "Key argument relevant to this study", "RQ", "Limitation / critique", "Use in thesis", "Priority to obtain"
]
ws.append(headers)

S = [
 # id, authors, year, title, type, outlet, doi, depth, verif, context, concept, method, argument, rq, limitation, use, priority
 ("L01","Aboalez, K.","2025","A Posthumanist Social Approach to Energy Transition","Conference paper","Procedia CIRP 135, 8–13","10.1016/j.procir.2025.01.042","Full","Verified in workspace (PDF, open access)","Global / engineering education","Critical posthumanism; social justice in engineering; posthumanist cartography","Conceptual / review","Engineering is constituted in tension with social justice; critical posthumanism supplies an ethics of engagement and a cartographic method for mapping power in energy assemblages.","RQ1, RQ4","Six pages; Braidotti/Barad only (no Haraway/Hayles); normative rather than analytical; education-focused, transfer to settlements asserted not shown.","Primary anchor for Ch3; supervisor's position; cite for posthumanist cartography and principles.","Obtained"),
 ("L02","Molepo, P.M., Aboalez, K. & Mathaba, T.N.D.","2025","Analysis of Barriers to South Africa's Energy Transition: Perspectives from industry experts","Journal article","Energy for Sustainable Development 88, 101777","10.1016/j.esd.2025.101777","Full","Verified in workspace (PDF, open access)","South Africa","AKTESP barrier framework (Trudgill 1990); energy transition barriers","15 semi-structured expert interviews; quantitative + qualitative content analysis; saturation at 9","Agreement, social and technological barriers dominate; grid connection and policy continuity cited by 100% of experts; barriers are multi-dimensional and interacting.","RQ4","Firm/expert-level, not household; no justice framing; supply-side only.","Ch1 context; Ch2 barriers; Ch4 mechanism design; methodological precedent for barrier analysis.","Obtained"),
 ("L03","Braidotti, R.","2013","The Posthuman","Book","Polity Press, Cambridge","TBC (ISBN to verify)","Canon","Cited in D-01; edition to verify","Continental philosophy","Nomadic subject; posthuman knowledge; affirmative ethics; critique of humanism; potestas/potentia","Philosophical","Humanism's 'Man' is a historically specific and exclusionary figure; critical posthumanism affirms difference and relational becoming rather than transcendence.","RQ1, RQ2","Abstract; Eurocentric reference points; little on infrastructure, households or empirical method.","Ch3 §§3.3, 3.4, 3.6, 3.11.","High"),
 ("L04","Braidotti, R.","2019","Posthuman Knowledge","Book","Polity Press, Cambridge","TBC","Canon","Cited in D-01; to verify","Knowledge production / university","Posthuman knowledge; socially relevant knowledge; academic responsibility","Philosophical","Knowledge production is situated and materially mediated; posthuman scholarship should be socially relevant and affirmative of difference.","RQ1, RQ5","Normative; limited operational guidance.","Ch3 §§3.3, 3.4, 3.13.","High"),
 ("L05","Haraway, D.","1988","Situated Knowledges: The Science Question in Feminism and the Privilege of Partial Perspective","Journal article","Feminist Studies 14(3), 575–599","TBC","Canon","Not in D-01; verify","Feminist epistemology","Situated knowledge; partial perspective; the 'god trick'","Philosophical","All knowledge is produced from somewhere; the claim to a view from nowhere is an epistemological and political manoeuvre.","RQ2, RQ3","Not an energy text; requires translation.","Ch3 §§3.3, 3.4, 3.7 (end-user knowledge).","High"),
 ("L06","Haraway, D.","2016","Staying with the Trouble: Making Kin in the Chthulucene","Book","Duke University Press, Durham","TBC","Canon","Not in D-01; verify","More-than-human / Anthropocene","Sympoiesis; response-ability; making kin; Chthulucene","Philosophical","Response-ability is relational and requires staying with problems rather than resolving them technocratically.","RQ1, RQ4","Dense; metaphorical vocabulary; limited direct energy application.","Ch3 §§3.3, 3.6, 3.8.","High"),
 ("L07","Barad, K.","2007","Meeting the Universe Halfway: Quantum Physics and the Entanglement of Matter and Meaning","Book","Duke University Press, Durham","TBC","Canon","Cited in D-01; verify","Philosophy of science / new materialism","Agential realism; intra-action; apparatus; phenomenon; agential cut; diffraction","Philosophical","Relata do not pre-exist their relations; boundaries and properties emerge through intra-actions within phenomena; the apparatus is part of what it measures.","RQ1, RQ2, RQ5","Demanding vocabulary; risk of mechanical application; downplays political economy.","Ch3 §§3.3, 3.4, 3.7, 3.8 (Living Lab as apparatus).","High"),
 ("L08","Barad, K.","2003","Posthumanist Performativity: Toward an Understanding of How Matter Comes to Matter","Journal article","Signs 28(3), 801–831","TBC","Canon","Verify","Feminist theory / new materialism","Performativity; matter; posthumanist performativity","Philosophical","Matter is agentive and not a passive surface on which discourse is inscribed.","RQ3","Abstract; requires careful exposition.","Ch3 §3.3.5 (materiality).","High"),
 ("L09","Hayles, N.K.","1999","How We Became Posthuman: Virtual Bodies in Cybernetics, Literature, and Informatics","Book","University of Chicago Press, Chicago","TBC","Canon","Verify","Cybernetics / literary theory","Informational vs embodied posthuman; technogenesis; the posthuman as material-informational","Philosophical / historical","The posthuman must not be reduced to disembodied information; embodiment and materiality remain constitutive.","RQ3","Less developed as justice theory; dated in places.","Ch3 §3.4 (corrective to abstraction); §3.7 (embodiment).","High"),
 ("L10","Wolfe, C.","2010","What Is Posthumanism?","Book","University of Minnesota Press, Minneapolis","TBC","Canon","Verify","Posthumanist theory","Posthumanism as anti-humanism; the 'posthumanist' vs 'transhumanist' distinction","Philosophical","Distinguishes critical posthumanism from transhumanist enhancement; posthumanism is a critique of humanism's exclusions.","RQ1","Taxonomic rather than constructive.","Ch3 §3.3.2 (differentiating posthumanism from transhumanism).","Medium"),
 ("L11","Forlano, L.","2017","Posthumanism and Design","Journal article","She Ji 3(1), 16–29","TBC","Canon","Cited in D-01; verify","Design theory","Posthuman design; human and non-human agency in design","Conceptual","Human and non-human actors share agency in shaping design outcomes; design must account for assemblages.","RQ2, RQ3","Design discipline focus; not energy-specific.","Ch3 §3.7 (design/participation).","High"),
 ("L12","Ulmer, J.B.","2017","Posthumanism as research methodology: inquiry in the Anthropocene","Journal article","International Journal of Qualitative Studies in Education 30(9), 832–848","TBC","Canon","Cited in D-01; verify","Qualitative methodology","Posthumanist methodology; justice as material and ecological","Methodological / conceptual","Posthumanism implies methodological consequences: research is an intravention, not an observation; justice extends beyond human relations.","RQ1, RQ5","Education focus; general.","Ch3 §3.3.7; Ch5 (methodology reconciliation).","High"),
 ("L13","Musango, J.K., Smit, S., Ceschin, F., Ambole, A., Batinge, B., Anditi, C., Petrulaityte, A. & Mukama, M.","2020","Mainstreaming gender to achieve security of energy services in poor urban environments","Journal article","Energy Research & Social Science 70, 101715","10.1016/j.erss.2020.101715","Secondary","DOI verified via Crossref; full text not yet read","South Africa / poor urban environments","Gender mainstreaming; energy security; energy services","Conceptual / review","Energy insecurity in poor urban environments is gendered and requires mainstreaming in planning, policy and innovation.","RQ1, RQ2, RQ3","Full text must be read before quotation; gender-framed, so requires careful translation under the end-user instruction.","Ch1 context; Ch2; Ch4 differentiation of end-users.","Critical"),
 ("L14","Mukama, M., Musango, J.K., Smit, S., Ceschin, F. & Petrulaityte, A.","2022","Development of living labs to support gendered energy technology innovation in poor urban environments","Journal article","Technology in Society 68, 101850","10.1016/j.techsoc.2021.101850","Secondary","DOI verified via Crossref; full text not yet read","Africa / poor urban environments","Living Lab development; energy technology innovation; gendered perspective","Case-based / conceptual","Develops Living Labs to support gendered energy technology innovation; the closest prior study to this project and therefore the key novelty reference.","RQ1, RQ5","Must be read in full to establish the precise novelty gap; published before the posthumanist turn.","Ch1 novelty claim; Ch2; Ch3 §3.8.","Critical"),
 ("L15","Anditi, C., Musango, J.K., Smit, S. & Ceschin, F.","2022","Addressing gender dimensions in energy innovations: A gender analysis framework for informal urban settlements in Africa","Journal article","Energy Research & Social Science 88, 102476","10.1016/j.erss.2021.102476","Secondary","DOI verified via Crossref; full text not yet read","Africa / informal urban settlements","Gender analysis framework; energy innovation","Framework development","Gender-blind innovation processes help explain low uptake in informal urban settlements; proposes a gender analysis framework.","RQ2, RQ3","Gender-framed; full text needed before use.","Ch1; Ch2; Ch4.","Critical"),
 ("L16","Petrulaityte, A., Ceschin, F., Musango, J.K., Mwiti, E., Anditi, C. & Njoroge, P.","2022","Supporting the Development of Gendered Energy Innovations for Informal Urban Settlements: GENS Codesign Toolkit for Multistakeholder Collaboration","Journal article","Sustainability 14, 6291","10.3390/su14106291","Secondary","DOI verified via Crossref; open access, full text not yet read","Kenya / South Africa (GENS project)","Co-design toolkit; multistakeholder collaboration; energy innovation","Toolkit development / case study","Provides a codified co-design toolkit (GENS) for multistakeholder energy innovation in informal settlements — the direct methodological precursor to the proposed Living Lab.","RQ2, RQ4, RQ5","Gender-framed; must be read to position the present study's posthumanist and justice contribution.","Ch2; Ch4; Ch5 (co-design method).","Critical"),
 ("L17","Okoye, A.","2020","TBC (Systems Reinforcing Model for Energy Access Sustainability)","Thesis / publication","TBC","TBC","Secondary","Cited 28× in draft; page numbers suggest a thesis","Nigeria / Africa","SRM; Energy Cultures; Sustainable Energy Access","TBC","Four system interconnections model energy access sustainability; used as the draft's core analytic scaffold.","RQ1","Secondary; the draft's use of SRM must be checked against the original.","Ch1 §1.2.1; Ch4 (model architecture).","Critical"),
 ("L18","Kovacic, Z. et al.","2019","TBC","Journal article","TBC","TBC","Secondary","Cited 9× in draft","Africa / informal settlements","Gender; energy poverty; informal settlements","TBC","Informal settlements are gendered; women bear disproportionate energy-poverty burdens.","RQ3","Secondary.","Ch2 (differentiation); Ch4.","High"),
 ("L19","Callaghan, R. & Herselman, M.","2015","TBC (Living Lab in education / ICT)","Journal article / chapter","TBC","TBC","Secondary","Cited 13× in draft","South Africa","Living Lab; education; ICT","TBC","Living Lab as methodology for co-creation in resource-constrained settings.","RQ1, RQ5","Secondary.","Ch2; Ch4 (SETILL structure adapted from this source).","High"),
 ("L20","Stephenson, J. et al.","2010","Energy cultures: A framework for understanding energy behaviours","Journal article","Energy Policy 38(10), 6120–6129","TBC","Canon","Cited 7× in draft; verify","New Zealand / general","Energy Cultures: material culture, cognitive norms, energy practices","Conceptual / empirical","Energy behaviour is shaped by the interaction of material culture, norms and practices — a relational alternative to individualist models.","RQ3","Not posthumanist; still practice-oriented and compatible.","Ch3 §3.5; Ch4 component 4.","High"),
 ("L21","Shrestha, R.M. & Acharya, J.","2015","TBC (Sustainable Energy Access framework)","Report / article","TBC","TBC","Secondary","Cited 7× in draft","Asia / global","SEA framework; sustainability pillars","TBC","Energy access sustainability assessed across economic, environmental, social, institutional and technical pillars.","RQ1","Secondary.","Ch1; Ch4.","Medium"),
 ("L22","Sovacool, B.K. & Dworkin, M.H.","2015","Energy justice: Conceptual insights and practical applications","Journal article","Applied Energy 142, 435–444","TBC","Canon","Verify","Global","Energy justice principles: availability, affordability, due process, good governance, sustainability, intra/intergenerational equity","Conceptual","Eight energy-justice principles derived from justice theory; justice must be applied across the energy lifecycle.","RQ1, RQ4","Anthropocentric; weak on design practice.","Ch3 §3.6; Ch4 component 10.","High"),
 ("L23","Jenkins, K. et al.","2016","Energy justice: A conceptual review","Journal article","Energy Research & Social Science 11, 174–182","TBC","Canon","Verify","Global","Distributive, procedural, recognition justice","Review","Synthesises energy justice into three core tenets; identifies the need for application to practice.","RQ4","Review; limited Global South detail.","Ch3 §3.6.","High"),
 ("L24","Walker, G. & Day, R.","2012","Fuel poverty as injustice: integrating distribution, recognition and procedure in the struggle for affordable warmth","Journal article","Energy Policy 49, 69–75","TBC","Canon","Verify","UK","Fuel poverty; distribution, recognition, procedure","Conceptual","Fuel poverty is a tripartite injustice; recognition and procedure matter alongside distribution.","RQ1, RQ4","Global North context.","Ch3 §3.6; Ch4.","High"),
 ("L25","Geels, F.W.","2002","Technological transitions as evolutionary reconfiguration processes: a multi-level perspective and a case-study","Journal article","Research Policy 31(8–9), 1257–1274","TBC","Canon","Verify","Netherlands / historical","Multi-level perspective: niche, regime, landscape","Conceptual / case study","Transitions occur through alignment and reconfiguration across levels, not through technology substitution alone.","RQ1, RQ5","Structuralist bias; weak on household agency.","Ch3 §3.9 (socio-technical complement).","High"),
 ("L26","Geels, F.W.","2011","The multi-level perspective on sustainability transitions: Responses to seven criticisms","Journal article","Environmental Innovation and Societal Transitions 1(1), 24–40","TBC","Canon","Verify","Global","MLP; agency and structure in transitions","Conceptual","Defends and refines MLP; acknowledges the need to theorise agency better.","RQ1","Not posthumanist.","Ch3 §3.9.","Medium"),
 ("L27","Davis, F.D.","1989","Perceived usefulness, perceived ease of use, and user acceptance of information technology","Journal article","MIS Quarterly 13(3), 319–340","TBC","Canon","Verify","Information systems","Technology Acceptance Model","Quantitative","Acceptance is predicted by perceived usefulness and ease of use.","RQ3 (as contrast)","Atomistic; artefact taken as given; no material or institutional context.","Ch3 §3.9 — used to explain why TAM is NOT adopted.","High"),
 ("L28","Venkatesh, V. et al.","2003","User acceptance of information technology: Toward a unified view","Journal article","MIS Quarterly 27(3), 425–478","TBC","Canon","Verify","Information systems","UTAUT","Quantitative","Unifies acceptance predictors (performance expectancy, effort expectancy, social influence, facilitating conditions).","RQ3 (as contrast)","Same ontological limits as TAM.","Ch3 §3.9.","Medium"),
 ("L29","Arnstein, S.R.","1969","A ladder of citizen participation","Journal article","Journal of the American Institute of Planners 35(4), 216–224","TBC","Canon","Verify","Urban planning / USA","Ladder of participation; tokenism","Conceptual","Distinguishes genuine participation from tokenism — essential for critiquing participatory claims.","RQ2, RQ4","Normative hierarchy; pre-dates relational ontologies.","Ch3 §3.7 (failure conditions of participation).","High"),
 ("L30","Checkland, P.","1981","Systems Thinking, Systems Practice","Book","Wiley, Chichester","TBC","Canon","Verify","Systems","Soft systems methodology; human activity systems","Methodological","Systems thinking must account for multiple worldviews and boundary judgements.","RQ1, RQ5","Human-centred; not materialist.","Ch3 §3.8, §3.9.","High"),
 ("L31","Meadows, D.H.","2008","Thinking in Systems: A Primer","Book","Chelsea Green, White River Junction","TBC","Canon","Verify","Systems","Stocks, flows, feedback, leverage points","Conceptual","System behaviour arises from structure and feedback, not from individual components.","RQ1, RQ4","General; no energy detail.","Ch3 §3.8; Ch4 feedback loops.","Medium"),
 ("L32","Midgley, G.","2000","Systemic Intervention: Philosophy, Methodology, and Practice","Book","Kluwer/Plenum, New York","TBC","Canon","Verify","Systems / critical systems","Boundary critique; systemic intervention","Methodological","Intervention requires reflection on boundaries — who and what is included.","RQ4, RQ5","Complex; needs selective use.","Ch3 §3.8 (boundary critique; links to posthumanism).","Medium"),
 ("L33","Leminen, S., Westerlund, M. & Nyström, A.-G.","2012","Living Labs as open-innovation networks","Journal article","Technology Innovation Management Review 2(9), 6–11","TBC","Canon","Cited in draft; verify","Finland / global","Living Lab typology (provider, utilizer, enabler, user-driven)","Conceptual / empirical","Living Labs are networks with different driving actors; typology shapes outcomes.","RQ1, RQ5","Global North; ICT bias.","Ch2; Ch3 §3.8; Ch4.","High"),
 ("L34","Dell'Era, C. & Landoni, P.","2014","Living Lab: A methodology between user-centred design and participatory design","Journal article","Creativity and Innovation Management 23(2), 137–154","TBC","Canon","Cited in draft; verify","Italy / Europe","Living Lab as design research methodology; aware users in real-life settings","Conceptual","Defines Living Lab as co-creation with aware users in real settings; distinguishes it from UCD and PD.","RQ1, RQ2","European; product-innovation focus.","Ch3 §3.8; Ch4.","High"),
 ("L35","Hossain, M., Leminen, S. & Westerlund, M.","2019","A systematic review of living lab literature","Journal article","Journal of Cleaner Production 213, 976–988","TBC","Canon","Cited in draft; verify","Global","Living Lab definitions; research themes","Systematic review","Synthesises Living Lab research; notes definitional fragmentation.","RQ1","Review; limited energy/justice content.","Ch2; Ch3 §3.8.","Medium"),
 ("L36","Schuurman, D. et al.","2015","TBC (Living Lab / user innovation)","Journal article","TBC","TBC","Secondary","Cited 11× in draft","Europe","Living Lab; user involvement","TBC","User involvement mechanisms in Living Labs.","RQ2","Secondary; title unverified.","Ch2; Ch3 §3.8.","Medium"),
 ("L37","Wüstenhagen, R., Wolsink, M. & Bürer, M.J.","2007","Social acceptance of renewable energy innovation: An introduction to the concept","Journal article","Energy Policy 35(5), 2683–2691","TBC","Canon","Cited in D-01 and draft; verify","Europe / global","Social acceptance: socio-political, community, market","Conceptual","Acceptance is multi-level; community acceptance differs from socio-political acceptance.","RQ3","Pre-dates relational and practice approaches.","Ch3 §3.7 (acceptance typology); Ch4.","High"),
 ("L38","Blessing, L.T.M. & Chakrabarti, A.","2009","DRM, a Design Research Methodology","Book","Springer, London","TBC","Canon","Cited in draft; verify","Design research","Design Research Methodology (RC, DS-I, PS, DS-II)","Methodological","Provides a staged framework for design research.","RQ5","Not sociological; needs pairing with social theory.","Ch5 (methodology).","High"),
 ("L39","Lincoln, Y.S. & Guba, E.G.","1985","Naturalistic Inquiry","Book","Sage, Beverly Hills","TBC","Canon","Cited in draft; verify","Qualitative research","Trustworthiness: credibility, transferability, dependability, confirmability","Methodological","Establishes quality criteria for qualitative inquiry.","RQ5","Constructivist; tension with posthumanist ontology.","Ch5 (trustworthiness; to be extended).","Medium"),
 ("L40","Easterby-Smith, M., Thorpe, R. & Jackson, P.","2012","Management Research","Book","Sage, London","TBC","Canon","Cited 6× in draft; verify edition","Management research","Ontology, epistemology, paradigm","Methodological","Frames management research philosophically; used for paradigm discussion.","RQ5","Generic; not posthumanist.","Ch5.","Medium"),
 ("L41","Miles, M.B. & Huberman, A.M.","1994","Qualitative Data Analysis: An Expanded Sourcebook","Book","Sage, Thousand Oaks","TBC","Canon","Cited in draft; verify","Qualitative methods","Conceptual framework as graphical/narrative model; coding","Methodological","Frameworks represent principal factors and presumed relationships.","RQ1","Not theoretical.","Ch4 §4.2; Ch5.","Medium"),
 ("L42","Booth, C. & Bennett, C.","2002","TBC (gender equality / 'equality stool')","Report / article","TBC","TBC","Secondary","Cited 3× in draft","UK / gender policy","Equality stool: equal treatment, positive action, gender mainstreaming","Conceptual","Three-legged approach to gender equality.","—","Likely to be removed under the end-user terminology instruction.","Ch2 (to be deleted or reframed).","Low"),
 ("L43","Postma, D.","2016","The ethics of becoming in a pedagogy for social justice: A posthumanist perspective","Journal article","South African Journal of Higher Education 30(3), 310–328","TBC","Canon","Cited in D-01; verify","South Africa","Posthumanist ethics; pedagogy; social justice","Conceptual","Posthumanist ethics of becoming applied to South African higher education.","RQ1","Education focus.","Ch3 §3.3.7 (Southern African posthumanist scholarship).","Medium"),
 ("L44","Motala, S.","2017","In/between science and art: Posthumanist ruminations on Geomatics education","Conference paper","Proceedings, 4th Biennial Conference of the South African Society for Engineering Education, 194–204","TBC","Canon","Cited in D-01; verify","South Africa","Posthumanism; engineering education; paradigm shift","Conceptual","Posthumanism as a deconstructive and generative stance in South African engineering education.","RQ1","Conference paper; education focus.","Ch3 §3.3 (localising the theory).","Medium"),
 ("L45","Heffron, R.J. & McCauley, D.","2017","The concept of energy justice across the disciplines","Journal article","Energy Policy 105, 658–667","TBC","Canon","Verify","Global","Energy justice across disciplines; restorative justice","Conceptual","Energy justice as a cross-disciplinary concept with restorative dimensions.","RQ4","Global North emphasis.","Ch3 §3.6.","Medium"),
 ("L46","Swilling, M. & Annecke, E.","2012","Just Transitions: Explorations of Sustainability in an Unfair World","Book","UCT Press, Cape Town","TBC","Canon","Verify","South Africa / global","Just transitions; sustainability; inequality","Conceptual","Just transition framed around inequality and sustainability in a South African context.","RQ1, RQ4","Pre-dates the post-2015 transition literature.","Ch3 §3.6 (South African just-transition framing).","High"),
 ("L47","Batinge, B., Musango, J.K., Ceschin, F., Ambole, A., Smit, S. & Petrulaityte, A.","2022","Modelling gendered innovation for the security of energy services in poor urban environments","Journal article","Systems Research and Behavioral Science 39(2), 287–304 (online 2021)","10.1002/sres.2793","Secondary","DOI verified via Crossref; full text not yet read","Africa / poor urban environments","Systems modelling; gendered innovation; energy services","Systems modelling","Models the system interactions through which gendered innovation affects security of energy services — a direct precedent for the systems dimension of this study.","RQ1, RQ4","Full text needed; does not use posthumanism or a Living Lab evaluation.","Ch2 (systems modelling); Ch3 §3.8; Ch4.","Critical"),
 ("L48","Musango, J.K.","2022","Assessing gender and energy in urban household energy transitions in South Africa: A quantitative storytelling from Groenheuwel informal settlement","Journal article","Energy Research & Social Science 88, 102525","10.1016/j.erss.2022.102525","Secondary","DOI verified via Crossref; full text not yet read","South Africa / Western Cape informal settlement","Gender; household energy transitions; quantitative storytelling","Mixed / quantitative storytelling","Empirical South African evidence on gender and household energy transitions in an informal settlement — essential contextual source.","RQ3","Single-site; gender-framed; full text needed.","Ch1 context; Ch2; Ch4; Ch5.","High"),
 ("L49","Oosthuizen, R., de Kock, I. & Musango, J.K.","2020","Exploring factors that influence the mainstreaming of gendered energy interventions in poor urban environments: A structured literature review","Journal article","South African Journal of Industrial Engineering 31(3)","10.7166/31-3-2421","Secondary","DOI verified via Crossref; full text not yet read","South Africa","Structured literature review; energy interventions; mainstreaming","Structured literature review","Identifies factors influencing mainstreaming of energy interventions in poor urban environments; useful for RQ4 mechanism design.","RQ4","Gender-framed; engineering management outlet (field-appropriate).","Ch2; Ch4 mechanisms.","Medium"),
]

for row in S:
    ws.append(list(row))

style_sheet(ws, [6, 26, 7, 42, 16, 30, 26, 10, 30, 24, 34, 22, 60, 12, 40, 40, 14], height=78)

# colour-code verification column (col 9) and read depth (col 8)
for r in range(2, ws.max_row + 1):
    v = str(ws.cell(row=r, column=9).value or "")
    fill = GREEN if v.startswith("Verified") else (RED if "corrupt" in v.lower() else AMBER)
    ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fill)
    d = str(ws.cell(row=r, column=8).value or "")
    ws.cell(row=r, column=8).fill = PatternFill("solid", fgColor=GREEN if d == "Full" else (AMBER if d in ("Canon", "Secondary", "Partial") else GREY))

# ---------------- Sheet 2: Gap map ----------------
ws = wb.create_sheet("Gap map")
ws.append(["Gap ID", "Gap", "Type", "Evidence that the gap exists", "RQ", "Chapter addressed", "Claimed contribution", "Defensibility", "Risk if overclaimed"])
gaps = [
 ("G1","No posthumanist account of energy technology innovation in informal settlements","Theoretical","Posthumanism is developed in education, geography and STS but absent from energy engineering management; D-01 is one of few exceptions and addresses engineering education, not settlements.","RQ1, RQ2, RQ3","Ch3 §§3.3–3.7","A posthumanist ontology of end-user participation","Strong","Low — the intersection is genuinely unoccupied"),
 ("G2","Energy justice is normatively rich but practically thin","Theoretical","Justice scholarship focuses on policy, siting and large infrastructure; rarely specifies design mechanisms.","RQ4","Ch3 §3.6; Ch4","Justice operationalised as design criteria","Moderate","Medium — must show the translation, not assert it"),
 ("G3","'Acceptance' is under-theorised and conflated with adoption/uptake","Theoretical","The draft uses five terms interchangeably; TAM, diffusion, social acceptance and practice theories rest on incompatible ontologies.","RQ3, RQ4","Ch3 §§3.4, 3.7, 3.9","A relational account of acceptance","Strong","Low if TAM is treated as contrast"),
 ("G4","Systems thinking and Living Lab theory are combined rhetorically, not theoretically","Theoretical","'Systems Living Lab' is used without deriving what 'systemic' implies for boundaries, stakeholders, feedback and evaluation.","RQ1, RQ5","Ch3 §3.8","Theoretical justification for a systemic Living Lab","Moderate","Medium"),
 ("G5","African energy Living Labs are thinly evidenced","Empirical","Evidence base is dominated by European and ICT/health cases; the draft itself calls the African literature 'lean'.","RQ5","Ch5–6","An evaluated South African case","Strong","Low"),
 ("G6","Urban informal settlements are under-represented relative to rural energy poverty","Empirical","Draft acknowledges the rural bias then repeatedly drifts into rural framing.","RQ2, RQ3","Ch2; Ch5","Context-specific analysis","Strong","Low"),
 ("G7","End-user knowledge is invoked but not characterised","Empirical / theoretical","The draft asserts end-users hold knowledge but never specifies type, elicitation or translation into design.","RQ2","Ch3 §3.7; Ch5","A typology of end-user knowledge","Moderate","Medium"),
 ("G8","Participation is assumed to produce acceptance and justice","Methodological","Literature on tokenism, non-use and capture contradicts the assumption; posthumanism highlights material constraints.","RQ4","Ch3 §3.7; Ch4","Conditions and failure modes of participation","Strong","Low — the critique strengthens the study"),
 ("G9","Living Lab evaluation lacks justice criteria","Methodological","Evaluation literature measures innovation outputs, not distributive/procedural/recognition outcomes.","RQ5","Ch3 §3.6; Ch6","Justice-oriented evaluation criteria","Strong","Low"),
 ("G10","The draft's claimed 'systematic review' is not systematic","Internal","No protocol, databases, strings, dates, criteria, counts or synthesis method reported.","—","Ch2","Methodological transparency","Procedural","High if left unfixed"),
]
for g in gaps: ws.append(list(g))
style_sheet(ws, [8, 46, 16, 56, 14, 20, 34, 14, 34], height=70)
for r in range(2, ws.max_row + 1):
    v = str(ws.cell(row=r, column=8).value or "")
    ws.cell(row=r, column=8).fill = PatternFill("solid", fgColor=GREEN if v=="Strong" else (AMBER if v=="Moderate" else RED))

# ---------------- Sheet 3: RQ-theory map ----------------
ws = wb.create_sheet("RQ-theory map")
ws.append(["RQ", "Question (abridged)", "Primary concepts", "Source IDs", "Analytical contribution", "Method implied", "What is lost without the theory"])
rq = [
 ("RQ1","How do Systems Living Lab concepts frame inequality, energy security, energy poverty, innovation and informal settlements?","Relational ontology; assemblage; energy justice; systemic boundaries","L03,L07,L22,L23,L30,L32,L46","Frames five apparent 'variables' as one configuration; makes the relations, not the elements, the object of analysis","Conceptual analysis + systematic review","The question becomes a definitional exercise"),
 ("RQ2","Which factors determine end-user involvement in decision-making, co-design, co-creation, co-research, prototyping and testing?","Distributed agency; intra-action; situated knowledge; participation and its failure modes","L05,L07,L11,L29,L33,L34","Explains participation as an effect of assembled relations (materials, institutions, expectations), not merely willingness","Interviews, workshops, stakeholder mapping, thematic analysis","Determinants collapse into a list of attitudinal factors"),
 ("RQ3","Which cultural elements, norms, practices and technological mechanisms influence acceptance?","Materiality; energy cultures; social acceptance; practice; embodiment","L08,L09,L20,L37,L27,L28","Explains acceptance as emerging from culture, practice, material conditions and technology together","Ethnography, observation, interviews, survey","Culture becomes a residual 'barrier'"),
 ("RQ4","Which mechanisms improve acceptance and end-user contribution to energy security services?","Leverage points; boundary critique; energy justice criteria; co-design","L22,L24,L30,L31,L32,L02","Identifies mechanisms that are just as well as effective; uses barrier evidence to design realistically","Co-design workshops, system dynamics, validation","Mechanisms become technical fixes"),
 ("RQ5","Can the proposed Systems Living Lab engage end-users and stakeholders in real-world settings to address energy poverty and enhance energy security and access?","Living Lab as apparatus; agential cut; response-ability; evaluation criteria","L06,L07,L12,L33,L34,L35,L38","Provides criteria for judging whether the Lab reconfigured the assemblage justly, and reflexivity about its own effects","Pilot/case study; mixed-method evaluation","Evaluation reduces to user satisfaction scores"),
]
for x in rq: ws.append(list(x))
style_sheet(ws, [7, 52, 42, 22, 54, 32, 38], height=90)

# ---------------- Sheet 4: Verification queue ----------------
ws = wb.create_sheet("Verification queue")
ws.append(["Order", "Task", "Why it matters", "Method", "Status", "Owner", "Target phase"])
vq = [
 (1,"Obtain and read the 8 most-cited sources in the draft (L13, L14, L15, L17, L16, L18, L19, L33–L34)","They carry the existing argument; without them nothing can be attributed accurately","Institutional library / UJ databases / open access / author contact","Not started","Researcher","Phase 6"),
 (2,"Resolve titles, journals, volumes, pages and DOIs for every L13–L19, L36, L42 entry marked 'Secondary' or 'TBC'","Prevents fabricated or garbled references","Crossref / publisher sites / library catalogue","Not started","Researcher + assistant","Phase 6"),
 (3,"Verify the seven Tier-1 theory works (L03–L09) including editions and page numbers for quotations","Chapter 3 depends on precise attribution","Publisher pages / library","Not started","Researcher","Phase 2–3"),
 (4,"Verify energy justice canon (L22–L24, L45, L46)","Chapter 3 §3.6","Publisher / Crossref","Not started","Researcher","Phase 3"),
 (5,"Verify TAM/UTAUT originals (L27, L28) and social acceptance (L37)","Chapter 3 §3.9 argument depends on representing them fairly","Library","Not started","Researcher","Phase 3"),
 (6,"Build the full reference list: every in-text citation in 68,000 words matched to a verified source","No reference list currently exists","Systematic extraction + matching","Not started","Assistant","Phase 6"),
 (7,"Remove or replace citations that cannot be matched to a real source","Citation integrity","Editorial pass","Not started","Assistant","Phase 6"),
 (8,"Create EndNote library, import verified records, export research_references.ris","Client requirement; institutional requirement (D-04 §12)","EndNote desktop / RIS","Not started","Researcher","Phase 6"),
 (9,"Download legally accessible PDFs to /research_work/references/","Client requirement","Open access, repositories, author manuscripts","Not started","Researcher","Phase 6"),
 (10,"Record DOI unresolvable items in missing_references.md rather than guessing","Prevents fabrication","Log","Not started","Assistant","Phase 6"),
]
for x in vq: ws.append(list(x))
style_sheet(ws, [7, 62, 42, 40, 14, 20, 14], height=60)

wb.save(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\literature_matrix.xlsx")
print("saved")
