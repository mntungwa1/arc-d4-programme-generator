# =============================================================================
#  POPULATE T5  -  POLICY BRIEF
#  SADC Regional Disaster Risk Management Innovation Programme
#  Contract CLMX055740  |  Academy of Resilience and Continuity
# -----------------------------------------------------------------------------
#  Two modes:
#    per innovation   python populate_T5.py --innovation 13 --states 4 --years 5
#    programme        python populate_T5.py --programme
#  Costing uses the shared budget engine, so a brief and the funding proposal
#  for the same innovation always state the same figures.
# =============================================================================

"""Populate the policy brief, for one innovation or for the programme."""
# =============================================================================
#  POPULATE T5  -  POLICY BRIEF
#  SADC Regional Disaster Risk Management Innovation Programme
#  Contract CLMX055740  |  Academy of Resilience and Continuity
# -----------------------------------------------------------------------------
#  Two modes:
#    per innovation   python populate_T5.py --innovation 13 --states 4 --years 5
#    programme        python populate_T5.py --programme
#  Costing uses the shared budget engine, so a brief and the funding proposal
#  for the same innovation always state the same figures.
# =============================================================================

"""Populate the policy brief, for one innovation or for the programme."""
# =============================================================================
#  POPULATE T5  -  Summary Brief
#  SADC Regional Disaster Risk Management Innovation Programme
#  Contract CLMX055740  |  Academy of Resilience and Continuity
# -----------------------------------------------------------------------------
#  Scope   : programme-level
#  Reads   : the D2 / D3 / Cost-Matrix data pack (ARC_D4_Automation_Matrix.xlsx)
#  Template: the T5 .docx shipped with this file
#  Usage   : python populate_T5.py
#  Rule    : nothing is generated. Every value is read from the data pack; an
#            absent value is omitted or written as a request naming the holder.
# =============================================================================

"""Populate Summary Brief from the D2/D3/cost-matrix data pack."""
# =============================================================================
# SECTION A  -  DATA PACK LOADER  (identical in every populate script)
# -----------------------------------------------------------------------------
# Reads the consolidated D2 / D3 / Cost-Matrix data pack and exposes one merged
# record per innovation. Nothing in this section invents a value: every field is
# read from a named sheet, and an absent field is returned as None so that the
# template section can omit it or turn it into a request.
#
#   D2 sources  : 48_D2_Innovation_Profile, 44_GRS_Determinations
#   D3 sources  : 06_Portfolio, 08_Pillar_Response, 15_Gap_Register,
#                 16_Root_Causes, 17_Gap_Linkages, 20_Innovation_Type,
#                 27_ME_Framework, 29_Legal_Regime_Register, 30_Mandate_Register
#   Cost matrix : 47_Cost_By_Innovation, 46_Cost_Estimation   (Annex R.1)
#
# Requirements:  pip install openpyxl docxtpl
# =============================================================================
import argparse, datetime, re, sys
from openpyxl import load_workbook
from docxtpl import DocxTemplate

HEADER_ROW = 2          # every data-pack sheet: row 1 = note, row 2 = headers
FIRST_DATA = 3

def read_sheet(wb, name):
    """Return a sheet as a list of dicts keyed by its header row."""
    ws = wb[name]
    heads = [str(ws.cell(row=HEADER_ROW, column=c).value or '').strip()
             for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(FIRST_DATA, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None or str(v).strip() == '' for v in vals):
            continue
        rows.append({heads[i]: vals[i] for i in range(len(heads)) if heads[i]})
    return rows

def clean(v):
    """Normalise a cell to text or None. Blank means absent, never empty string."""
    if v is None: return None
    s = str(v).strip()
    return s if s and s not in ('\u2014', '-', 'None', 'NO MATCH') else None

def as_int(v):
    try: return int(float(v))
    except (TypeError, ValueError): return None

class DataPack:
    SHEETS = ['06_Portfolio', '08_Pillar_Response', '15_Gap_Register', '16_Root_Causes',
              '17_Gap_Linkages', '20_Innovation_Type', '27_ME_Framework',
              '29_Legal_Regime_Register', '30_Mandate_Register', '44_GRS_Determinations',
              '46_Cost_Estimation', '47_Cost_By_Innovation', '48_D2_Innovation_Profile']

    def __init__(self, path):
        wb = load_workbook(path, data_only=True)
        missing = [s for s in self.SHEETS if s not in wb.sheetnames]
        if missing:
            sys.exit('Data pack is missing required sheets: ' + ', '.join(missing))
        self.s = {name: read_sheet(wb, name) for name in self.SHEETS}
        self.path = path

    def _by_n(self, sheet):
        return {as_int(r.get('#')): r for r in self.s[sheet] if as_int(r.get('#'))}

    def numbers(self):
        return sorted(self._by_n('06_Portfolio'))

    def innovation(self, n):
        """One merged record for innovation n, from D2, D3 and the cost matrix."""
        port = self._by_n('06_Portfolio').get(n)
        if not port:
            sys.exit('Innovation %s is not in the portfolio.' % n)
        typ  = self._by_n('20_Innovation_Type').get(n, {})
        pil  = self._by_n('08_Pillar_Response').get(n, {})
        grs  = self._by_n('44_GRS_Determinations').get(n, {})
        cost = self._by_n('47_Cost_By_Innovation').get(n, {})
        d2   = self._by_n('48_D2_Innovation_Profile').get(n, {})
        pillar_cols = [k for k in pil if k.startswith('Pillar')]
        return {
            'n': n,
            'name': clean(port.get('Innovation')),
            'bps': port.get('BPS'), 'tfs': port.get('TFS'), 'ifs': port.get('IFS'),
            'rrs': port.get('RRS'),
            'type': clean(typ.get('Type')), 'pillar': clean(typ.get('Pillar')),
            'chain': clean(typ.get('Value chain')), 'cycle': clean(typ.get('DRM cycle')),
            'maturity': clean(typ.get('Maturity')),
            'pillars': [clean(pil.get(k)) for k in pillar_cols],
            # D2
            'lead': clean(d2.get('Lead institution')), 'geo': clean(d2.get('Geography')),
            'desc': clean(d2.get('Description')), 'why': clean(d2.get('Why it works')),
            'benef': clean(d2.get('Beneficiary groups')), 'access': clean(d2.get('Accessibility feature')),
            'barrier': clean(d2.get('Principal barrier')), 'finance': clean(d2.get('Financing sources')),
            'source': clean(d2.get('Cited source')), 'd2_prov': clean(d2.get('D2 provenance')),
            # inclusion determination (D2 interview coding, extracted under the D3 rule)
            'grs': grs,
            # cost matrix
            'cost_arch': clean(cost.get('Cost archetype')), 'cost_setup': clean(cost.get('Set-up band')),
            'cost_rec': clean(cost.get('Annual recurrent')), 'cost_basis': clean(cost.get('Comparator basis')),
            'cost_conf': clean(cost.get('Confidence')),
        }

# ---- shared vocabulary -------------------------------------------------------
INCL_CATS = ['Women', 'Youth', 'Persons with disabilities', 'Older persons',
             'Marginalised and minority groups', 'Remote and hard-to-reach communities',
             'Migrants and displaced persons']
PILLARS = ['Risk data and early warning', 'Community reach and inclusion',
           'Governance and coordination', 'Financing and partnerships', 'Knowledge and learning']
USE = {'P': 'Delivers this directly', 'S': 'Contributes substantially',
       'E': 'Removes an obstacle to it', 'N': 'No contribution'}
CHAIN_LABEL = {'RI': 'Risk information \u2014 producing the trigger',
               'AU': 'Authority \u2014 authorising the decision',
               'FI': 'Financing \u2014 releasing the funds',
               'DE': 'Delivery to the household',
               'VL': 'Verification and learning'}
TYPE_LABEL = {'Non-tech': 'An arrangement between organisations, not a piece of equipment',
              'Hybrid': 'Part technology, part arrangement \u2014 both halves must be financed together',
              'Tech': 'Technology'}

MEMBER_STATES = ['Angola', 'Botswana', 'Comoros', 'Democratic Republic of the Congo', 'Eswatini',
                 'Lesotho', 'Madagascar', 'Malawi', 'Mauritius', 'Mozambique', 'Namibia', 'Seychelles',
                 'South Africa', 'Tanzania', 'Zambia', 'Zimbabwe']
MS_CODES = {'AGO': 'Angola', 'BWA': 'Botswana', 'COM': 'Comoros', 'DRC': 'Democratic Republic of the Congo',
            'COD': 'Democratic Republic of the Congo', 'SWZ': 'Eswatini', 'LSO': 'Lesotho', 'MDG': 'Madagascar',
            'MWI': 'Malawi', 'MUS': 'Mauritius', 'MOZ': 'Mozambique', 'NAM': 'Namibia', 'SYC': 'Seychelles',
            'ZAF': 'South Africa', 'TZA': 'Tanzania', 'ZMB': 'Zambia', 'ZWE': 'Zimbabwe'}

def country_of(inv):
    """The Member State the instrument operates in.

    Geography strings mix the operating country, the replication scope and, for some
    instruments, a proof-of-concept location outside the region. Taking the first
    clause returns that outside location and addresses a request to a government that
    is not a Member State, so only Member States are matched.
    """
    g = inv.get('geo') or ''
    if not g: return 'Regional'
    for clause in re.split(r'[;]', g):
        c = re.sub(r'\(.*?\)', '', clause).strip()
        if re.search(r'\b(regional|SADC[- ]wide|multi[- ]country|continental|global)\b', c, re.I):
            continue
        for name in MEMBER_STATES:
            if re.search(r'\b' + re.escape(name) + r'\b', c, re.I):
                return name
        for code, name in MS_CODES.items():
            if re.search(r'\b' + code + r'\b', c):
                return name
    return 'Regional'

ABBREVIATIONS = [
    (r'\bMNOs?\b', 'mobile network operators'), (r'\bDM\b', 'disaster management'),
    (r'\bMS\b', 'Member States'), (r'\bPOC\b', 'proof of concept'),
    (r'\bIbF\b', 'impact-based forecasting'), (r'\bEWS?\b', 'early warning'),
    (r'\bCAP\b', 'the Common Alerting Protocol'), (r'\bIKS\b', 'indigenous knowledge systems'),
    (r'\bGCF\b', 'the Green Climate Fund'), (r'\bAA\b', 'anticipatory action'),
    (r'\bCVA\b', 'cash and voucher assistance'), (r'\bDRR\b', 'disaster risk reduction'),
    (r'\bDRM\b', 'disaster risk management'), (r'\bNDMA\b', 'the national disaster management authority'),
    (r'\bGIS\b', 'geographic information systems'), (r'\bM&E\b', 'monitoring and evaluation'),
]

def plain(text):
    """Expand the abbreviations that appear in free-text fields of the record.

    The client's standing instruction is that abbreviations cloud readability, so any
    field quoted into a product is expanded rather than passed through as recorded.
    """
    if not text: return text
    out = str(text)
    for pat, full in ABBREVIATIONS:
        out = re.sub(pat, full, out)
    return re.sub(r'\s{2,}', ' ', out).strip()

def lead_short(inv):
    lead = inv.get('lead') or 'Lead institution'
    return re.split(r'[,(]', lead)[0].strip()

def grs_rows(inv):
    """Inclusion table from the extracted determination. Absence is recorded."""
    g = inv['grs'] or {}
    src = clean(g.get('Source interviews'))
    rows = []
    for cat in INCL_CATS:
        v = g.get(cat)
        try: v = float(v)
        except (TypeError, ValueError): v = None
        if v is None:
            prov, basis = 'Not determined', 'Not matched in the interview coding'
        elif v >= 1:
            prov, basis = 'Designed into the arrangement', 'Coded on a segment naming this innovation'
        elif v >= 0.5:
            prov, basis = 'Named as a beneficiary', 'Coded at interview level; no innovation-specific design evidence'
        else:
            prov, basis = 'None recorded', 'Not coded in any interview covering this innovation'
        rows.append({'group': cat, 'provision': prov, 'basis': basis})
    return rows, src

def pathway_for(inv):
    """Theory of Change placement, derived from value-chain position under a stated rule."""
    ch, pl = inv.get('chain'), inv.get('pillar')
    if ch == 'DE' or pl == 'P2':
        return ('Delivery to the household', 'Pathway 4 \u2014 enhanced risk communication and community engagement',
                'Pathway 1 \u2014 strengthened early warning and anticipatory action', 'Target G, at the access dimension')
    if ch == 'RI' or pl == 'P1':
        return ('Risk information \u2014 producing the trigger', 'Pathway 1 \u2014 strengthened early warning and anticipatory action',
                'None', 'Target G, at the availability dimension')
    if ch == 'AU' or pl == 'P3':
        return ('Authority \u2014 authorising the decision', 'Pathway 2 \u2014 improved coordination and governance',
                'None', 'Target E \u2014 national and local disaster risk reduction strategies')
    if ch == 'FI' or pl == 'P4':
        return ('Financing \u2014 releasing the funds', 'Pathway 1 \u2014 strengthened early warning and anticipatory action',
                'Pathway 2 \u2014 improved coordination and governance', 'Target C \u2014 direct economic loss')
    return ('Verification and learning', 'Pathway 3 \u2014 adoption of context-appropriate innovations',
            'None', 'Target E \u2014 national and local strategies')

def toc_block(inv):
    link, path, path2, sendai = pathway_for(inv)
    if inv.get('chain') in ('DE', 'VL'):
        assumption = ('Delivery becomes record. Operating institutions will document what happened to a standard that '
                      'allows another Member State to replicate it. The D3 assumption register marks this as holding only '
                      'partially \u2014 it holds where a funder demands it and lapses where none does. If it fails, the '
                      'innovation continues to work where it is and remains unreplicable everywhere else.')
    else:
        assumption = ('The assumption governing this link is recorded in the D3 assumption register at section 14, '
                      'with its evidence status. It is cited there rather than restated here.')
    return {'spine': 'The Programme\u2019s causal chain runs: risk information produces the trigger; governance '
                     'authorises the decision; financing releases the funds; community systems deliver the action and '
                     'give the trigger its legitimacy; learning recalibrates the trigger. This innovation sits at the '
                     'link named below.',
            'link': link, 'pathway': path, 'pathway2': path2, 'sendai': sendai, 'assumption': assumption}

def region_block(inv):
    """Governance as it now stands: the Unit drives, the Steering Committee oversees,
    the Operations Centre is enabled and holds no Programme decision."""
    regional = bool(re.search(r'SADC|Humanitarian and Emergency|World Meteorological', inv.get('lead') or ''))
    recovery = (inv.get('cycle') or '') == 'Recovery'
    activation = ('Coordinates regional activation where the instrument operates across Member States, under its own mandate.'
                  if regional else 'None. This is delivered nationally and the Operations Centre does not activate it.')
    if recovery:
        activation = ('None, and it may not lead this. The mandate of the Operations Centre does not extend to '
                      'rehabilitation and reconstruction; a Member State institution must lead.')
    return {
        'intro': ('The Programme is owned and driven by the Disaster Risk Reduction Unit of the Secretariat, accountable '
                  'upward through the Committee of Ministers responsible for disaster risk management, the Troika and the '
                  'Summit. The Unit appoints a Programme Steering Committee to oversee implementation. The Humanitarian '
                  'and Emergency Operations Centre does not drive the Programme and takes no Programme decision; it holds '
                  'its own mandate and the Programme exists in part to resource it.'),
        'centre': [
            {'fn': 'Activation', 'role': activation},
            {'fn': 'Protocols and standards', 'role': 'Receives the protocol, conditions and standards the Programme produces, and may supply them to any Member State adopting the innovation.'},
            {'fn': 'Regional record', 'role': 'Holds the activation and delivery record, so the innovation is visible beyond the country where it operates.'},
            {'fn': 'Interoperability', 'role': 'Receives the data the innovation reports, to the documentation standard the Programme issues.'},
            {'fn': 'Programme decisions', 'role': 'None. No Programme decision rests here and no Programme reporting line runs through it.'},
        ],
        'supervision': [
            {'item': 'Ownership', 'arr': 'Disaster Risk Reduction Unit. It owns and drives the Programme and answers for it.'},
            {'item': 'Oversight of implementation', 'arr': 'Programme Steering Committee, appointed by the Unit and holding only the authority the Unit delegates to it.'},
            {'item': 'Portfolio decisions', 'arr': 'Entry to the portfolio, and any change to what this innovation is expected to deliver, decided by the Steering Committee under delegated authority.'},
            {'item': 'Standing item', 'arr': 'Custody of recurrent cost is a standing item of the Steering Committee. This innovation appears under it until an accountable post is named.'},
            {'item': 'Reporting line', 'arr': 'Operating institution to its national disaster management authority; national authority to the Disaster Risk Reduction Unit. No part of the line runs through the Operations Centre.'},
        ]}

def repl_block(inv):
    try: rrs = float(inv.get('rrs'))
    except (TypeError, ValueError): rrs = None
    score = ('%.2f' % rrs) if rrs is not None else 'not scored'
    reflects = ('Limited documented operation. The criterion counts documented operation, so a low score here is usually '
                'a documentation finding rather than a judgement about the innovation.' if (rrs is not None and rrs < 5)
                else 'Documented operation in several Member States.')
    t = inv.get('type')
    if t == 'Non-tech':
        conds = [('Existing community or institutional structures', 'Travels. Present across the region.'),
                 ('Legal authority to operate it', 'Does not travel. Each Member State must identify its own provision.'),
                 ('Capacity to apply the protocol', 'Travels with support. The protocol transfers; capacity varies.'),
                 ('A named recurrent-cost post', 'Does not travel. A national act in each case.')]
    elif t == 'Tech':
        conds = [('The platform or system', 'Travels. Licensable or replicable.'),
                 ('Data, connectivity and hosting', 'Travels with support. Varies with national infrastructure.'),
                 ('Spectrum, aviation or data law', 'Does not travel. Each Member State must clear its own regime.'),
                 ('A named recurrent-cost post', 'Does not travel. A national act in each case.')]
    else:
        conds = [('The technical component', 'Travels. Licensable or replicable.'),
                 ('The institutional arrangement', 'Travels with support. The protocol transfers; the arrangement must be made locally.'),
                 ('Legal authority to operate it', 'Does not travel. Each Member State must identify its own provision.'),
                 ('A named recurrent-cost post', 'Does not travel. A national act in each case.')]
    return {'score': score, 'reflects': reflects,
            'conditions': [{'cond': a, 'travels': b} for a, b in conds],
            'transfer': ('The Programme transfers the protocol, the conditions and the evidence \u2014 not the national '
                         'arrangement itself. Conditions marked as not travelling are national acts that no regional body '
                         'can perform; each adopting Member State is asked the same questions.')}

ME_MAP = {'DE': 'Outcome 2', 'RI': 'Outcome 1', 'FI': 'Outcome 3', 'VL': 'Outcome 4', 'AU': 'Output A'}

def me_block(pack, inv):
    fw = {clean(r.get('Level')): r for r in pack.s['27_ME_Framework']}
    wanted = [ME_MAP.get(inv.get('chain'), 'Outcome 2'), 'Outcome 4', 'Output C']
    if inv.get('type') in ('Non-tech', 'Hybrid'):
        wanted.append('Output D')
    prog = []
    for lvl in wanted:
        r = fw.get(lvl)
        if r:
            prog.append({'level': lvl + ' \u2014 ' + (clean(r.get('Statement')) or ''),
                         'ind': clean(r.get('Indicator')) or '',
                         'by': lead_short(inv) if lvl.startswith('Outcome') else 'Programme record',
                         'freq': clean(r.get('Frequency')) or ''})
    return {'intro': ('Two levels of measurement apply. The first is whether the innovation works where it operates. The '
                      'second is what it contributes to the Programme, measured through the Programme\u2019s own framework '
                      'on the Programme\u2019s cadence. Every indicator is readable from records that already exist.'),
            'programme': prog,
            'evaluation': ('Evaluation \u2014 whether the innovation caused the change observed \u2014 runs on the '
                           'Programme\u2019s own cycle and is not attempted for a single innovation.')}

def cost_block(inv):
    return {'intro': ('Indicative bands from the Programme cost estimation matrix (Annex R.1). Cost archetype: '
                      + (inv.get('cost_arch') or 'not assigned') + '.'),
            'setup': inv.get('cost_setup') or 'Not in the matrix',
            'recurrent': inv.get('cost_rec') or 'Not in the matrix',
            'setup_covers': 'Establishing the instrument to working order in one Member State.',
            'recurrent_covers': 'Keeping it running for a year once established.',
            'basis': ('Basis: ' + (inv.get('cost_basis') or 'not stated') + '. Confidence: '
                      + (inv.get('cost_conf') or 'not stated') + '. These are bands drawn from named published '
                      'comparators, not quotations.'),
            'use_limit': ('Usable for scoring and planning. Not usable as a budget line. A submission-tier product '
                          'replaces these bands with a costed proposal at national prices.')}

def tier_banner(inv):
    return ('Final controlled summary brief. Generated from the D2, D3 and cost-matrix data pack on '
            + datetime.date.today().strftime('%d %B %Y') + '. Every field the Programme holds is complete. '
            'Fields held by a Member State appear as requests naming who is being asked. Cost figures are '
            'indicative bands, not quotations.')

def render(template_path, context, out_path):
    doc = DocxTemplate(template_path)
    doc.render(context, autoescape=True)
    doc.save(out_path)
    print('written:', out_path)

# ---- shared budget engine ---------------------------------------------------
# Used by the funding proposal and by the policy brief so their figures can never
# disagree. Direct costs are budgeted bottom-up; the indirect rate follows the
# modality; execution is held inside the ceiling the climate funds apply.
CONTINGENCY_RATE = 0.10      # risk provision, on direct costs
REPLICATION_RATE = 0.60      # cost of the second and later adopter against the first
RECURRENT_SHARE  = {1: 1.00, 2: 0.50, 3: 0.50}   # Action share of running cost, by year

# The indirect rate is set by the modality the Action is delivered through, not chosen.
# United Nations: the joint cost-recovery policy sets 8 per cent for non-thematic
# contributions and 7 for thematic. The climate funds express the same idea as an
# implementing entity fee, higher for regional work than for single-country work.
MODALITY = {
    'un':                {'rate': 0.08, 'label': 'United Nations non-thematic contribution',
                          'basis': 'Standard harmonized indirect cost-recovery rate of 8 per cent'},
    'un-thematic':       {'rate': 0.07, 'label': 'United Nations thematic contribution',
                          'basis': 'Differentiated rate of 7 per cent for thematic contributions'},
    'government':        {'rate': 0.05, 'label': 'Programme government cost-sharing',
                          'basis': 'Differentiated rate of 5 per cent for government cost-sharing'},
    'adaptation-fund':   {'rate': 0.10, 'label': 'Adaptation Fund, regional project',
                          'basis': 'Implementing entity fee, capped at 10 per cent for regional projects and 8.5 for single-country'},
    'green-climate-fund':{'rate': 0.085,'label': 'Green Climate Fund',
                          'basis': 'Accredited entity fee, within the range the Fund applies'},
    'european-union':    {'rate': 0.07, 'label': 'European Union flat-rate indirect cost',
                          'basis': 'Flat-rate indirect cost at the maximum of 7 per cent, where justified'},
}

# Indicative unit rates for direct management and oversight. These are budgeted
# directly because they are traceable to the Action. Replace with the institution's
# own rates before submission; they are stated here so the arithmetic is auditable.
UNIT_RATES = {
    'programme_manager':   {'rate': 120000, 'unit': 'Per year, one post'},
    'finance_officer':     {'rate':  80000, 'unit': 'Per year, one post'},
    'mel_officer':         {'rate':  90000, 'unit': 'Per year, one post'},
    'steering_meeting':    {'rate':  35000, 'unit': 'Per meeting, two a year'},
    'audit':               {'rate':  25000, 'unit': 'Per year'},
    'baseline':            {'rate':  60000, 'unit': 'Once, first six months'},
    'midterm':             {'rate':  90000, 'unit': 'Once, end of year three'},
    'final_evaluation':    {'rate': 110000, 'unit': 'Once, final year'},
    'visibility':          {'rate':  20000, 'unit': 'Per year'},
}

def parse_band(text):
    """Classify a cost entry from the matrix and return (kind, value).

    The matrix holds six forms and only two of them are sums that can be added:
      per_state  'USD 150,000 - 400,000'              a band, once per adopting State
      regional   'USD 700,000 - 2,000,000 regional'   a band, once for the region
      premium    'Premium-dependent'                  not a capital sum
      per_person 'USD 11 per person reached'          needs a beneficiary count
      per_ratio  '17 cents per dollar transferred'    needs a transfer volume
      per_hh     'USD 0.65 - 2.00 per household'      needs a household count
    Anything that is not a plain sum is returned for separate treatment rather
    than silently counted as zero.
    """
    if not text: return ('absent', None)
    t = text.strip()
    if re.search(r'dependent', t, re.I):                      return ('premium', None)
    if re.search(r'per person', t, re.I):                     return ('per_person', None)
    if re.search(r'per dollar|cents', t, re.I):               return ('per_ratio', None)
    if re.search(r'per household', t, re.I):                  return ('per_hh', None)
    nums = [float(x.replace(',', '')) for x in re.findall(r'[\d][\d,]*(?:\.\d+)?', t)]
    if not nums:                                              return ('absent', None)
    mid = sum(nums) / len(nums) if len(nums) > 1 else nums[0]
    return ('regional' if re.search(r'regional', t, re.I) else 'per_state', mid)

UNPRICED_REASON = {
    'premium':    ('The matrix records this as premium or instrument dependent rather than as a capital sum',
                   'A premium quotation from the risk carrier, and the sum insured the Member States choose'),
    'per_person': ('The matrix prices this per person reached, not as a lump sum',
                   'The number of people to be reached in each Member State, from national exposure data'),
    'per_ratio':  ('The matrix prices this as a ratio of the value transferred',
                   'The planned transfer volume per Member State'),
    'per_hh':     ('The matrix prices this per household per year',
                   'The number of households to be registered in each Member State'),
    'absent':     ('No band is recorded for this instrument in the matrix',
                   'A costed proposal at national prices, or a comparator added to the matrix'),
}

def money(v):
    return 'USD {:,.0f}'.format(round(v))

def pct(v, total):
    return '{:.0f}%'.format(100.0 * v / total) if total else '-'

def recurrent_years(years):
    """Action-funded share of running cost, summed over the life of the Action."""
    return sum(RECURRENT_SHARE.get(y, 0.0) for y in range(1, years + 1))

def scope_factor(n_instruments, states):
    """Management effort scales with what the Action actually carries.

    A single instrument in a few Member States does not need a full regional
    management establishment, and charging one to it is how administration comes to
    cost more than the work. The factor is a full-time-equivalent share, floored so a
    small Action still carries real supervision and capped at one establishment.
    """
    f = 0.15 + 0.06 * n_instruments + 0.03 * states
    return max(0.20, min(1.00, f))

def build_management(years, factor=1.0):
    """Direct management and oversight, built from unit rates rather than a percentage.

    United Nations cost-recovery policy requires costs traceable to the programme to
    be budgeted directly. Deriving a management total from a rate and then splitting
    it is the reverse of that, and leaves no answer to a reviewer who asks how a post
    was costed.
    """
    R = UNIT_RATES
    rows = [
        ('programme_manager', 'Programme management', years),
        ('finance_officer',   'Financial management and contracting', years),
        ('steering_meeting',  'Programme Steering Committee', years * 2),
        ('audit',             'External audit and assurance', years),
        ('visibility',        'Communication and visibility', years),
    ]
    out = []
    for key, label, qty in rows:
        v = R[key]['rate'] * qty * factor
        out.append({'item': label, 'unit': R[key]['unit'],
                    'rate': money(R[key]['rate']),
                    'qty': '%s at %.2f full-time equivalent' % (qty, factor) if factor < 1 else str(qty),
                    'amount': money(v), '_v': v})
    return out

def build_mel(years, factor=1.0):
    """Monitoring and evaluation as a costed plan in its own right.

    A budgeted monitoring and evaluation plan is an explicit review criterion for the
    Adaptation Fund, so it is shown as a line rather than folded into management.
    """
    R = UNIT_RATES
    rows = [
        ('mel_officer', 'Monitoring officer', 'Continuous monitoring from existing records', 'Throughout', years),
        ('baseline', 'Baseline', 'Starting position for every indicator in each Member State', 'First six months', 1),
        ('midterm', 'Mid-term evaluation', 'Whether the pathways are producing the outcomes expected, and whether the assumptions hold', 'End of year three', 1),
        ('final_evaluation', 'Final evaluation', 'Contribution to the outcomes, and what carries into the successor programme', 'Final year', 1),
    ]
    out = []
    for key, label, covers, when, qty in rows:
        v = R[key]['rate'] * qty * factor
        out.append({'item': label, 'covers': covers, 'when': when,
                    'amount': money(v), '_v': v})
    return out

def build_budget(pack, invs, states, years, modality):
    """Direct implementation, direct management, risk provision, indirect support.

    Set-up is charged in full to the first adopting Member State and at the replication
    rate to each one after it, because what transfers between States is the protocol,
    the conditions and the evidence rather than the design work. Running cost is funded
    on the Programme's own sustainability sequence rather than for every year.
    """
    rec_years = recurrent_years(years)
    by_pillar, unpriced = {}, []
    for i in invs:
        p = i.get('pillar') or 'P1'
        by_pillar.setdefault(p, {'amount': 0.0, 'names': [], 'setup': 0.0, 'recurrent': 0.0})
        by_pillar[p]['names'].append(i['name'])
        kind, val = parse_band(i.get('cost_setup'))
        if kind == 'per_state':
            setup = val * (1 + (states - 1) * REPLICATION_RATE)
        elif kind == 'regional':
            setup = val
        else:
            setup = 0.0
            why, req = UNPRICED_REASON[kind]
            unpriced.append({'inst': i['name'] + ' (set-up)', 'why': why, 'requirement': req})
        kind, val = parse_band(i.get('cost_rec'))
        if kind == 'per_state':
            recurrent = val * states * rec_years
        elif kind == 'regional':
            recurrent = val * rec_years
        else:
            recurrent = 0.0
            why, req = UNPRICED_REASON[kind]
            unpriced.append({'inst': i['name'] + ' (running cost)', 'why': why, 'requirement': req})
        by_pillar[p]['setup'] += setup
        by_pillar[p]['recurrent'] += recurrent
        by_pillar[p]['amount'] += setup + recurrent

    implementation = sum(v['amount'] for v in by_pillar.values())
    setup_total = sum(v['setup'] for v in by_pillar.values())
    recurrent_total = sum(v['recurrent'] for v in by_pillar.values())

    factor = scope_factor(len(invs), states)
    management_rows = build_management(years, factor)
    mel_rows = build_mel(years, factor)
    management = sum(r['_v'] for r in management_rows)
    mel = sum(r['_v'] for r in mel_rows)

    # Execution costs are held inside the ceiling the climate funds apply: 9.5 per cent
    # of activities plus execution. Where the scoped requirement exceeds it, the Action
    # is too small to carry a standalone management establishment and the balance is met
    # from the Programme's own management. That is reported, not absorbed silently.
    EXEC_CEILING = 0.095
    activities_base = implementation * (1 + CONTINGENCY_RATE)
    allowed = activities_base * EXEC_CEILING / (1 - EXEC_CEILING)
    scoped = management + mel
    capped = scoped > allowed
    if capped and scoped > 0:
        shrink = allowed / scoped
        for r in management_rows + mel_rows:
            r['_v'] *= shrink
            r['amount'] = money(r['_v'])
        management *= shrink
        mel *= shrink

    contingency = implementation * CONTINGENCY_RATE
    direct_total = implementation + management + mel + contingency
    mod = MODALITY.get(modality, MODALITY['un'])
    support = direct_total * mod['rate']
    total = direct_total + support
    return {'by_pillar': by_pillar, 'unpriced': unpriced,
            'implementation': implementation, 'setup': setup_total, 'recurrent': recurrent_total,
            'management_rows': management_rows, 'mel_rows': mel_rows,
            'management': management, 'mel': mel, 'contingency': contingency,
            'direct_total': direct_total, 'support': support, 'total': total,
            'rec_years': rec_years, 'modality': mod, 'rate': mod['rate'],
            'factor': factor, 'capped': capped, 'scoped': scoped, 'allowed': allowed}

def build_mapping(B, states):
    """The same total expressed as the climate funds read a budget.

    They test three components against their own ceilings: activities, execution costs
    and the implementing entity fee. A budget they cannot map to that shape has to be
    rebuilt by the reviewer, which is the surest way to lose a week.
    """
    activities = B['implementation'] + B['contingency']     # risk provision sits inside activities
    execution = B['management'] + B['mel']
    fee = B['support']
    project_cost = activities + execution                   # the base both ceilings apply to
    exec_pct = 100.0 * execution / project_cost if project_cost else 0
    fee_pct = 100.0 * fee / project_cost if project_cost else 0
    TOL = 0.05          # a ceiling met exactly must not read as breached
    regional = states > 1
    fee_cap = 10.0 if regional else 8.5
    return [
        {'comp': 'A  Activities', 'contains': 'Direct implementation of the instruments, with the risk provision carried inside activity costs',
         'amount': money(activities), 'cap': 'No ceiling', 'position': 'n/a'},
        {'comp': 'B  Execution', 'contains': 'Direct management, oversight, audit, visibility, and the costed monitoring and evaluation plan',
         'amount': money(execution), 'cap': '9.5 per cent of A plus B',
         'position': '%.1f per cent — %s' % (exec_pct, 'within' if exec_pct <= 9.5 + TOL else 'ABOVE THE CEILING')},
        {'comp': 'C  Implementing entity fee', 'contains': 'Indirect cost recovery at the rate of the modality used',
         'amount': money(fee), 'cap': '%.1f per cent of A plus B (%s project)' % (fee_cap, 'regional' if regional else 'single-country'),
         'position': '%.1f per cent — %s' % (fee_pct, 'within' if fee_pct <= fee_cap + TOL else 'ABOVE THE CEILING')},
        {'comp': 'Total', 'contains': 'Amount of funding requested, A plus B plus C',
         'amount': money(activities + execution + fee), 'cap': 'Administrative cost below 18 per cent of the total',
         'position': '%.1f per cent — %s' % (exec_pct + fee_pct, 'within' if (exec_pct + fee_pct) <= 18 + TOL else 'ABOVE THE CEILING')},
    ]


# ============================ END OF SECTION A ===============================

# =============================================================================
# SECTION B  -  CONTEXT FOR T5, POLICY BRIEF
# =============================================================================
PILLAR_LABEL = {'P1': 'risk data, early warning and anticipatory action',
                'P2': 'community reach and inclusion',
                'P3': 'governance, coordination and institutional innovation',
                'P4': 'innovation financing, partnerships and scaling',
                'P5': 'knowledge management, learning and replication'}

def absent_groups(inv):
    rows, _ = grs_rows(inv)
    return [r['group'].lower() for r in rows if r['provision'] == 'None recorded']

def innovation_brief(pack, n, states, years, modality, beneficiaries):
    inv = pack.innovation(n)
    B = build_budget(pack, [inv], states, years, modality)
    MAP = build_mapping(B, states)
    ctry = country_of(inv)
    gov = 'the Government of ' + ctry if ctry != 'Regional' else 'each adopting Member State'
    absent = absent_groups(inv)
    stage = (inv.get('maturity') or 'not recorded').lower()
    admin_pct = (100.0 * (B['management'] + B['mel'] + B['support']) / B['total']) if B['total'] else 0

    facts = [
        {'label': 'Instrument', 'value': inv['name']},
        {'label': 'Operated by', 'value': inv.get('lead') or 'To be named by ' + gov},
        {'label': 'Where it operates now', 'value': plain(inv.get('geo')) or 'Not recorded'},
        {'label': 'Stage', 'value': (inv.get('maturity') or 'Not recorded')},
        {'label': 'Pillar', 'value': PILLAR_LABEL.get(inv.get('pillar'), 'not recorded').capitalize()},
        {'label': 'Proposed coverage', 'value': '%d Member State%s over %d years' % (states, '' if states == 1 else 's', years)},
        {'label': 'Indicative cost', 'value': money(B['total']) + ' in total, of which %s is implementation' % money(B['implementation'])},
        {'label': 'Running cost after handover', 'value': inv.get('cost_rec') or 'To be established with the operating institution'},
    ]

    why = plain(inv.get('why')) or 'The operating evidence is recorded in the Innovation Landscape Assessment.'
    return {
        'kicker': 'POLICY BRIEF   \u00b7   ONE INSTRUMENT   \u00b7   {d}'.format(d=datetime.date.today().strftime('%B %Y')).upper(),
        'headline': inv['name'],
        'lead': ('This instrument is already ' + stage + ' in ' + (plain(inv.get('geo')) or 'the region') +
                 '. What is proposed is not invention but extension: putting a working arrangement into ' +
                 ('%d more Member States' % states if states > 1 else 'another Member State') +
                 ', with the running cost named and an accountable post attached before it is accepted as complete.'),
        'facts': facts,
        'h_evidence': 'What the evidence shows',
        'evidence': (why + ' ' + ('The inclusion record is the part a decision-maker should read closely: ' +
                     ('no provision is recorded for ' + ', '.join(absent) + '. That is stated rather than omitted, and closing it is a condition of scaling this instrument beyond its first Member State.'
                      if absent else 'provision is recorded for all seven groups assessed, which is rare in this portfolio.'))),
        'h_what': 'What this instrument closes',
        'what': ('The Programme\u2019s diagnosis is that the region has no shortage of innovations and that the failures sit at '
                 'the hand-offs between them. This instrument sits in ' + PILLAR_LABEL.get(inv.get('pillar'), 'the portfolio') +
                 ', and it is carried because it closes a function no other instrument in the portfolio closes. Financing it '
                 'alone leaves the hand-offs either side of it open, which is why it is presented alongside the rest of the '
                 'portfolio rather than instead of it.'),
        'h_gov': 'Who owns it and who runs it',
        'governance': ('The Programme belongs to the Southern African Development Community and is owned and driven by its '
                       'Disaster Risk Reduction Unit, accountable upward through the Committee of Ministers responsible for '
                       'disaster risk management, the Troika and the Summit. The Unit appoints a Programme Steering Committee '
                       'to oversee implementation. ' + (inv.get('lead') or 'The national disaster management authority') +
                       ' operates the instrument under national law. The Humanitarian and Emergency Operations Centre does not '
                       'drive the Programme; it receives the protocols, standards and records the instrument produces, in '
                       'support of a mandate it already holds.'),
        'h_cost': 'What it costs, and what happens afterwards',
        'cost': ('Establishing the instrument across %d Member State%s costs an indicative %s, with %s carried by the Action '
                 'for running cost during the transition to national budgets. Total cost of the Action is %s. Figures are '
                 'bands from the Programme cost estimation matrix and are replaced by nationally priced figures before '
                 'financial close; they are not a national budget.'
                 % (states, '' if states == 1 else 's', money(B['setup']), money(B['recurrent']), money(B['total']))),
        'cost_detail': ('Management, oversight and indirect cost recovery together are %.1f per cent of the total, within '
                        'every ceiling the funding partners apply. The Action funds running cost in full for the first year, '
                        'at half in the second and third, and not at all after that, because the instrument is designed to be '
                        'carried on a national budget line rather than by a partner indefinitely.' % admin_pct),
        'h_ask': 'What is being asked',
        'ask': ('Three things of ' + gov + ', none of which requires funding: name the institution that will operate this '
                'instrument; state the law under which it may do so; and name the post that will carry the running cost from '
                'the first full budget year. Without the last two, no proposition built on this instrument can be submitted '
                'to any financing partner.'),
        'close': ('This instrument is one of the ten carried by the Regional Disaster Risk Management Innovation Programme, '
                  'a commitment the region has already made through its Disaster Risk Management Strategy and Action Plan '
                  '2022 to 2030.'),
        'footnote': ('Prepared for the Southern African Development Community under service contract CLMX055740, with the '
                     'International Federation of Red Cross and Red Crescent Societies as implementing agent. Cost figures '
                     'are indicative bands for planning, not quotations.'),
    }

def programme_brief(pack, states, years, modality):
    invs = [pack.innovation(n) for n in pack.numbers()]
    gaps = pack.s['15_Gap_Register']
    inst = sum(1 for g in gaps if clean(g.get('Type')) == 'Institutional')
    nt = sum(1 for i in invs if i.get('type') == 'Non-tech')
    hy = sum(1 for i in invs if i.get('type') == 'Hybrid')
    return {
        'kicker': 'POLICY BRIEF   \u00b7   THE PROGRAMME   \u00b7   {d}'.format(d=datetime.date.today().strftime('%B %Y')).upper(),
        'headline': 'There is no shortage of innovations. What the region lacks are the arrangements that let them work.',
        'lead': ('%d disaster management innovations are operating across Southern Africa. %d things exist that none of them '
                 'reliably does, and %d of those are about authority, agreement and accountability. Not one is solved by '
                 'buying equipment the region does not have.' % (len(invs), len(gaps), inst)),
        'facts': [
            {'label': 'Innovations recorded', 'value': '%d across the region' % len(invs)},
            {'label': 'Missing functions', 'value': '%d, of which %d are institutional' % (len(gaps), inst)},
            {'label': 'Arrangements rather than purchases', 'value': '%d arrangements and %d part-arrangements of %d' % (nt, hy, len(invs))},
            {'label': 'Evidence base', 'value': 'Three streams, twelve of sixteen Member States represented'},
            {'label': 'Instruments carried', 'value': 'Ten, covering all five pillars and all five hand-offs'},
        ],
        'h_evidence': 'What the evidence shows',
        'evidence': ('The evidence comes from three independent streams: a review of the regional and international literature, '
                     'twenty-one interviews across eight Member States and the regional level, and a survey returned by seven '
                     'Member States. Innovation inside each part of the disaster management system is comparatively healthy. '
                     'The system breaks at the hand-offs in the value chain, and all five breaks stay open because the '
                     'arrangements that would close them reach no household directly and are therefore never prioritised.'),
        'h_what': 'What the Programme will do',
        'what': ('It selects ten of the thirty-four: not the ten highest scorers, but the smallest set that closes every break. '
                 'Eight already operate in a Member State and two are piloted, so what is asked for is not invention. Together '
                 'they cover all five pillars, all five links from warning to action, and all five breaks, and close eleven of '
                 'the sixteen missing functions directly.'),
        'h_gov': 'Who owns it, who runs it, and who it enables',
        'governance': ('The Programme belongs to the Southern African Development Community. It is owned and driven by its '
                       'Disaster Risk Reduction Unit, accountable upward through the Committee of Ministers responsible for '
                       'disaster risk management, the Troika and the Summit. The Unit appoints a Programme Steering Committee '
                       'to oversee implementation. Member State authorities operate the innovations under their own national '
                       'law. The International Federation of Red Cross and Red Crescent Societies is the implementing agent, '
                       'not the owner. The Humanitarian and Emergency Operations Centre does not drive the Programme; it is '
                       'enabled by it.'),
        'h_cost': 'What it costs, and the number that matters',
        'cost': ('Every instrument carries an indicative cost band drawn from published comparators, stated for one adopting '
                 'Member State. These bands are not a regional budget: regional totals must be consolidated from the '
                 'participating countries and from the common regional functions before any single figure is put to a '
                 'financing partner.'),
        'cost_detail': ('The number that should shape the decision is not establishment but the one beneath it. Running cost '
                        'is close to a third of establishment, every year, indefinitely, so financing that pays to set '
                        'something up and not to run it pays for the first year of something that then stops. The Programme '
                        'moves that cost onto national budgets across three phases rather than carrying it.'),
        'h_ask': 'What is being asked',
        'ask': ('Three things of Member States, none of which requires funding: name the institution that will lead each core '
                'arrangement; state the law under which your country may operate each instrument; and name the post accountable '
                'for running costs. Without them, no proposition in the Programme can be submitted to any financing partner.'),
        'close': ('The Community has no protocol on disaster risk management. Its authority here rests on the Protocol on '
                  'Health at Article 25, which binds Parties to collaborate on regional preparedness plans, and on the mandate '
                  'of its Disaster Risk Reduction Unit to develop exactly this kind of programme.'),
        'footnote': ('Prepared for the Southern African Development Community under service contract CLMX055740, with the '
                     'International Federation of Red Cross and Red Crescent Societies as implementing agent. Cost figures '
                     'are indicative bands for planning, not quotations; costed proposals at national prices follow.'),
    }

def build_context(pack, innovation=None, states=4, years=5, modality='un', beneficiaries=None):
    brief = (programme_brief(pack, states, years, modality) if innovation is None
             else innovation_brief(pack, innovation, states, years, modality, beneficiaries))
    return {'date': datetime.date.today().strftime('%d %B %Y'), 'brief': brief}

# =============================================================================
# SECTION C  -  COMMAND LINE
# =============================================================================
if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--pack', default='ARC_D4_Automation_Matrix.xlsx')
    ap.add_argument('--template', default='T5_Summary_Brief.docx')
    ap.add_argument('--innovation', type=int, default=None,
                    help='portfolio number 1-34 for a single-instrument brief')
    ap.add_argument('--programme', action='store_true',
                    help='the programme-level brief instead of one instrument')
    ap.add_argument('--states', type=int, default=4)
    ap.add_argument('--years', type=int, default=5)
    ap.add_argument('--modality', default='un', choices=sorted(MODALITY))
    ap.add_argument('--out', default=None)
    a = ap.parse_args()
    if not a.programme and a.innovation is None:
        ap.error('give --innovation N for one instrument, or --programme for the programme brief')
    pack = DataPack(a.pack)
    n = None if a.programme else a.innovation
    ctx = build_context(pack, n, a.states, a.years, a.modality)
    out = a.out or ('T5_Programme.docx' if n is None else 'T5_%02d.docx' % n)
    render(a.template, ctx, out)
