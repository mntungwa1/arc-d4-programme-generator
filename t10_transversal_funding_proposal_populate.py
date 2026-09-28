# =============================================================================
#  POPULATE T10  -  TRANSVERSAL FUNDING PROPOSAL
#  SADC Regional Disaster Risk Management Innovation Programme
#  Contract CLMX055740  |  Academy of Resilience and Continuity
# -----------------------------------------------------------------------------
#  Scope   : a transversal Action covering several innovations across pillars
#  Reads   : the D2 / D3 / Cost-Matrix data pack
#  Usage   : python populate_T10.py --innovations 1,2,3,4,6,8,13,15,18,19 --states 8 --years 5
#  Budget  : built from the cost matrix. Contingency is set at 10 per cent by
#            default and never below it. Governance and administration are shown
#            as separate lines so a funder can apply its own ceilings.
# =============================================================================

"""Populate the transversal funding proposal from the D2/D3/cost-matrix data pack."""
# =============================================================================
#  POPULATE T2  -  Project Concept Note
#  SADC Regional Disaster Risk Management Innovation Programme
#  Contract CLMX055740  |  Academy of Resilience and Continuity
# -----------------------------------------------------------------------------
#  Scope   : one innovation
#  Reads   : the D2 / D3 / Cost-Matrix data pack (ARC_D4_Automation_Matrix.xlsx)
#  Template: the T2 .docx shipped with this file
#  Usage   : python populate_T2.py --innovation 1
#  Rule    : nothing is generated. Every value is read from the data pack; an
#            absent value is omitted or written as a request naming the holder.
# =============================================================================

"""Populate Project Concept Note from the D2/D3/cost-matrix data pack."""
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

def country_of(inv):
    """The country where the innovation operates.

    Geography strings name the operating country first and the replication scope
    second, for example 'Malawi; SADC recurrent-hazard MS'. Only the first clause
    decides the country; searching the whole string wrongly returns Regional for
    every nationally operated innovation.
    """
    g = inv.get('geo') or ''
    if not g: return 'Regional'
    first = re.split(r'[;]', g)[0]
    first = re.sub(r'\(.*?\)', '', first).strip()
    if re.search(r'\b(regional|SADC|multi[- ]country|continental|global)\b', first, re.I):
        return 'Regional'
    first = re.split(r'[,(]', first)[0].strip()
    first = re.sub(r'\s+(piloted|operational|scaling|since|from)\b.*$', '', first, flags=re.I).strip()
    first = re.sub(r'\s+[A-Z]{2,6}$', '', first).strip()      # trailing institution code
    first = re.sub(r'\s+\d{4}$', '', first).strip()            # trailing year
    CODES = {'BWA': 'Botswana', 'SWZ': 'Eswatini', 'MWI': 'Malawi', 'MOZ': 'Mozambique',
             'ZAF': 'South Africa', 'ZMB': 'Zambia', 'ZWE': 'Zimbabwe', 'NAM': 'Namibia',
             'AGO': 'Angola', 'LSO': 'Lesotho', 'TZA': 'Tanzania', 'DRC': 'Democratic Republic of the Congo'}
    return CODES.get(first.upper(), first) or 'Regional' 

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
    return ('Validation-ready draft. Generated from the D2, D3 and cost-matrix data pack on '
            + datetime.date.today().strftime('%d %B %Y') + '. Every field the Programme holds is complete. '
            'Fields held by a Member State appear as requests naming who is being asked. Cost figures are '
            'indicative bands, not quotations.')

def render(template_path, context, out_path):
    doc = DocxTemplate(template_path)
    doc.render(context, autoescape=True)
    doc.save(out_path)
    print('written:', out_path)
# ============================ END OF SECTION A ===============================

# =============================================================================
# SECTION B  -  CONTEXT FOR T10, TRANSVERSAL FUNDING PROPOSAL
# =============================================================================
PILLAR_NAME = {'P1': 'Risk data, early warning and anticipatory action',
               'P2': 'Community-centred and inclusive disaster risk management',
               'P3': 'Governance, coordination and institutional innovation',
               'P4': 'Innovation financing, partnerships and scaling',
               'P5': 'Knowledge management, learning and regional replication'}

CONTINGENCY_RATE = 0.10      # never lower; a funder capping it lower is handled in the note
GOVERNANCE_RATE  = 0.08      # governance and administration of the Action
SUPPORT_RATE     = 0.07      # programme support costs on the subtotal

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

def build_budget(pack, invs, states, years):
    """Direct cost by pillar, then governance, contingency and support costs.

    Set-up is incurred once per adopting Member State for a national instrument and
    once for the region for a regional one. Running cost is carried for the years
    inside the Action after the first. Entries that are not sums are excluded from
    the arithmetic and reported separately so the budget is never quietly short.
    """
    by_pillar = {}
    unpriced = []
    for i in invs:
        p = i.get('pillar') or 'P1'
        by_pillar.setdefault(p, {'amount': 0.0, 'names': []})
        by_pillar[p]['names'].append(i['name'])
        amount = 0.0
        for field, span in (('cost_setup', 1), ('cost_rec', max(years - 1, 1))):
            kind, val = parse_band(i.get(field))
            if kind == 'per_state':
                amount += val * states * span
            elif kind == 'regional':
                amount += val * span
            else:
                why, req = UNPRICED_REASON[kind]
                unpriced.append({'inst': i['name'] + (' (set-up)' if field == 'cost_setup' else ' (running cost)'),
                                 'why': why, 'requirement': req})
        by_pillar[p]['amount'] += amount
    direct = sum(v['amount'] for v in by_pillar.values())
    governance = direct * GOVERNANCE_RATE
    subtotal = direct + governance
    contingency = subtotal * CONTINGENCY_RATE
    before_support = subtotal + contingency
    support = before_support * SUPPORT_RATE
    return {'by_pillar': by_pillar, 'unpriced': unpriced, 'direct': direct,
            'governance': governance, 'subtotal': subtotal, 'contingency': contingency,
            'support': support, 'total': before_support + support}

def build_context(pack, numbers, states, years):
    invs = [pack.innovation(n) for n in numbers]
    B = build_budget(pack, invs, states, years)
    arrangements = sum(1 for i in invs if i.get('type') in ('Non-tech', 'Hybrid'))

    # the Action adapts to what it actually carries: one instrument or several,
    # inside one pillar or across them
    n_inst = len(invs); n_pill = len(B['by_pillar'])
    single = (n_inst == 1)
    transversal = (n_pill > 1)
    if single:
        act_title = 'Implementation of ' + (invs[0]['name'] or 'a disaster risk management innovation')
        act_kind = 'a single-instrument Action'
        shape = ('This Action carries one instrument: ' + (invs[0]['name'] or '') + '. It is presented on its own because '
                 'it closes a function that no other instrument in the portfolio closes, and because it can be '
                 'implemented without waiting for the rest of the portfolio to be financed.')
    elif transversal:
        act_title = 'Closing the institutional hand-offs in Southern African disaster risk management'
        act_kind = 'a transversal Action'
        shape = ('This Action is transversal by necessity. Each hand-off runs between two pillars, so an instrument that '
                 'closes one sits in neither pillar alone. Financing the pillars separately funds the parts that already '
                 'work and leaves the joints between them open, which is the failure the evidence identifies.')
    else:
        only = PILLAR_NAME.get(list(B['by_pillar'])[0], 'a single pillar')
        act_title = 'Strengthening ' + only[0].lower() + only[1:]
        act_kind = 'a single-pillar Action'
        shape = ('This Action carries ' + str(n_inst) + ' instruments, all within ' + only + '. It is presented as one '
                 'Action because the instruments depend on each other within that pillar and financing them separately '
                 'would leave those dependencies unmanaged.')

    unpriced_note = ('Every instrument in this Action is banded in the cost estimation matrix and the budget above is complete.'
                     if not B['unpriced'] else
                     'The items below are carried by this Action but cannot be banded from the cost estimation matrix, because '
                     'the matrix prices them as a premium, a ratio or a per-person figure rather than as a sum. They are '
                     'excluded from the arithmetic above rather than entered as zero, and each is priced before financial '
                     'close from the information named in the third column.')

    # inclusion across the set, from the determinations
    tally = {}
    for cat in INCL_CATS:
        designed = named = none = 0
        for i in invs:
            v = (i['grs'] or {}).get(cat)
            try: v = float(v)
            except (TypeError, ValueError): continue
            if v >= 1: designed += 1
            elif v >= 0.5: named += 1
            else: none += 1
        tally[cat] = (designed, named, none)
    incl_rows = []
    for cat, (dg, nm, no) in tally.items():
        if dg == 0 and nm == 0:
            prov = 'No provision recorded in any instrument'
            act  = 'Designed provision is a condition of entry to the scale-up phase'
        elif dg == 0:
            prov = 'Named in %d instruments, designed into none' % nm
            act  = 'Convert naming into designed provision before scale-up'
        else:
            prov = 'Designed into %d, named in %d' % (dg, nm)
            act  = 'Maintain and evidence provision at each acceptance point'
        incl_rows.append({'group': cat, 'provision': prov, 'action': act})

    outputs = []
    for ref, (p, v) in enumerate(sorted(B['by_pillar'].items()), 1):
        outputs.append({'ref': 'Output %d' % ref,
                        'statement': PILLAR_NAME.get(p, p) + ' strengthened through the instruments named',
                        'activities': '; '.join(v['names'])})

    direct_rows = []
    for ref, (p, v) in enumerate(sorted(B['by_pillar'].items()), 1):
        direct_rows.append({'ref': 'Output %d' % ref, 'desc': PILLAR_NAME.get(p, p),
                            'unit': 'Per Member State, %d States' % states,
                            'amount': money(v['amount']), 'share': pct(v['amount'], B['direct'])})

    summary = [
        {'ref': 'A', 'line': 'Direct costs of implementation', 'amount': money(B['direct']),
         'note': 'By output at section 7.3'},
        {'ref': 'B', 'line': 'Governance and administration of the Action', 'amount': money(B['governance']),
         'note': '%.0f per cent of A' % (GOVERNANCE_RATE * 100)},
        {'ref': 'C', 'line': 'Subtotal (A + B)', 'amount': money(B['subtotal']), 'note': ''},
        {'ref': 'D', 'line': 'Contingency reserve', 'amount': money(B['contingency']),
         'note': '%.0f per cent of C' % (CONTINGENCY_RATE * 100)},
        {'ref': 'E', 'line': 'Subtotal (C + D)', 'amount': money(B['subtotal'] + B['contingency']), 'note': ''},
        {'ref': 'F', 'line': 'Programme support costs', 'amount': money(B['support']),
         'note': '%.0f per cent of E' % (SUPPORT_RATE * 100)},
        {'ref': 'G', 'line': 'Total estimated cost of the Action', 'amount': money(B['total']), 'note': 'E + F'},
    ]

    return {
        'date': datetime.date.today().strftime('%d %B %Y'),
        'act': {
            'kicker': ('FUNDING PROPOSAL   \u00b7   ' +
                       ('SINGLE-INSTRUMENT ACTION' if single else
                        ('TRANSVERSAL ACTION' if transversal else 'SINGLE-PILLAR ACTION'))),
            'cover_note': (('This proposal covers one instrument. Figures marked indicative are drawn from the Programme '
                            'cost estimation matrix and are replaced by nationally priced figures before submission.') if single else
                           (('This proposal is transversal: it addresses functions that cut across the pillars of the Programme '
                             'and cannot be resolved inside any one of them. ') if transversal else
                            ('This proposal covers several instruments within one pillar, which depend on each other. ')) +
                           'Figures marked indicative are drawn from the Programme cost estimation matrix and are replaced by '
                           'nationally priced figures before submission.'),
            'title': act_title,
            'subtitle': (('One instrument in %d Member State%s over %d years' % (states, '' if states == 1 else 's', years))
                         if single else
                         ('%d instruments across %d pillar%s in %d Member State%s over %d years'
                          % (n_inst, n_pill, '' if n_pill == 1 else 's', states, '' if states == 1 else 's', years))),
            'applicant': 'Southern African Development Community Secretariat, Disaster Risk Reduction Unit',
            'coapplicants': 'National disaster management authorities of the participating Member States; implementing partners named at section 4.2',
            'modality': 'Regional Action implemented by the Disaster Risk Reduction Unit, with delivery by national institutions under their own law',
            'coverage': '%d Member States of the Southern African Development Community' % states,
            'duration': '%d months' % (years * 12),
            'sector': 'Disaster risk reduction and preparedness; institutional capacity building',
            'sdgs': 'Goal 1.5, Goal 11.5, Goal 11.b and Goal 13.1',
            'sendai': 'Targets B, C, D, E and G, with the contribution to Target G at the access dimension',
            'reference': 'To be assigned on submission',
            'summary': ('The region has no shortage of disaster risk management innovations. Thirty-four are operating or piloted across '
                        'Southern Africa. What fails is the hand-offs between them: the authority to issue a warning, the finance that moves '
                        'before a disaster rather than after it, the channel that carries a warning to a household in a language it uses, and '
                        'the record that lets one Member State copy what works in another. This Action addresses those hand-offs directly. It '
                        'addresses those hand-offs directly. ' + shape + ' %d of the %d instrument%s carried %s an arrangement between '
                        'institutions rather than a purchase of equipment, which is the part of the portfolio that conventional capital '
                        'financing does not reach.' % (arrangements, n_inst, '' if n_inst == 1 else 's', 'is' if n_inst == 1 else 'are')),
            'target_groups': 'National disaster management authorities, sub-national authorities, community structures and the institutions that operate each instrument',
            'final_beneficiaries': 'Populations exposed to hydrometeorological and related hazards in the participating Member States, with priority to the groups identified at section 3.3',
            'reach': 'To be confirmed per Member State from national exposure data before submission',
            'beneficiary_note': ('Inclusion is treated as a determination rather than a commitment. Each group below was assessed across the '
                                 'instruments in this Action from the coded interview record, and absence is stated rather than omitted.'),
        },
        'rel': {
            'context': ('Evidence was assembled from three independent streams and triangulated: a review of the regional and international '
                        'literature, twenty-one key informant interviews across eight Member States and the regional level, and a survey '
                        'returned by seven Member States. Twelve of the sixteen Member States are represented in at least one stream. The '
                        'assessment found that innovation inside each part of the disaster risk management value chain is comparatively '
                        'healthy and that the system fails at the hand-offs between them.'),
            'needs_intro': 'Sixteen functions exist that no institution in the region reliably performs. Ten concern authority, agreement and accountability rather than equipment. Those most directly addressed by this Action are set out below.',
            'needs': [
                {'need': 'No designated authority to issue a public warning in several Member States',
                 'evidence': 'Interview and survey record; legal register of the Draft Programme',
                 'consequence': 'Warning channels are built that no institution may lawfully use'},
                {'need': 'No instrument converts a financing decision into a delivery instruction',
                 'evidence': 'Gap register of the Draft Programme',
                 'consequence': 'Pre-arranged finance arrives without a route to the household'},
                {'need': 'No obligation on any institution to document what operates',
                 'evidence': 'Gap register; absence of a regional record',
                 'consequence': 'What works in one Member State cannot be adopted by another'},
                {'need': 'Recurrent cost is unassigned in most instruments',
                 'evidence': 'Cost estimation matrix; Member State engagement',
                 'consequence': 'Instruments stop when founding financing ends'},
                {'need': 'No instrument records provision for displaced people',
                 'evidence': 'Inclusion determinations extracted from the coded interview record',
                 'consequence': 'The group most directly produced by these hazards is not reached'},
            ],
            'transversal': shape + (' The Action is accountable for the connection between instruments rather than for each part alone.'
                                    if transversal else ''),
            'alignment': [
                {'name': 'Protocol on Health, Article 25', 'how': 'Gives effect to the obligation to collaborate on regional awareness, risk reduction, preparedness and management plans'},
                {'name': 'Protocol on Politics, Defence and Security Cooperation, Article 2', 'how': 'Enhances regional capacity in disaster management and the coordination of humanitarian assistance'},
                {'name': 'Regional Indicative Strategic Development Plan 2020 to 2030', 'how': 'Delivers the disaster risk reduction commitments it carries'},
                {'name': 'Disaster Risk Management Strategy and Action Plan 2022 to 2030', 'how': 'Implements commitments recorded in the Strategy'},
                {'name': 'Sendai Framework for Disaster Risk Reduction', 'how': 'Contributes to Targets B, C, D, E and G, and to the reporting obligation that accompanies them'},
                {'name': 'Sustainable Development Goals', 'how': 'Contributes to Goals 1.5, 11.5, 11.b and 13.1'},
                {'name': 'Africa Regional Strategy for Disaster Risk Reduction', 'how': 'Aligns the regional contribution to the continental programme of action'},
            ],
            'complementarity': ('The Action works through instruments that already operate and through institutions that already hold the mandate. '
                                'It does not create parallel structures, does not duplicate an existing regional mechanism, and is coordinated '
                                'through the Disaster Risk Reduction Unit, which holds the portfolio and the record. Where an instrument is '
                                'externally capitalised, the Action finances the arrangement around it rather than the instrument itself.'),
        },
        'logic': {
            'overall': 'Reduced disaster mortality, displacement and economic loss in the participating Member States',
            'specific': 'The hand-offs between risk information, authority, financing, delivery and learning function reliably across the participating Member States',
            'outputs_short': '; '.join(o['ref'] + ': ' + o['statement'] for o in outputs),
            'outputs': outputs,
            'toc': ('The Action rests on a contribution logic rather than an attribution claim. Risk information produces a trigger; governance '
                    'authorises a decision; financing releases funds; community systems deliver the action and give the trigger its legitimacy; '
                    'and learning recalibrates the trigger. The Action intervenes at the points between these links. Its claim is that it makes '
                    'a difference the pathway would not otherwise show, and each evaluation tests the assumptions explicitly and reports which failed.'),
            'matrix': [
                {'level': 'Impact', 'result': 'Reduced disaster mortality, displacement and economic loss',
                 'indicator': 'Sendai Framework Targets A, B and C', 'verification': 'Member State returns to the Sendai monitoring system',
                 'assumption': 'Hazard exposure does not change beyond the range observed'},
                {'level': 'Outcome', 'result': 'The hand-offs function reliably',
                 'indicator': 'Number of hand-offs with a designated institution and a working instrument, of five',
                 'verification': 'Programme record; Member State confirmation', 'assumption': 'Member States designate authority where it is absent'},
                {'level': 'Output', 'result': 'Instruments implemented and operating',
                 'indicator': 'Instruments operating to specification, by Member State',
                 'verification': 'Host institution confirmation at acceptance', 'assumption': 'Gating arrangements are in place before delivery'},
                {'level': 'Output', 'result': 'Inclusion provision designed in',
                 'indicator': 'Determinations completed across seven groups, including recorded absences',
                 'verification': 'Programme record', 'assumption': 'Data can be disaggregated at source'},
                {'level': 'Output', 'result': 'Recurrent cost assigned',
                 'indicator': 'Instruments with a named accountable post and a budget line',
                 'verification': 'Member State budget documentation', 'assumption': 'Budget cycles allow a new line within the Action period'},
            ],
        },
        'incl': {'rows': incl_rows},
        'impl': {
            'modality': ('The Action is implemented by the Disaster Risk Reduction Unit of the Secretariat, which owns and drives it and is '
                         'accountable upward through the Committee of Ministers responsible for disaster risk management, the Troika and the '
                         'Summit. The Unit appoints a Programme Steering Committee to oversee implementation. Instruments are operated by '
                         'national institutions under their own national law; the Action supplies the arrangement, the standard and the '
                         'evidence, and does not operate inside a Member State.'),
            'partners': [
                {'partner': 'Disaster Risk Reduction Unit', 'role': 'Owner and driver of the Action; contracting authority; holder of the record',
                 'basis': 'Institutional mandate to develop frameworks and programmes and to coordinate implementation and evaluation'},
                {'partner': 'Programme Steering Committee', 'role': 'Oversight of implementation under delegated authority',
                 'basis': 'Appointed by the Unit'},
                {'partner': 'National disaster management authorities', 'role': 'Operate the instruments; confirm acceptance; carry recurrent cost after transition',
                 'basis': 'National law and the adoption plan agreed per Member State'},
                {'partner': 'Humanitarian and Emergency Operations Centre', 'role': 'Receives protocols, standards and records in support of its own mandate',
                 'basis': 'Enabled by the Action; holds no Action decision and no reporting line'},
                {'partner': 'Implementing and technical partners', 'role': 'Deliver against terms of reference issued per instrument',
                 'basis': 'Competitive procurement by the Unit'},
            ],
            'workplan': [
                {'period': 'Year 1', 'focus': 'Quick wins and pilots; arrangements that gate other instruments',
                 'outputs': 'All outputs initiated', 'milestone': 'Gating arrangements in place; first instruments operating'},
                {'period': 'Years 2 to 3', 'focus': 'Scale-up and replication across Member States',
                 'outputs': 'All outputs advanced', 'milestone': 'Provision for displaced people designed in before scale-up'},
                {'period': 'Years 4 to 5', 'focus': 'Institutionalisation in policy and budgets',
                 'outputs': 'Sustainability conditions met', 'milestone': 'Named accountable post and budget line per instrument'},
            ],
            'mel': ('Monitoring is continuous and reads from records that already exist or that the instruments themselves produce. Evaluation '
                    'is separate and occurs three times: a baseline within the first six months, a mid-term evaluation at the end of year three '
                    'which is the decision point for reshaping the portfolio, and a final evaluation in year five. Evaluation uses contribution '
                    'analysis and tests the assumptions in the logical framework explicitly.'),
            'monitoring': [
                {'level': 'Impact', 'what': 'Mortality, displacement and economic loss', 'by': 'National authorities through existing Sendai returns', 'freq': 'Annual'},
                {'level': 'Outcome', 'what': 'Hand-offs functioning', 'by': 'Programme record with Member State confirmation', 'freq': 'Semi-annual'},
                {'level': 'Output', 'what': 'Instruments operating; inclusion determinations; recurrent cost assigned', 'by': 'Host institutions and the Programme record', 'freq': 'Quarterly'},
                {'level': 'Process', 'what': 'Evidence integrity: sources cited, derivations confirmed, requests distinguished from gaps', 'by': 'Programme record', 'freq': 'Continuous'},
            ],
            'risks': [
                {'risk': 'An instrument is scaled before the arrangement that gates it exists', 'rating': 'Extreme', 'owner': 'Steering Committee',
                 'mitigation': 'Sequencing rule: no instrument enters scale-up until its gating arrangement is in place'},
                {'risk': 'Community-held knowledge is used without written authorisation', 'rating': 'Extreme', 'owner': 'Disaster Risk Reduction Unit',
                 'mitigation': 'Written consent, attribution and benefit sharing are conditions precedent; delivery halts where absent'},
                {'risk': 'Recurrent cost is not assigned before the Action closes', 'rating': 'High', 'owner': 'Member States',
                 'mitigation': 'Named post required before any instrument is accepted as complete; standing item of the Steering Committee'},
                {'risk': 'Legal provision cannot be confirmed in a Member State', 'rating': 'High', 'owner': 'Member States',
                 'mitigation': 'Instrument referred to the governance workstream rather than dropped'},
                {'risk': 'Provision for displaced people is not designed in', 'rating': 'High', 'owner': 'Disaster Risk Reduction Unit',
                 'mitigation': 'Condition of entry to the scale-up phase, carried in the implementation matrix'},
                {'risk': 'Currency movement or inflation erodes the budget', 'rating': 'Medium', 'owner': 'Disaster Risk Reduction Unit',
                 'mitigation': 'Contingency reserve at section 7.5; annual budget review with the funding partner'},
            ],
        },
        'cross': {
            'intro': ('The commitments below are treated as obligations of delivery rather than as statements of intent. Each is evidenced at an '
                      'acceptance point and reported against.'),
            'commitments': [
                {'item': 'Gender equality and the empowerment of women and girls',
                 'how': 'Inclusion is determined instrument by instrument against seven groups, women among them, and provision is evidenced at acceptance rather than asserted in design'},
                {'item': 'Leave no one behind',
                 'how': 'The Action records the groups no instrument currently reaches and makes designed provision a condition of entry to the scale-up phase. Displaced people are the priority case'},
                {'item': 'Human rights-based approach',
                 'how': 'Community-held knowledge is used only under written authorisation, with attribution and benefit sharing agreed before access and ownership not transferring'},
                {'item': 'Accountability to affected populations',
                 'how': 'Community structures confirm that warnings produced an action, and that confirmation is part of the monitoring record rather than a separate exercise'},
                {'item': 'Do no harm and conflict sensitivity',
                 'how': 'Instruments are tested against the mandate of the institutions that will operate them before delivery, so no capability is created that no institution may lawfully use'},
                {'item': 'Protection from sexual exploitation and abuse',
                 'how': 'Binding on every implementing and technical partner through the conditions of contract, with safeguarding obligations where work touches communities'},
                {'item': 'Environmental and social safeguards',
                 'how': 'Applied to any instrument engaging water use, land or community infrastructure, under the law of the Member State concerned'},
                {'item': 'Data protection',
                 'how': 'Beneficiary information is handled under the law of the Member State in which it is collected and is not transferred without written authority'},
            ],
        },
        'sus': {
            'intro': ('Sustainability is treated as a sequence with conditions rather than as an expectation. The Action is designed to end, and '
                      'the arrangements it establishes are designed to continue without it.'),
            'institutional': ('Each instrument is operated by a national institution under its own law, named in national policy, and assigned to '
                              'a post rather than to an institution in the abstract. The Action supplies the arrangement, the standard and the '
                              'evidence; it does not operate anything itself, so there is no structure to dismantle at closure.'),
            'financial': ('Running cost across this portfolio is close to a third of establishment cost, every year, indefinitely. That figure is '
                          'the determinant of sustainability and it is transitioned deliberately across three phases rather than left to the '
                          'closing months of the Action.'),
            'transition': [
                {'phase': 'Year 1', 'who': 'The Action carries running cost in full',
                 'condition': 'The instrument is operating and the host institution can use it unaided'},
                {'phase': 'Years 2 to 3', 'who': 'A national budget line is created alongside the Action contribution',
                 'condition': 'The line exists in the national budget and a post is named against it'},
                {'phase': 'Years 4 to 5', 'who': 'The national budget line carries the full recurrent cost',
                 'condition': 'An instrument that has not begun this transition by the end of year three is reported to the Steering Committee as at risk'},
            ],
            'exit': ('Capability transfer is a condition of final payment on every implementation contract. The host institution must be able to '
                     'operate the instrument unaided, with documentation, training and a stated annual running cost. The Programme record, the '
                     'templates and the populating code remain with the Disaster Risk Reduction Unit as an asset that continues to generate '
                     'propositions after the Action closes.'),
        },
        'budget': {
            'total_action': money(B['total']),
            'requested': money(B['total']) + ' (co-financing to be confirmed at section 7.8)',
            'cofinancing': 'Member State contributions in kind through staff time, existing structures and the recurrent budget lines created during the Action',
            'basis': ('Costs are necessary for the Action, reasonable, verifiable and incurred during implementation. Direct costs are built from '
                      'the Programme cost estimation matrix, which bands every instrument against named published comparators. Those bands are '
                      'indicative: they support appraisal and budgeting and are replaced by nationally priced figures before financial close. '
                      'Governance, contingency and support costs are shown as separate lines so that a funding partner may apply its own ceilings '
                      'without reworking the direct costs.'),
            'direct': direct_rows,
            'by_output': [{'ref': r['ref'], 'desc': r['desc'], 'amount': r['amount'], 'share': r['share']} for r in direct_rows],
            'governance_note': ('Governance and administration are budgeted at %.0f per cent of direct costs. They are shown explicitly rather than '
                                'absorbed, because the Action is regional and its oversight is the mechanism by which a funding partner\u2019s '
                                'contribution is controlled.' % (GOVERNANCE_RATE * 100)),
            'governance': [
                {'item': 'Programme management', 'covers': 'Action management within the Disaster Risk Reduction Unit: planning, contracting, supervision and reporting',
                 'amount': money(B['governance'] * 0.45)},
                {'item': 'Steering Committee and oversight', 'covers': 'Meetings of the Programme Steering Committee, Member State focal point coordination and Board reporting',
                 'amount': money(B['governance'] * 0.20)},
                {'item': 'Monitoring, evaluation and learning', 'covers': 'Baseline, mid-term and final evaluation, and the monitoring system',
                 'amount': money(B['governance'] * 0.20)},
                {'item': 'Audit and assurance', 'covers': 'External audit and verification of expenditure',
                 'amount': money(B['governance'] * 0.15)},
            ],
            'contingency_note': ('A contingency reserve of %.0f per cent of direct and governance costs is included, equal to %s. It covers currency '
                                 'movement, inflation in national markets, and the replacement of indicative bands with nationally priced figures. '
                                 'It is released only on the written authority of the Disaster Risk Reduction Unit and is reported on separately. '
                                 'Where a funding partner caps a contingency reserve below this level, the balance is carried as a separately '
                                 'justified risk provision inside direct costs, and the Unit will present it that way on request.'
                                 % (CONTINGENCY_RATE * 100, money(B['contingency']))),
            'summary': summary,
            'unpriced_note': unpriced_note,
            'unpriced': B['unpriced'],
            'vfm': ('Three features carry the value-for-money case. The Action extends instruments that already operate rather than financing '
                    'untested designs, which lowers delivery risk. It concentrates on arrangements between institutions, which are the cheapest '
                    'instruments in the portfolio to establish and to run and which reach the most households. And it requires a named post and a '
                    'budget line for recurrent cost before an instrument is accepted as complete, so the contribution buys a continuing capability '
                    'rather than a period of activity.'),
            'cofin': [
                {'source': 'Participating Member States', 'nature': 'Staff time, existing community and district structures, and the recurrent budget lines created during the Action', 'value': 'To be quantified per Member State before submission'},
                {'source': 'Southern African Development Community Secretariat', 'nature': 'Disaster Risk Reduction Unit staff time and regional convening', 'value': 'To be quantified before submission'},
                {'source': 'Implementing and technical partners', 'nature': 'Co-financed instruments already operating in the region', 'value': 'To be quantified before submission'},
            ],
        },
        'vis': {'text': ('Communication and visibility follow the funding partner\u2019s guidelines and are agreed in a plan within the first three '
                         'months. Visibility is treated as an obligation of the Action and is budgeted within governance and administration. All '
                         'material identifies the Action as an initiative of the Southern African Development Community, delivered by its Disaster '
                         'Risk Reduction Unit with the support of the funding partner. Where an instrument produces public warnings, visibility '
                         'requirements never take precedence over the clarity or speed of the warning itself.')},
        'annexes': [
            {'n': '1', 'name': 'Logical framework matrix', 'status': 'At section 3.5; supplied separately in the funding partner\u2019s format on request'},
            {'n': '2', 'name': 'Detailed budget', 'status': 'Supplied in the funding partner\u2019s format; indicative pending national pricing'},
            {'n': '3', 'name': 'Innovation Landscape Assessment', 'status': 'Available'},
            {'n': '4', 'name': 'Draft Programme and annexes, including the gap and legal registers', 'status': 'Available'},
            {'n': '5', 'name': 'Cost estimation matrix', 'status': 'Available; internal, supplied on request to the funding partner only'},
            {'n': '6', 'name': 'Inclusion determinations', 'status': 'Available'},
            {'n': '7', 'name': 'Member State adoption plans', 'status': 'Prepared per Member State on confirmation of participation'},
        ],
        'decl': {'text': ('The applicant confirms that the information in this proposal is accurate, that the figures marked indicative are drawn '
                          'from the Programme cost estimation matrix and will be replaced by nationally priced figures before financial close, and '
                          'that the Action does not duplicate an existing regional mechanism.'),
                 'prepared_by': 'Programme, from the data pack',
                 'approved_by': 'Disaster Risk Reduction Unit'},
    }

# =============================================================================
# SECTION C  -  COMMAND LINE
# =============================================================================
if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--pack', default='ARC_D4_Automation_Matrix.xlsx')
    ap.add_argument('--template', default='T10_Transversal_Funding_Proposal.docx')
    ap.add_argument('--innovations', default='1,2,3,4,6,8,13,15,18,19',
                    help='comma-separated portfolio numbers carried by the Action')
    ap.add_argument('--states', type=int, default=8, help='Member States adopting')
    ap.add_argument('--years', type=int, default=5, help='duration of the Action in years')
    ap.add_argument('--out', default='T10_Funding_Proposal.docx')
    a = ap.parse_args()
    nums = [int(x) for x in a.innovations.split(',') if x.strip()]
    pack = DataPack(a.pack)
    render(a.template, build_context(pack, nums, a.states, a.years), a.out)
