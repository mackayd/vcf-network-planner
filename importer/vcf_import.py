"""Import entered Broadcom VCF 9.1.1 planning data into Firewall Planner.

Python 3.10+ standard library. Writing/recalculation uses desktop Excel on Windows
through the companion PowerShell script. Source files are read as ZIP/XML only.
"""
from __future__ import annotations
import argparse
import collections
import datetime
import hashlib
import ipaddress
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
q = lambda s: '{' + NS + '}' + s
SENSITIVE = re.compile(r'password|passphrase|credential|secret|token|license|licence|private.?key|access.?key|api.?key|activation.?code|username|user_name|_user$|_pwd|_pass$', re.I)
SOURCE_SHEETS = {'Deploy Management Domain', 'Configure Management Domain', 'Deploy Workload Domain', 'Configure Workload Domain', 'Deploy Cluster', 'Deploy Fleet Management Day-N', 'VCF & VVF Planning'}
UNMAPPED_SHEETS = SOURCE_SHEETS | {'Additional Racks'}
PLACEHOLDERS = {'value missing', 'not specified', 'n/a', 'not applicable', 'none', 'tbd', 'enter value', '<enter value>'}

def colnum(s):
    n = 0
    for c in s: n = n * 26 + ord(c) - 64
    return n

def colname(n):
    s = ''
    while n:
        n, r = divmod(n-1, 26); s = chr(65+r) + s
    return s

def cells_in_range(ref):
    a, _, b = ref.replace('$','').partition(':'); b = b or a
    m = re.fullmatch(r'([A-Z]+)(\d+)', a); z = re.fullmatch(r'([A-Z]+)(\d+)', b)
    if not m or not z: raise ValueError('Only finite A1 ranges are supported')
    for r in range(int(m[2]), int(z[2])+1):
        for c in range(colnum(m[1]), colnum(z[1])+1): yield f'{colname(c)}{r}'

class Xlsx:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(path) as z:
            if sum(i.file_size for i in z.infolist()) > 300_000_000:
                raise ValueError('Workbook is larger than the supported expanded size')
            self.parts = {n:z.read(n) for n in z.namelist()}
        self.workbook = ET.fromstring(self.parts['xl/workbook.xml'])
        rels = ET.fromstring(self.parts['xl/_rels/workbook.xml.rels'])
        targets = {r.get('Id'):r.get('Target') for r in rels if r.get('TargetMode') != 'External'}
        self.sheets = {}
        self.sheet_paths = {}
        for sheet in self.workbook.find(q('sheets')):
            name = sheet.get('name'); target = targets[sheet.get('{'+REL+'}id')]
            path = target.lstrip('/') if target.startswith('/') else 'xl/'+target
            self.sheet_paths[name] = path
            self.sheets[name] = ET.fromstring(self.parts[path])
        self.strings = []
        if 'xl/sharedStrings.xml' in self.parts:
            self.strings = [''.join(n.itertext()) for n in ET.fromstring(self.parts['xl/sharedStrings.xml'])]
        self.maps = {s:{c.get('r'):c for c in root.iter(q('c'))} for s,root in self.sheets.items()}
        self.names = {}
        self.duplicate_names = set()
        for n in self.workbook.findall(q('definedNames')+'/'+q('definedName')):
            name = n.get('name','').lower()
            if name in self.names: self.duplicate_names.add(name)
            self.names[name] = n.text or ''
        styles = ET.fromstring(self.parts['xl/styles.xml'])
        self.styles = list(styles.find(q('cellXfs')))

    def value(self, sheet, cell):
        c = self.maps[sheet].get(cell)
        if c is None: return None
        if c.get('t') == 'inlineStr': return ''.join(c.find(q('is')).itertext())
        v = c.findtext(q('v'))
        if v is None: return None
        if c.get('t') == 's': return self.strings[int(v)]
        if c.get('t') in ('str','e'): return v
        if c.get('t') == 'b': return bool(int(v))
        try:
            f = float(v); return int(f) if f.is_integer() else f
        except ValueError: return v

    def formula(self,sheet,cell):
        c = self.maps[sheet].get(cell)
        return c is not None and c.find(q('f')) is not None

    def unlocked(self,sheet,cell):
        c = self.maps[sheet].get(cell)
        style = self.styles[int(c.get('s','0')) if c is not None else 0]
        p = style.find(q('protection'))
        return p is not None and p.get('locked','1') == '0'

    def named_ref(self,name):
        name = name.lower()
        if name in self.duplicate_names: raise ValueError(f'Ambiguous defined name: {name}')
        f = self.names.get(name,'')
        m = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+)", f)
        if not m: return None
        return (m[1].replace("''", "'") if m[1] else m[2], m[3]+m[4])

class Source:
    def __init__(self,path,overrides=None):
        self.book = Xlsx(path)
        if not {'Deploy Management Domain','Deploy Workload Domain','Deploy Cluster'}.issubset(self.book.sheets):
            raise ValueError('Unsupported P&P workbook: expected VCF 9.1.1 deployment worksheets')
        if 'input_mgmt_az1_mgmt_gateway_cidr' not in self.book.names:
            raise ValueError('Unsupported P&P schema: named network input fields are missing')
        self.used = {}; self.notices = []; self.seen_notices = set()
        self.overrides=overrides or {}
        if not isinstance(self.overrides,dict):raise ValueError('source_overrides must be a JSON object')
        for name in self.overrides:
            ref=self.book.named_ref(name)
            if not name.startswith('input_') or SENSITIVE.search(name) or not ref or ref[0] not in SOURCE_SHEETS:
                raise ValueError('Source override is not a supported non-sensitive visible input field')
        if self.overrides:self.issue('SOURCE_OVERRIDES','Explicit operator corrections from the config replace these inputs for this import only; the source file is unchanged.',self.overrides.keys())

    def issue(self,code,message,refs=()):
        key = (code,message,tuple(refs))
        if key not in self.seen_notices:
            self.seen_notices.add(key); self.notices.append({'code':code,'message':message,'sourceFields':list(refs)})

    def get(self,name,mark=True):
        if SENSITIVE.search(name): raise ValueError('Sensitive source field is outside importer scope')
        ref = self.book.named_ref(name)
        if not ref: return None
        sheet,cell = ref
        if sheet not in SOURCE_SHEETS: return None
        v = self.overrides[name] if name in self.overrides else self.book.value(sheet,cell)
        if name not in self.overrides and self.book.formula(sheet,cell):
            self.issue('FORMULA_INPUT', 'Formula-based input was skipped; replace it with a confirmed value or review manually.', [name])
            return None
        if v is None or v == '' or isinstance(v,str) and v.strip().lower() in PLACEHOLDERS: return None
        if not isinstance(v,(str,int,float,bool)):raise ValueError('Input overrides must be literal strings, numbers or null')
        if isinstance(v,str):
            v = v.strip()
            if not v: return None
            if v.startswith(('=','+','@')) or v.startswith('-') and not re.fullmatch(r'-?\d+(\.\d+)?',v):
                self.issue('UNSAFE_LITERAL','Input text beginning with a spreadsheet expression marker was skipped.',[name]); return None
            if len(v)>300: self.issue('LONG_INPUT','Oversized input was skipped; review the field manually.',[name]); return None
        if mark: self.used[name] = {'sheet':sheet,'cell':cell}
        return v

    def first(self,*names):
        found = [(n,self.get(n)) for n in names]
        found = [(n,v) for n,v in found if v is not None]
        if not found:return None
        if len({str(v) for n,v in found})>1:
            self.issue('CONFLICTING_INPUTS','Alternative input fields disagree; this value was not imported.',[n for n,v in found]); return None
        return found[0][1]

    def populated_inputs(self):
        out = {}
        for name in self.book.names:
            if not name.startswith('input_') or SENSITIVE.search(name): continue
            ref = self.book.named_ref(name)
            if ref and ref[0] in UNMAPPED_SHEETS and not self.book.formula(*ref):
                value = self.book.value(*ref)
                if value is not None and value != '': out[name] = ref
        return out

    def refs(self,*names):
        return ', '.join(f'{n} ({self.book.named_ref(n)[0]}!{self.book.named_ref(n)[1]})' for n in names if self.book.named_ref(n))

def ipv4(value):
    return str(ipaddress.IPv4Address(str(value).strip()))

def gateway_network(value):
    if '/' not in str(value): raise ValueError('Expected gateway/prefix or gateway/netmask')
    interface = ipaddress.IPv4Interface(str(value).strip())
    if interface.network.prefixlen>30: raise ValueError('Planner VLAN networks require prefix lengths 1–30')
    if interface.network.prefixlen<1 or interface.ip in (interface.network.network_address,interface.network.broadcast_address):
        raise ValueError('Gateway must be a usable host address')
    return str(interface.network.network_address),interface.network.prefixlen,str(interface.ip)

def clean_int(value,low,high):
    if isinstance(value,bool): raise ValueError('Expected integer')
    f = float(value)
    if not f.is_integer() or not low <= f <= high: raise ValueError('Integer outside supported range')
    return int(f)

def fqdn(value):
    s = str(value).strip().rstrip('.')
    labels=s.split('.')
    if len(s)>253 or len(labels)<2 or any(not re.fullmatch(r'[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?',label) for label in labels):
        raise ValueError('Invalid host name')
    return s.lower()

PURPOSES = {'mgmt':(0,[0,1,2]),'mgmt_vm':(1,[3]),'vcf_mgmt':(2,[4]),
            'vmotion':(3,[5,6,7]),'vsan':(4,[8,9,10]),'host_overlay':(5,[11,12,13]),
            'edge_mgmt':(6,[14,15,16]),'edge_overlay':(7,[17,18,19]),
            'uplink01':(8,[20,21,22]),'uplink02':(9,[23,24,25]),
            'witness':(10,[26]),'nfs':(11,[27,28,29])}
ACTIVE_OFF = {'exclude','unselected','no','disabled','not selected'}
BLUEPRINTS = {'Single site - minimal','Single site - standard','Multiple sites - one region','Multiple regions','Stretched region + other regions'}

class Planner:
    def __init__(self,source,template,config=None):
        self.source = source; self.template = Xlsx(template); self.config = config or {}
        self.edits = {}; self.resources = []; self.networks = []; self.clusters = []; self.domains = []
        self.manual_row = 139; self.external_networks = {}; self.nsx_groups = []; self.hosts = {}
        self.mtu_values = collections.defaultdict(set)
        self.interface_assignments = self.read_interface_assignments()
        self.used_interface_assignments = set()
        for sheet in ['Design','Domains & clusters','Deployment','Network plan','Resources','Firewall matrix']:
            if sheet not in self.template.sheets: raise ValueError('Unsupported firewall template: missing '+sheet)
        for sheet,cell,value in [('Resources','R8','Witness managed by vCenter ID'),('Deployment','I19','NET-C1-01'),('Domains & clusters','B8','DOM-01'),('Domains & clusters','B17','CL-01')]:
            if self.template.value(sheet,cell)!=value: raise ValueError('Unsupported firewall template schema; use v6.1')
        self.clear = []
        for sheet in ['Design','Domains & clusters','Deployment','Network plan','Resources']:
            for cell in self.template.maps[sheet]:
                if self.template.unlocked(sheet,cell):
                    if self.template.formula(sheet,cell): raise ValueError('Unexpected editable formula in template')
                    self.clear.append({'sheet':sheet,'cell':cell})
        self.allowed = {(c['sheet'],c['cell']) for c in self.clear}

    def put(self,sheet,cell,value):
        if value is None:return
        if (sheet,cell) not in self.allowed: raise ValueError('Target is not an input: '+sheet+'!'+cell)
        self.edits[(sheet,cell)] = value

    def get_config(self,section,key,default=None):
        return self.config.get(section,{}).get(key,default)

    def warn(self,code,msg,refs=()): self.source.issue(code,msg,refs)

    def validated(self,name,validator):
        v = self.source.get(name)
        if v is None:return None
        try:return validator(v)
        except (ValueError,TypeError):
            self.warn('INVALID_VALUE','Invalid value was not imported; correct this source field.',[name]);return None

    def defaults(self):
        if self.config.get('confirmation_mode','Required')!='Required' and ('Design','G6') not in self.allowed:raise ValueError('Confirmation mode requires the v6.2 template')
        if ('Design','G6') in self.allowed:
            mode=self.config.get('confirmation_mode','Required')
            if mode not in ('Required','Skip confirmations'):raise ValueError('Invalid confirmation_mode')
            self.put('Design','G6',mode)
        for cell in ['C22','G22','G23','C24','G24','C25']:self.put('Design',cell,'No')
        for r in range(8,12):
            for col,value in [('A','No'),('D','Management' if r==8 else 'Workload'),('J','No')]:self.put('Domains & clusters',f'{col}{r}',value)
        for r in range(17,21):
            for col,value in [('A','No'),('J',0),('O','None')]:self.put('Domains & clusters',f'{col}{r}',value)
            if r>17:
                for col,value in [('K',0),('Q','No')]:self.put('Domains & clusters',f'{col}{r}',value)
        for r in range(26,30):
            for col in 'BFGHI':
                if ('Domains & clusters',f'{col}{r}') in self.allowed:self.put('Domains & clusters',f'{col}{r}','No')
        for r in range(35,39):
            for col,value in [('A','No'),('F','No')]:self.put('Domains & clusters',f'{col}{r}',value)
        for r in range(44,49):
            self.put('Domains & clusters',f'A{r}','Yes' if r in (44,45) else 'No')
        self.put('Domains & clusters','D44','Imported platform services; review required')
        self.put('Domains & clusters','D45','Imported DNS / NTP / external endpoints; review required')
        self.put('Domains & clusters','G44','EXT-01')
        for r in range(9,169):self.put('Resources',f'A{r}','No')
        for r in range(139,163):self.put('Deployment',f'A{r}','No')
        for cell in ['D8','G8','I8','I9','D10','G10','I10','D11','I11','D12']:self.put('Deployment',cell,'No')
        self.put('Deployment','I7','Live deployment')
        self.put('Deployment','G11',443)
        self.put('Deployment','D5','Imported P&P data — review incomplete fields')
        self.put('Deployment','D6','Initial import; source inputs are not an approved firewall design')

    def meaningful(self,prefix):
        exact = ['name','domain_name','vc_fqdn','nsxt_vip_fqdn'] if prefix=='cluster' else ['sddc_domain','cl01_cluster','vc_fqdn','vc_ip']
        tests=[f'input_{prefix}_{x}' for x in exact]
        pattern=re.compile(r'^input_'+re.escape(prefix)+r'_(?:az[12]_(?:host\d+_(?:fqdn|mgmt_ip)|(?:mgmt|mgmt_vm|vcf_mgmt|vmotion|vsan|host_overlay|principal_storage)_(?:vlan|cidr|gateway_cidr|network|mask|pool_start_ip|pool_end_ip|gateway_ip))|nsxt_(?:mgr[abc]|vip)_(?:fqdn|ip))$')
        tests += [name for name in self.source.book.names if pattern.fullmatch(name)]
        return any(self.source.get(n,False) is not None for n in tests)

    def storage(self,prefix):
        n=f'{prefix}_principal_storage_chosen';v=self.source.get(n)
        if v is None:return None
        t=str(v).lower()
        if 'fibre channel' in t or t in ('fc','vmfs on fc'):return 'External FC'
        if 'vsan' in t and 'compute' not in t and 'storage cluster' not in t:return 'vSAN HCI'
        if 'nfs' in t and re.search(r'4[.]1',t):return 'NFS 4.1'
        self.warn('STORAGE_WORKFLOW','Storage workflow is not fully supported by this planner; review storage networking manually.',[n])
        return 'Other - manual design'

    def network_model(self,prefix):
        names=[f'{prefix}_nsx_vpc_type_chosen',f'{prefix}_vpc_type_chosen']
        v=self.source.first(*names)
        if v is not None:
            t=str(v).lower()
            if 'full stack' in t:return 'Full stack overlay'
            if 'vlan' in t:return 'VLAN-backed VPC'
        configured=self.get_config(prefix,'network_model')
        if configured in ('Full stack overlay','VLAN-backed VPC'):return configured
        self.warn('NETWORK_MODEL','Select the domain network model in Domains & clusters; it was not determined from explicit source inputs.',names)
        return None

    def cluster(self,prefix,index,domain,active):
        row=17+index; ident=f'CL-0{index+1}'
        name=self.source.get('input_cluster_name' if prefix=='cluster' else f'input_{prefix}_cl01_cluster')
        stretched=self.source.get(f'{prefix}_stretched_cluster_chosen')
        has_az2=any(self.source.get(f'input_{prefix}_az2_host{i}_{field}',False) is not None for i in range(1,17) for field in ('mgmt_ip','fqdn'))
        has_az2=has_az2 or any(self.source.get(f'input_{prefix}_az2_mgmt_{field}',False) is not None for field in ('vlan','cidr','gateway_ip','gateway_cidr'))
        azs=2 if has_az2 or str(stretched).lower() in {'selected','include','yes','enabled'} else 1
        if has_az2 and str(stretched).lower() in ACTIVE_OFF:
            self.warn('AZ_SELECTOR_CONFLICT','AZ2 host entries exist while the stretch selector is off. AZ2 data is retained; confirm the cluster design.',[f'{prefix}_stretched_cluster_chosen'])
        data={'prefix':prefix,'index':index,'id':ident,'domain':domain,'active':active,'azs':azs,'name':name,'storage':self.storage(prefix),'networks':{}}
        self.clusters.append(data)
        self.put('Domains & clusters',f'A{row}','Yes' if active else 'No')
        self.put('Domains & clusters',f'C{row}',domain)
        if index==0:
            self.put('Design','G13',name);self.put('Design','G11',data['storage'] if data['storage']!='Other - manual design' else None)
            chosen=self.get_config('design','blueprint')
            if chosen and chosen not in BLUEPRINTS:raise ValueError('Unsupported blueprint in config')
            self.put('Design','C6',chosen or ('Multiple sites - one region' if azs==2 else 'Single site - standard'))
            self.put('Design','C10',self.get_config('design','data_sites',azs))
            for c,key in [('G7','fleets'),('G8','instances'),('G9','regions')]:self.put('Design',c,self.get_config('design',key,1))
            for c,key in [('C7','fleet'),('C9','region')]:self.put('Design',c,self.get_config('design',key))
            self.put('Design','C12',self.get_config(prefix,'infrastructure','AZ-local subnets'))
            for c,key in [('C17','az1_site'),('C18','az2_site'),('C19','witness_site')]:self.put('Design',c,self.get_config(prefix,key))
            self.put('Design','E17','AZ1');self.put('Design','E18','AZ2' if azs==2 else None)
            self.warn('FLEET_SCOPE_REVIEW','The blueprint and fleet counts describe a provisional local pattern. Confirm fleet/region/site details on Design or provide a config file.')
        else:
            for col,val in [('D',name),('E',azs),('F',self.get_config(prefix,'az1_site')),('G',self.get_config(prefix,'az2_site')),('H',self.get_config(prefix,'infrastructure','AZ-local subnets')),('I',data['storage']),('N',self.get_config(prefix,'witness_site'))]:self.put('Domains & clusters',f'{col}{row}',val)
        if not self.get_config(prefix,'az1_site') or azs==2 and not self.get_config(prefix,'az2_site'):
            self.warn('SITE_NAMES_REQUIRED','Physical site names are missing from the import configuration; complete the cluster site inputs. Datacenter names are not assumed to be physical sites.',[prefix])
        if prefix!='mgmt':
            racks=self.source.get(f'{prefix}_multi_rack_count_chosen')
            if racks is not None and str(racks).lower() not in ACTIVE_OFF | {'1','single rack'}:
                self.warn('ADDITIONAL_RACKS','Only the base rack is imported. Additional Racks inputs need a separate mapping and are listed as unsupported.',[f'{prefix}_multi_rack_count_chosen'])
        return data

    def domain(self,prefix,index,active):
        row=8+index;ident=f'DOM-0{index+1}';name=self.source.get(f'input_{prefix}_sddc_domain')
        data={'prefix':prefix,'index':index,'id':ident,'name':name,'active':active,'model':self.network_model(prefix)}
        self.domains.append(data)
        for col,val in [('A','Yes' if active else 'No'),('C',name),('D','Management' if index==0 else 'Workload'),('E',f'vc-{prefix}'),('G',f'CL-0{index+1}'),('H',data['model']),('I','EXT-01'),('J','No'),('L','Imported from entered P&P fields. Review topology and controller placement.')]:self.put('Domains & clusters',f'{col}{row}',val)
        return data

    def read_network(self,base):
        vals={suffix:self.source.get('input_'+base+'_'+suffix) for suffix in ['vlan','gateway_cidr','cidr','network','mask','gateway_ip','mtu','pool_start_ip','pool_end_ip']}
        if not any(v is not None for k,v in vals.items() if k!='mtu'):return None
        out={'base':base,'refs':['input_'+base+'_'+k for k,v in vals.items() if v is not None],**vals}
        try:
            if vals['gateway_cidr'] is not None:
                interface=ipaddress.IPv4Interface(re.sub(r'\s+','',str(vals['gateway_cidr'])))
                if interface.network.prefixlen<31 and interface.ip in (interface.network.network_address,interface.network.broadcast_address):
                    subnet,prefix,gw=str(interface.network.network_address),interface.network.prefixlen,None
                    self.warn('GATEWAY_NOT_HOST','The gateway/CIDR input contains a network or broadcast address. Its subnet and prefix were imported; the gateway is left blank for correction.',['input_'+base+'_gateway_cidr'])
                else:subnet,prefix,gw=gateway_network(str(interface))
            else:
                cidr=vals['cidr'];gw=ipv4(vals['gateway_ip']) if vals['gateway_ip'] is not None else None
                if cidr is not None and '/' in str(cidr):n=ipaddress.IPv4Network(str(cidr).strip(),strict=True)
                elif vals['network'] is not None and (vals['mask'] is not None or cidr is not None):n=ipaddress.IPv4Network(f"{vals['network']}/{vals['mask'] or cidr}",strict=True)
                elif gw and cidr is not None and str(cidr).isdigit():n=ipaddress.IPv4Interface(f'{gw}/{cidr}').network
                else:raise ValueError('Subnet or prefix missing')
                subnet,prefix=str(n.network_address),n.prefixlen
                if gw and (ipaddress.IPv4Address(gw) not in n or prefix<31 and ipaddress.IPv4Address(gw) in (n.network_address,n.broadcast_address)):raise ValueError('Gateway outside subnet')
            out.update(subnet=subnet,prefix=prefix,gateway=gw)
        except ValueError:
            out.update(subnet=None,prefix=None,gateway=None)
            self.warn('NETWORK_ADDRESS_REVIEW','Subnet, mask or gateway could not be validated. Valid VLAN/MTU entries are retained; complete the network address manually.',out['refs'])
        for key,low,high in [('vlan',1,4094),('mtu',576,65535)]:
            try:out[key]=clean_int(vals[key],low,high) if vals[key] is not None else None
            except (ValueError,TypeError):out[key]=None;self.warn('INVALID_'+key.upper(),'Invalid '+key+' was not imported.', ['input_'+base+'_'+key])
        return out

    def add_network(self,cluster,purpose,az,base):
        net=self.read_network(base)
        if not net:return None
        p,offsets=PURPOSES[purpose];index=cluster['index'];offset=offsets[min(az-1,len(offsets)-1)]
        row=19+index*30+offset;key=f'NET-C{index+1}-{offset+1:02d}'
        net.update(row=row,key=key,cluster=cluster['id'],purpose=purpose,az=az)
        self.networks.append(net);cluster['networks'][(purpose,az)]=net
        local_name=self.source.first(f'input_{cluster["prefix"]}_cl01_az{az}_{purpose}_pg',f'input_{base}_pg')
        for c,v in [('D',net['vlan']),('E',net['subnet']),('F',net['prefix']),('G',net['gateway']),('H','VCF'),('K',f'P&P input: {base}. Confirm L2/routing and physical site.'),('L',local_name)]:self.put('Deployment',f'{c}{row}',v)
        planrow=8+12*index+p
        self.put('Network plan',f'{"K" if az==1 else "L"}{planrow}',key)
        if net['mtu'] is not None:self.mtu_values[planrow].add(net['mtu'])
        return net

    def resource(self,ident,role,fn=None,ip=None,network=None,end=None,cluster=None,domain=None,az=1,nsx=None,service=None,hosted=None,active=True,interface='Management',notes='',location=None):
        if len(self.resources)>=160:raise ValueError('Resource capacity exceeded (160 interfaces); no workbook written')
        row=9+len(self.resources)
        if fn:
            try:fn=fqdn(fn)
            except ValueError:self.warn('INVALID_FQDN','Invalid resource hostname was omitted; check the source named field.',[ident]);fn=None
        if ip:
            try:ip=ipv4(ip)
            except ValueError:self.warn('INVALID_IP','Invalid or unsupported resource IP was omitted.',[ident]);ip=None
        if role=='ESXi Management' and fn:
            if fn in self.hosts:
                self.warn('DUPLICATE_HOST','Repeated host FQDN was not added twice; reconcile cluster/AZ/IP ownership.',[ident]);return
            self.hosts[fn]=(cluster,az,ip)
        if network and ip and network.get('subnet'):
            if ipaddress.IPv4Address(ip) not in ipaddress.IPv4Network(f"{network['subnet']}/{network['prefix']}"):
                self.warn('IP_NETWORK_MISMATCH','Assigned address does not belong to its source-selected network. The entry is retained for planner validation.',[ident,network['base']])
        record={'A':'Yes' if active else 'No','B':ident,'C':role,'D':fn,'E':interface,'F':network['key'] if network else None,'G':ip,'H':end,'I':'No','K':notes or 'Imported from P&P. Confirm address, allocation and placement.','L':location or f'AZ{az}','M':domain,'N':cluster,'O':hosted,'P':nsx,'Q':service}
        for col,value in record.items():self.put('Resources',f'{col}{row}',value)
        self.resources.append({'row':row,'id':ident,'role':role,'cluster':cluster,'domain':domain,'active':active,'pool':end is not None,'hasFqdn':fn is not None,'hasIp':ip is not None})
        if active and (not ip or (not fn and role!='Runtime Pool')):
            self.warn('RESOURCE_INCOMPLETE','Resource has no explicit FQDN or IPv4 in P&P. Imported available fields; fill the missing details before rule approval.',[ident])

    def pool(self,net,role,cluster,domain,active=True,assignments=False):
        start,end=net['pool_start_ip'],net['pool_end_ip']
        if start is None and end is None:return
        try:
            start=ipv4(start);end=ipv4(end)
            if int(ipaddress.IPv4Address(start))>int(ipaddress.IPv4Address(end)):raise ValueError('Reversed pool')
            if net.get('subnet'):
                n=ipaddress.IPv4Network(f"{net['subnet']}/{net['prefix']}")
                if ipaddress.IPv4Address(start) not in n or ipaddress.IPv4Address(end) not in n:raise ValueError('Pool outside subnet')
                if n.prefixlen<31 and (ipaddress.IPv4Address(start)==n.network_address or ipaddress.IPv4Address(end)==n.broadcast_address):raise ValueError('Pool includes network/broadcast')
                if net.get('gateway') and int(ipaddress.IPv4Address(start))<=int(ipaddress.IPv4Address(net['gateway']))<=int(ipaddress.IPv4Address(end)):raise ValueError('Pool contains gateway')
        except (ValueError,TypeError):
            self.warn('POOL_REVIEW','Pool bounds are incomplete, reversed or outside the subnet. No individual addresses were allocated. Review the source range labels; one vendor AZ2 storage mapping has reversed named endpoints.',net['refs']);return
        self.resource(net['base']+'-pool',role,ip=start,end=end,network=net,cluster=cluster['id'] if role.startswith('ESXi') else None,domain=domain if role.startswith('ESXi') else None,az=net['az'],hosted='CL-01' if role=='Runtime Pool' else None,service='PLAT-01' if role=='Runtime Pool' else None,active=active,interface='Allocation pool',notes='Source allocation range, not assigned host interfaces. No addresses were allocated by this importer.')
        if role.startswith('ESXi') and not assignments:
            self.warn('HOST_INTERFACES_REQUIRED','Only an allocation pool is available. Add actual per-host VMkernel addresses; this range does not count as assigned host interfaces.',[net['base']])

    def read_interface_assignments(self):
        """Read explicit assignments only; never distribute a source pool."""
        records=list(self.config.get('host_interfaces',[]))
        book=self.source.book
        if 'Example design' in book.sheets:
            in_schedule=False
            for row in range(1,1000):
                a,b=book.value('Example design',f'A{row}'),book.value('Example design',f'B{row}')
                if a=='Host allocation schedule' and b=='FQDN; management; vMotion; vSAN; TEP addresses':
                    in_schedule=True;continue
                if not in_schedule:continue
                if a is None and b is None:break
                if book.formula('Example design',f'A{row}') or book.formula('Example design',f'B{row}'):raise ValueError('Host allocation schedule must contain literal inputs')
                m=re.fullmatch(r'Mgmt ([0-9.]+); vMotion ([0-9.]+); vSAN ([0-9.]+); TEP ([0-9.]+) / ([0-9.]+)',str(b))
                if not m:raise ValueError('Unrecognized explicit host allocation row')
                for purpose,address,label in [('vmotion',m[2],'vmk-vmotion'),('vsan',m[3],'vmk-vsan'),('host_overlay',m[4],'tep1'),('host_overlay',m[5],'tep2')]:
                    records.append({'fqdn':a,'purpose':purpose,'ip':address,'interface':label,'management_ip':m[1],'source':f'Example design!A{row}:B{row}'})
        result={};seen=set()
        for item in records:
            fn=fqdn(item['fqdn']);purpose=item['purpose'];address=ipv4(item['ip'])
            if purpose not in ('vmotion','vsan','host_overlay','nfs'):raise ValueError('Unsupported host interface purpose')
            label=str(item.get('interface') or purpose)
            key=(fn,label)
            if key in seen:raise ValueError('Duplicate explicit host interface assignment')
            seen.add(key);record=dict(item,fqdn=fn,ip=address,interface=label)
            result.setdefault((fn,purpose),[]).append(record)
        return result

    def assigned_interfaces(self,c,az,purpose,net,role):
        assigned=[];expected=0
        for i in range(1,17):
            base=f'input_{c["prefix"]}_az{az}_host{i}'
            fn=self.source.get(base+'_fqdn');mgmt=self.source.get(base+'_mgmt_ip')
            if not fn:continue
            expected+=1;fn=fqdn(fn)
            items=self.interface_assignments.get((fn,purpose),[])
            if items:self.used_interface_assignments.add((fn,purpose))
            if not items:
                value=self.source.get(base+'_'+purpose+'_ip')
                if value is not None:items=[{'ip':value,'interface':purpose,'source':base+'_'+purpose+'_ip'}]
            for item in items:
                address=ipv4(item['ip']);n=ipaddress.IPv4Network(f"{net['subnet']}/{net['prefix']}")
                if ipaddress.IPv4Address(address) not in n or address in (str(n.network_address),str(n.broadcast_address),net.get('gateway')):raise ValueError('Explicit host interface outside usable source network')
                if item.get('management_ip') and ipv4(item['management_ip'])!=ipv4(mgmt):raise ValueError('Host allocation schedule disagrees with management input')
                if net.get('pool_start_ip') and net.get('pool_end_ip') and not (int(ipaddress.IPv4Address(net['pool_start_ip']))<=int(ipaddress.IPv4Address(address))<=int(ipaddress.IPv4Address(net['pool_end_ip']))):raise ValueError('Explicit host interface outside its allocation pool')
                assigned.append((i,fn,address,item))
        if len({fn for _,fn,_,_ in assigned}) not in (0,expected):self.warn('PARTIAL_HOST_INTERFACES','Explicit addresses cover only some hosts. Pool retained inactive; complete missing host interfaces.',[net['base']])
        if len({a for _,_,a,_ in assigned})!=len(assigned):raise ValueError('Duplicate explicit host interface IP')
        for i,fn,address,item in assigned:
            self.resource(f'{c["prefix"]}-a{az}-h{i}-{item["interface"]}',role,fn,address,net,cluster=c['id'],domain=c['domain'],az=az,active=c['active'],interface=item['interface'],notes='Explicit source assignment: '+str(item.get('source','import configuration')))
        return bool(assigned)

    def external_network(self,ip,role):
        try:ip=ipv4(ip)
        except (ValueError,TypeError):return None
        if ip in self.external_networks:return self.external_networks[ip]
        if self.manual_row>162:raise ValueError('Manual network capacity exceeded; no workbook written')
        row=self.manual_row;self.manual_row+=1;name=f'Endpoint {row-138}'
        key=f'External / {name} / External'
        for col,val in [('A','Yes'),('B','External'),('C',name),('D','External'),('E',ip),('F',32),('H','External endpoints'),('K','Host /32 object from explicit source IP. This is not an inferred physical subnet.')]:self.put('Deployment',f'{col}{row}',val)
        net={'key':key,'subnet':ip,'prefix':32,'base':role,'row':row,'purpose':'external'}
        self.external_networks[ip]=net
        return net

    def cluster_data(self,c):
        prefix=c['prefix'];active=c['active'];index=c['index'];domain=c['domain']
        purposes=[('mgmt','mgmt'),('vmotion','vmotion'),('host_overlay','host_overlay')]
        if prefix=='mgmt':purposes += [('mgmt_vm','mgmt_vm'),('vcf_mgmt','vcf_mgmt'),('vsan','vsan')]
        elif c['storage'] in ('vSAN HCI','NFS 4.1'):
            purposes += [('vsan' if c['storage']=='vSAN HCI' else 'nfs','principal_storage')]
        for az in range(1,c['azs']+1):
            for purpose,suffix in purposes:
                if az==2 and purpose in ('mgmt_vm','vcf_mgmt'):continue
                net=self.add_network(c,purpose,az,f'{prefix}_az{az}_{suffix}')
                if net and purpose in ('vmotion','vsan','host_overlay','nfs'):
                    role={'vmotion':'ESXi vMotion','vsan':'ESXi vSAN','host_overlay':'ESXi TEP','nfs':'ESXi NFS'}[purpose]
                    assigned=self.assigned_interfaces(c,az,purpose,net,role)
                    self.pool(net,role,c,domain,active and not assigned,assignments=assigned)
            hostcount=0
            for i in range(1,17):
                base=f'input_{prefix}_az{az}_host{i}';fn=self.source.get(base+'_fqdn');ip=self.source.get(base+'_mgmt_ip')
                if fn is None and ip is None:continue
                self.resource(f'{prefix}-a{az}-h{i}', 'ESXi Management', fn,ip,c['networks'].get(('mgmt',az)),cluster=c['id'],domain=domain,az=az,active=active,interface='vmk0',notes=f'P&P named inputs: {base}_fqdn and {base}_mgmt_ip.')
                hostcount+=1
            self.put('Domains & clusters',f'{"J" if az==1 else "K"}{17+index}',hostcount) if index>0 or az==1 else None
            if az==2 and index==0:
                self.warn('PLANNED_HOST_COUNTS','Management AZ2 planned host count follows the planner design formulas; compare recorded hosts with the intended stretch design.')
        if c['azs']==2:
            self.warn('STRETCH_REVIEW','Stretched infrastructure, routing, witness and cross-AZ evidence are not certified by an import. Review both AZ network plans and witness placement.',[prefix])

    def controller_data(self,d,c):
        prefix=d['prefix'];active=d['active'];net=self.clusters[0]['networks'].get(('mgmt_vm',1))
        fn=self.source.get(f'input_{prefix}_vc_fqdn');ip=self.source.get(f'input_{prefix}_vc_ip')
        if fn is not None or ip is not None:self.resource(f'vc-{prefix}','vCenter',fn,ip,net,domain=d['id'],hosted='CL-01',active=active,notes='P&P vCenter. Served domain imported; hosting on the management cluster requires confirmation.')
        nodes=[]
        for kind in ['mgra','mgrb','mgrc','vip']:
            base=f'input_{prefix}_nsxt_{kind}';f=self.source.get(base+'_fqdn');a=self.source.get(base+'_ip')
            if f is not None or a is not None:nodes.append((kind,f,a))
        if not nodes:
            self.warn('NSX_BINDING_REQUIRED','No explicit NSX endpoint set was found for this domain. Select its actual NSX group; sharing is not assumed.',[prefix]);return
        signature=tuple(sorted((str(f).lower(),str(a)) for k,f,a in nodes))
        existing=next((x for x in self.nsx_groups if x['signature']==signature),None)
        if existing:
            self.put('Domains & clusters',f'F{8+d["index"]}',existing['id']);return
        group=f'NSX-0{len(self.nsx_groups)+1}';row=35+len(self.nsx_groups)
        if row>38:raise ValueError('NSX group capacity exceeded')
        self.nsx_groups.append({'id':group,'signature':signature})
        for col,val in [('A','Yes' if active else 'No'),('C','Imported '+prefix+' NSX'),('D','CL-01'),('E',sum(k!='vip' for k,f,a in nodes)),('F','No'),('H','Confirm shared NSX compatibility and physical placement.'),('I','EXT-01')]:self.put('Domains & clusters',f'{col}{row}',val)
        self.put('Domains & clusters',f'F{8+d["index"]}',group)
        for kind,fn,ip in nodes:self.resource(prefix+'-nsx-'+kind,'NSX VIP' if kind=='vip' else 'NSX Manager',fn,ip,net,nsx=group,hosted='CL-01',active=active)

    def platform_data(self):
        c=self.clusters[0];vmnet=c['networks'].get(('mgmt_vm',1));runtime=c['networks'].get(('vcf_mgmt',1))
        for name,role in [('sddc_mgr','SDDC Manager'),('vcf_installer','Installer'),('flt_lc','License Server')]:
            fn=self.source.get(f'input_{name}_fqdn');ip=self.source.get(f'input_{name}_ip')
            if fn is None and ip is None:continue
            shared=name=='flt_lc' and 'additional' in str(self.source.get('mgmt_domain_chosen')).lower()
            outside=shared or (name=='vcf_installer' and 'outside' in str(self.source.get('mgmt_vcf_installer_location_chosen')).lower())
            net=self.external_network(ip,role) if outside else vmnet
            self.resource(name,role,fn,ip,net,hosted=None if outside else 'CL-01',service='PLAT-01',location='Shared service' if shared else ('External' if outside else 'AZ1'),notes='Existing fleet licence service referenced by an additional instance; explicit address retained as a /32 endpoint.' if shared else ('P&P explicit outside-management placement; planner hosting binding needs manual review.' if outside else 'P&P named appliance. Confirm its network and physical placement.'))
            if outside and not shared:self.warn('EXTERNAL_INSTALLER','P&P places the Installer outside the management domain. Its address is imported as an external host object; planner platform hosting validation needs manual review.',['mgmt_vcf_installer_location_chosen'])
        if runtime:
            runtime['pool_start_ip']=self.source.get('input_flt_vcfms_node_pool_start_ip');runtime['pool_end_ip']=self.source.get('input_flt_vcfms_node_pool_end_ip')
            self.pool(runtime,'Runtime Pool',c,None)
        for name in ['flt_fc','flt_ic','flt_vidb_vip','flt_sr']:
            if self.source.get('input_'+name+'_fqdn') is not None:
                self.warn('RUNTIME_ROLE_REVIEW','Named runtime service has no safe one-to-one planner role mapping. The runtime pool is imported when available; assign individual service roles and addresses manually.',['input_'+name+'_fqdn'])
        deferred=str(self.source.get('mgmt_domain_ops_automation_later_chosen')).lower() in {'selected','yes','include'}
        existing_ops=str(self.source.get('mgmt_domain_existing_vcf_operations_chosen')).lower() in {'selected','yes','include'}
        for name,role in [('xreg_vrops_nodea','Operations'),('xreg_vrops_nodeb','Operations'),('xreg_vrops_nodec','Operations'),('xreg_vrops_collector','Cloud Proxy')]:
            fn=self.source.get('input_'+name+'_fqdn');ip=self.source.get('input_'+name+'_ip')
            if fn is None and ip is None:continue
            self.resource(name,role,fn,ip,self.external_network(ip,role) if existing_ops else vmnet,hosted=None if existing_ops else 'CL-01',service='PLAT-01',active=not deferred or existing_ops,location='External' if existing_ops else 'AZ1',notes='Explicit P&P Operations endpoint. Confirm physical hosting and deferred deployment choice.')
            if deferred and not existing_ops:self.warn('OPERATIONS_DEFERRED','Operations endpoint inputs were retained inactive because the source defers Operations/Automation deployment.',['mgmt_domain_ops_automation_later_chosen'])
            if existing_ops:self.warn('EXISTING_OPERATIONS_PLACEMENT','Existing Operations endpoints were imported, but their hosting is unresolved in the current planner platform model.',['mgmt_domain_existing_vcf_operations_chosen'])
        if self.source.get('input_xreg_vrops_virtual_fqdn') is not None or self.source.get('input_xreg_vrops_virtual_ip') is not None:
            self.warn('OPERATIONS_VIP_ROLE','Operations virtual endpoint requires a distinct VIP binding; it was not misclassified as another Operations node.',['input_xreg_vrops_virtual_fqdn','input_xreg_vrops_virtual_ip'])
        external_services={}
        for kind,role in [('dns','DNS'),('ntp','NTP')]:
            for i in (1,2):
                name=f'input_region_{kind}{i}_'+('ip' if kind=='dns' else 'server');value=self.source.get(name)
                if value is None:continue
                try:ip=ipv4(value);fn=None
                except ValueError:ip=None;fn=value
                key=('ip',ip) if ip else ('fqdn',str(fn).casefold())
                item=external_services.setdefault(key,{'ip':ip,'fqdn':fn,'roles':set(),'fields':[]})
                item['roles'].add(role);item['fields'].append(name)
        for i,item in enumerate(external_services.values(),1):
            role='DNS / NTP' if len(item['roles'])==2 else next(iter(item['roles']))
            self.resource(f'external-service-{i}',role,item['fqdn'],item['ip'],self.external_network(item['ip'],role),service='EXT-01',location='External',interface=role,notes='P&P '+', '.join(item['fields'])+'. /32 is an endpoint only. Identical DNS/NTP endpoints are combined.')
        depot=str(self.source.get('mgmt_offline_depot_chosen') or '').lower()
        if depot=='online':self.put('Deployment','D8','Yes')
        elif 'offline' in depot:self.put('Deployment','D11','Yes')
        depot_name=self.source.get('input_mgmt_offline_depot_fqdn')
        if depot_name is not None:
            self.resource('offline-depot','Offline Depot',depot_name,service='EXT-01',location='External',active='offline' in depot,notes='P&P offline depot FQDN. Supply resolved service IP and network; import performs no DNS lookup.')
            port=self.validated('input_mgmt_offline_depot_port',lambda x:clean_int(x,1,65535))
            if port is not None:self.put('Deployment','G11',port)
        self.warn('SERVICE_ENDPOINTS_REVIEW','Confirm public depot/licensing, admin, logging and optional feature endpoints. Unentered services are not invented; generic example endpoints were cleared.')

    def build(self):
        self.defaults()
        version=self.source.get('vcf_version_chosen')
        if version is None or not re.fullmatch(r'9\.1\.1(?:\.\d+)?',str(version).strip()):raise ValueError('Missing or unsupported source release; this mapping is tested for VCF 9.1.1 / 9.1.1.x')
        if not self.meaningful('mgmt'):raise ValueError('No management-domain identity, hosts or networks found in Your Value fields. No workbook written.')
        for cell,node in self.source.book.maps['Deploy Management Domain'].items():
            formula=node.findtext(q('f'))
            if formula=='mgmt_az1_mgmt_gateway_ip' and cell.startswith('L') and self.source.book.value('Deploy Management Domain','J'+cell[1:])=='Gateway CIDR':
                self.warn('SOURCE_SUMMARY_PREFIX','The locked Gateway CIDR summary references the IP-only named value and drops the prefix. Import uses input_mgmt_az1_mgmt_gateway_cidr, preserving the original prefix. The source summary formula should reference mgmt_az1_mgmt_gateway_cidr.',['Deploy Management Domain!'+cell,'input_mgmt_az1_mgmt_gateway_cidr'])
        self.put('Design','C8',self.source.get('input_vcf_instance_name'))
        pos=str(self.source.get('mgmt_domain_chosen') or '')
        self.put('Design','C13','Additional instance' if 'additional' in pos.lower() else 'First instance')
        d=self.domain('mgmt',0,True);c=self.cluster('mgmt',0,d['id'],True);self.cluster_data(c)
        self.controller_data(d,c)
        workload_present=self.meaningful('wld')
        if workload_present:
            gate=self.source.get('wld_domain_chosen');active=str(gate).lower() not in ACTIVE_OFF
            if not active:self.warn('WORKLOAD_EXCLUDED','Workload entries exist but the source selector is Exclude/Unselected. They were collected into inactive planner rows; confirm intent before enabling.',['wld_domain_chosen'])
            d=self.domain('wld',1,active);c=self.cluster('wld',1,d['id'],active);self.cluster_data(c);self.controller_data(d,c)
        else:self.warn('MANAGEMENT_ONLY','Only management-domain data was found. No workload domain was created. Populate Deploy Workload Domain and rerun to collect workload networking and hosts.')
        if self.meaningful('cluster'):
            owner=self.source.get('input_cluster_domain_name')
            matches=[d for d in self.domains if d['name'] and owner and str(d['name']).casefold()==str(owner).casefold()]
            domain=matches[0]['id'] if len(matches)==1 else None
            active=domain is not None and matches[0]['active'] and str(self.source.get('cluster_chosen')).lower() not in ACTIVE_OFF
            if domain is None:self.warn('CLUSTER_OWNER_UNRESOLVED','Additional cluster has no unambiguous imported owning domain. Data was collected into inactive, unbound rows; select the correct domain before enabling.',['input_cluster_domain_name'])
            if domain:
                parent=matches[0]['prefix']
                for extra,base in [('input_cluster_vc_fqdn',f'input_{parent}_vc_fqdn'),('input_cluster_nsxt_vip_fqdn',f'input_{parent}_nsxt_vip_fqdn')]:
                    a,b=self.source.get(extra),self.source.get(base)
                    if a is not None and b is not None and str(a).casefold()!=str(b).casefold():
                        active=False;self.warn('CLUSTER_CONTROLLER_CONFLICT','Additional cluster controller identity disagrees with its named owning domain. Data was retained inactive for reconciliation.',[extra,base])
            c=self.cluster('cluster',len(self.clusters),domain,active);self.cluster_data(c)
        self.platform_data()
        for row,values in self.mtu_values.items():
            if len(values)==1:self.put('Network plan',f'G{row}',next(iter(values)))
            else:self.warn('MTU_CONFLICT','AZs provide different MTUs for one planner purpose; select the verified end-to-end MTU manually.',[f'Network plan!G{row}'])
        self.put('Deployment','I8','Yes' if any(d['model']=='Full stack overlay' for d in self.domains if d['active']) else 'No')
        # Connectivity and all engineering/port approval inputs intentionally remain unconfirmed.
        self.warn('IMPORT_REVIEW','Import is not design approval. Complete site/fleet details, connectivity, placement, missing interfaces and port review before using firewall rules.')
        unsupported=[{'name':name,'sheet':ref[0],'cell':ref[1]} for name,ref in self.source.populated_inputs().items() if name not in self.source.used]
        if unsupported:self.warn('UNMAPPED_INPUTS',f'{len(unsupported)} populated non-sensitive input fields are outside this mapping. Their names and source cells are listed in the report.')
        if set(self.interface_assignments)-self.used_interface_assignments:raise ValueError('Explicit host interfaces reference an unknown host or an unavailable network purpose')
        return {'schemaVersion':1,'clearInputs':True,'clearCells':self.clear,'edits':[{'sheet':s,'cell':c,'value':v} for (s,c),v in self.edits.items()],
                'report':{'sourceRelease':version,'status':'Imported data requires engineering review','managementOnly':not workload_present,
                          'counts':{'domainsCollected':len(self.domains),'activeDomains':sum(d['active'] for d in self.domains),'clustersCollected':len(self.clusters),'networks':len(self.networks),'externalHostObjects':len(self.external_networks),'resourceRows':len(self.resources),'hostManagementRows':sum(r['role']=='ESXi Management' for r in self.resources),'allocationPools':sum(r['pool'] for r in self.resources)},
                          'warnings':self.source.notices,'sourceFieldsRead':self.source.used,'unmappedInputs':unsupported,'resourceRows':self.resources,
                          'scope':'One management domain, one workload domain and one additional cluster from this P&P instance. Additional racks, brownfield conversion, Edge/BGP, Supervisor, recovery and specialized storage workflows require manual mapping.',
                          'notes':['Reads explicitly named Your Value fields and approved visible selectors; never sample columns or hidden canonical caches.','No passwords, user accounts, licence keys or other authentication data are imported.','No per-host addresses are allocated from pools; no DNS queries or connections to infrastructure are made.','Generic input examples are cleared from the new copy. Source and template stay unchanged.']}}

def write_report(report,path):
    path.write_text(json.dumps(report,indent=2,ensure_ascii=False),encoding='utf-8')
    lines=['VCF planning workbook import report','',report['status'],'',f"Management-only source: {'Yes' if report['managementOnly'] else 'No'}",'']
    lines += [f'{k}: {v}' for k,v in report['counts'].items()]
    lines += ['', 'Review items:']
    for w in report['warnings']:
        lines.append(f"- [{w['code']}] {w['message']}" + (' Fields: '+', '.join(w['sourceFields']) if w['sourceFields'] else ''))
    lines += ['', 'Unmapped populated fields (names and cells only):']
    lines += [f"- {x['name']}: {x['sheet']}!{x['cell']}" for x in report['unmappedInputs']]
    lines += ['',report['scope'],'']+report['notes']
    path.with_suffix('.txt').write_text('\n'.join(lines)+'\n',encoding='utf-8')

def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',required=True,type=Path)
    parser.add_argument('--template',required=True,type=Path)
    parser.add_argument('--output',required=True,type=Path)
    parser.add_argument('--config',type=Path,help='Optional JSON overrides for site names and fleet/design choices')
    parser.add_argument('--dry-run',action='store_true',help='Inspect/map and produce reports only; do not start Excel or create a planner')
    args=parser.parse_args(argv)
    source=args.source.resolve();template=args.template.resolve();output=args.output.resolve()
    if source==template or output in (source,template):raise ValueError('Source, template and output must be distinct files')
    if output.suffix.lower()!='.xlsx':raise ValueError('Output must be .xlsx')
    reportpath=output.with_suffix('.import-report.json')
    viewreportpath=output.with_suffix('.views-report.json')
    if any(p.exists() for p in (output,reportpath,reportpath.with_suffix('.txt'),viewreportpath)):raise ValueError('Output/report already exists; choose a new output name')
    config=json.loads(args.config.read_text(encoding='utf-8-sig')) if args.config else {}
    if not isinstance(config,dict):raise ValueError('Config must be a JSON object')
    for section in ('design','mgmt','wld','cluster'):
        if section in config and not isinstance(config[section],dict):raise ValueError('Config sections must be JSON objects')
    plan=Planner(Source(source,config.get('source_overrides')),template,config).build()
    output.parent.mkdir(parents=True,exist_ok=True)
    report=plan.pop('report');report['runMode']='Dry run' if args.dry_run else 'Workbook import'
    report['createdUtc']=datetime.datetime.now(datetime.timezone.utc).isoformat()
    report['sourceFile']=source.name
    report['sourceSha256']=hashlib.sha256(source.read_bytes()).hexdigest()
    if not args.dry_run:
        if os.name!='nt':raise ValueError('Writing requires Windows with desktop Excel; --dry-run works on other platforms')
        writer=Path(__file__).with_name('Apply-FirewallImport.ps1')
        if not writer.exists():raise ValueError('Companion PowerShell writer is missing')
        refresher=Path(__file__).with_name('Update-FirewallViews.ps1')
        if not all(refresher.with_name(name).exists() for name in ('Update-FirewallViews.ps1','EngineerViews.xlsx','BlueprintDiagram.ps1')):
            raise ValueError('View refresher, BlueprintDiagram.ps1 or EngineerViews.xlsx scaffold is missing')
        with tempfile.TemporaryDirectory(prefix='vcf-firewall-import-') as tmp:
            planpath=Path(tmp)/'plan.json';planpath.write_text(json.dumps(plan,ensure_ascii=False),encoding='utf-8')
            seeded=Path(tmp)/(output.stem+'.seeded.xlsx')
            enhanced=Path(tmp)/output.name
            shell=shutil.which('pwsh')
            if not shell:raise ValueError('PowerShell 7 (pwsh) is required for writing; install it or use --dry-run')
            result=subprocess.run([shell,'-NoProfile','-NonInteractive','-File',str(writer),'-TemplatePath',str(template),'-PlanPath',str(planpath),'-OutputPath',str(seeded)],text=True,encoding='utf-8',capture_output=True)
            if result.returncode:
                try:
                    failure=json.loads(result.stdout.strip().lstrip('\ufeff'))
                    detail=failure.get('message','')+' Stage: '+failure.get('stage','unknown')+' Diagnostics: '+json.dumps({k:v for k,v in failure.items() if k not in ('message','success')})
                except ValueError:detail=result.stderr.strip()[-1200:]
                raise RuntimeError('Excel import failed. '+detail)
            report['excelVerification']=json.loads(result.stdout.strip().lstrip('\ufeff'))
            result=subprocess.run([shell,'-NoProfile','-NonInteractive','-File',str(refresher),'-WorkbookPath',str(seeded),'-OutputWorkbook',str(enhanced)],text=True,encoding='utf-8',capture_output=True)
            if result.returncode:
                raise RuntimeError('Engineer-view refresh failed; no final workbook published. '+result.stderr.strip()[-1600:])
            staged_viewreport=enhanced.with_suffix('.views-report.json')
            view_report=json.loads(staged_viewreport.read_text(encoding='utf-8-sig'))
            if (not view_report.get('success') or not view_report.get('corePreserved')
                    or view_report.get('blueprintDiagram',{}).get('status') != 'Diagram layout current'):
                raise RuntimeError('Engineer-view verification did not pass; no final workbook published')
            report['seededWorkbookSha256']=hashlib.sha256(seeded.read_bytes()).hexdigest()
            report['outputSha256']=hashlib.sha256(enhanced.read_bytes()).hexdigest()
            report['engineerViewVerification']=view_report
            report['engineerViewVerification']['outputPath']=str(output)
            report['notes'].append('Engineer rules and Blueprint view are refreshed automatically after import. After manual input edits, save and run Update-FirewallViews.ps1 to a new filename.')
            staged_report=Path(tmp)/reportpath.name
            write_report(report,staged_report)
            staged_viewreport.write_text(json.dumps(view_report,indent=2),encoding='utf-8')
            # Exclusive publication never overwrites a file that appeared during Excel work.
            created=[]
            try:
                for src,dst in ((staged_report,reportpath),(staged_report.with_suffix('.txt'),reportpath.with_suffix('.txt')),(staged_viewreport,viewreportpath),(enhanced,output)):
                    with dst.open('xb') as target:
                        created.append(dst)
                        with src.open('rb') as origin:shutil.copyfileobj(origin,target)
            except Exception:
                for path in created:path.unlink(missing_ok=True)
                raise
    else:write_report(report,reportpath)
    print(json.dumps({'status':'Dry run complete' if args.dry_run else 'Import complete','counts':report['counts'],'managementOnly':report['managementOnly'],'warningCount':len(report['warnings']),'output':str(output) if not args.dry_run else None,'report':str(reportpath)},indent=2))
    if report['managementOnly']:print('WARNING: Only management-domain data found; no workload domain was created.')
    print('Review the import report and complete unconfirmed planner inputs before using firewall rules.')
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except (ValueError,KeyError,OSError,RuntimeError,zipfile.BadZipFile,ET.ParseError) as exc:
        print('Import stopped: '+str(exc),file=sys.stderr);sys.exit(1)

