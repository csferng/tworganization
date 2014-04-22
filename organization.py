# -*- coding: utf-8 -*-

import datetime
import itertools
import re
import sys
import time
import mechanize

INFO = re.compile(r'<span id="UcCoQueryResult1_dgA_Main__ctl([0-9]+)_([A-Za-z_]*)"[^>]*>(.*)</span>')
TAG = re.compile(r'<[^>]*>')
FIELDS = ['dgLblA_SeqNo_Main', u'lblCO_NAME', u'lblCO_CLASS', u'lblCO_TYPE', u'lblDIRECTOR_TITLE', u'lblDIRECTOR_NAME', u'lblESTABLISH_DTE', u'lblE_ASSOCIATION_TEL', u'lblE_AGREE_ESTABLISH_DOC', u'lblE_ASSOCIATION_ADDR']
FIELD_NAME = [u'編號', u'名稱', u'分類', u'子分類', u'會長頭銜', u'會長姓名', u'成立日期', u'電話', u'核准立案字號', u'會址']
def extract(html):
	html = html.decode('utf-8')
	items = INFO.findall(html)
	for (k,g) in itertools.groupby(items, key=lambda x:x[0]):
		d = { x[1]:TAG.sub('',x[2]) for x in g }
		if 'lblCO_NAME' in d:
			yield '\t'.join( d.get(f,'') for f in FIELDS ).encode('utf-8')

def find_control_by_suffix(br, sfx):
	return [ c.name for c in br.form.controls if hasattr(c, 'name') and c.name.endswith(sfx) ][0]

def find_all_control_by_suffix(br, sfx):
	return [ c.name for c in br.form.controls if hasattr(c, 'name') and c.name.endswith(sfx) ]

def disable_controls(br, controls):
	for sfx in controls:
		ctrls = find_all_control_by_suffix(br, sfx)
		if len(ctrls) == 1:
			br.find_control(name=ctrls[0]).disabled = True
		else:
			for (i,ctrl) in enumerate(ctrls):
				br.find_control(name=ctrl, nr=i).disabled = True

def download(org_type):
	br = mechanize.Browser()
	br.set_handle_robots(False)
	br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
	print >> sys.stderr, 'Reading http://cois.moi.gov.tw/moiweb/web/frmHome.aspx'
	br.open('http://cois.moi.gov.tw/moiweb/web/frmHome.aspx')

	br.select_form(nr=0)
	br.set_all_readonly(False)
	disable_controls(br, ['btnRead', 'btnOne', 'btnPre', 'btnNext', 'btnLast', 'ibtnANN_MORE', 'ibtn_LOGIN', 'ibtn_FORGET'])

	time.sleep(10)
	print >> sys.stderr, 'Goto search'
	br.submit(name="UcCoQuery2:gbtnSearch", coord=(24,7))

	br.select_form(nr=0)
	disable_controls(br, ['btnOne', 'btnPre', 'btnNext', 'btnLast', 'ibtn_LOGIN', 'ibtn_FORGET'])
	class_control = find_control_by_suffix(br, 'drpQ_CO_CLASS')
	br.find_control(name=class_control).set_value_by_label([org_type.encode('utf-8')])

	time.sleep(10)
	print >> sys.stderr, 'Query', org_type.encode('utf-8')
	br.submit(name="UcCoQueryResult1:btnQuery")

	br.select_form(nr=0)
	br.set_all_readonly(False)
	disable_controls(br, ['btnOne', 'btnPre', 'btnNext', 'btnLast', 'btnPrint', 'ibtn_LOGIN', 'ibtn_FORGET'])
	page_size_ctrl = find_control_by_suffix(br, 'txtPageSize')
	br['__EVENTTARGET'] = page_size_ctrl
	br[page_size_ctrl] = '999'

	time.sleep(10)
	print >> sys.stderr, 'Set 999 items per page, read page 1'
	resp = br.submit()
	yield resp.read()

	br.select_form(nr=0)
	page_cnt_ctrl = find_control_by_suffix(br, 'drpPageCount')
	all_pages =  [ itm.attrs['label'] for itm in br.find_control(page_cnt_ctrl).get_items() ]
	cur_page = br[page_cnt_ctrl][0]

	for pg in all_pages:
		if pg == cur_page: continue
		br.select_form(nr=0)
		br.set_all_readonly(False)
		disable_controls(br, ['btnOne', 'btnPre', 'btnNext', 'btnLast', 'btnPrint', 'ibtn_LOGIN', 'ibtn_FORGET'])
		br['__EVENTTARGET'] = page_cnt_ctrl
		br[page_cnt_ctrl] = [pg]

		time.sleep(10)
		print >> sys.stderr, 'Read page', pg
		resp = br.submit()
		yield resp.read()

if __name__ == "__main__":
	today = datetime.date.today().strftime("%Y%m%d")
	with open('output/'+today+'-social.tsv', 'w') as f:
		print >> f, '\t'.join(FIELD_NAME).encode('utf-8')
		for html in download(u"社會團體"):
			for item in extract(html):
				print >> f, item
	with open('output/'+today+'-professional.tsv', 'w') as f:
		print >> f, '\t'.join(FIELD_NAME).encode('utf-8')
		for html in download(u"職業團體"):
			for item in extract(html):
				print >> f, item
