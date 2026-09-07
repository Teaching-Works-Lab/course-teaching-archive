from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import re, copy, zipfile, shutil, os, subprocess, json

HERE=Path(__file__).resolve().parent
MD=HERE/'智能工厂集成技术_教案_16次课_最终版.md'
MANIFEST=json.loads((HERE/'manifest.json').read_text(encoding='utf-8'))
# Template location: when run from delivered source bundle, place school template next to script or set TEMPLATE env.
TEMPLATE=Path(os.environ.get('LESSON_PLAN_TEMPLATE','/mnt/data/机械测试技术教案.docx'))
OUT=Path(os.environ.get('LESSON_PLAN_OUTPUT','/mnt/data/智能工厂集成技术_教案_16次课_最终版.docx'))

COURSE={'name':'智能工厂集成技术','english':'Smart Factory Integration Technology','code':'25JD32406','category':'专业类','nature':'选修课','language':'中文','semester':'第7学期','credits':'2','major':'智能制造工程','college':'机械与动力工程学院','textbook':'王军利, 白海清.《智能集成制造系统》[M]. 北京：机械工业出版社，2024. ISBN 9787111768845','assessment':'最终成绩100%=平时成绩20%＋课程作业20%＋期末考核60%；期末采用智能工厂集成综合方案报告并结合方案答辩或个人核验。'}
HEADS=['课次信息','教学目标','课程思政','专创融合','教学重难点','教学方法与用具','教学设计','课前任务','互动导入','传授新知与案例训练','过关检测','课堂小结','作业布置','考勤','课后教学反思','本章节参考文献']

def parse_md():
    text=MD.read_text(encoding='utf-8'); chunks=re.split(r'(?=^## 第\d{2}次课 )',text,flags=re.M)[1:]; out=[]
    for ch in chunks:
        h=re.match(r'^## 第(\d{2})次课 (.+)$',ch,re.M); fields={}
        for i,hd in enumerate(HEADS):
            m=re.search(rf'^### {re.escape(hd)}\s*$',ch,re.M)
            if not m: continue
            start=m.end(); ends=[]
            for hd2 in HEADS[i+1:]:
                m2=re.search(rf'^### {re.escape(hd2)}\s*$',ch[start:],re.M)
                if m2: ends.append(start+m2.start())
            fields[hd]=ch[start:min(ends) if ends else len(ch)].strip()
        im=re.search(r'!\[知识脉络图\]\((images/lesson-\d{2}-knowledge-map\.png)\)',ch); fields['image']=im.group(1) if im else None
        out.append({'no':int(h.group(1)),'topic':h.group(2).strip(),'fields':fields})
    assert len(out)==16
    return out

def clean_md(s):
    s=re.sub(r'!\[[^\]]*\]\([^\)]*\)','',s); s=s.replace('**','').replace('`',''); s=re.sub(r'^- ','',s,flags=re.M); s=re.sub(r'\n{3,}','\n\n',s); return s.strip()
def info(s):
    d={}
    for line in s.splitlines():
        m=re.match(r'- \*\*(.+?)\*\*：(.+)',line.strip())
        if m:d[m.group(1)]=m.group(2).strip()
    return d

def remove_table(t): t._element.getparent().remove(t._element)
def set_text(cell,text,size=9.2,bold_prefixes=None,align=None):
    cell.text=''; p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.05
    if align is not None:p.alignment=align
    for i,line in enumerate(text.split('\n')):
        if i:p.add_run().add_break()
        r=p.add_run(line); r.font.name='宋体'; r._element.get_or_add_rPr(); r._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体'); r.font.size=Pt(size)
        if bold_prefixes and any(line.strip().startswith(x) for x in bold_prefixes):r.bold=True

def margins(cell,top=55,start=70,bottom=55,end=70):
    pr=cell._tc.get_or_add_tcPr(); mar=pr.first_child_found_in('w:tcMar')
    if mar is None: mar=OxmlElement('w:tcMar'); pr.append(mar)
    for k,v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        x=mar.find(qn('w:'+k))
        if x is None:x=OxmlElement('w:'+k); mar.append(x)
        x.set(qn('w:w'),str(v)); x.set(qn('w:type'),'dxa')

def repeat_header(row):
    pr=row._tr.get_or_add_trPr(); x=OxmlElement('w:tblHeader'); x.set(qn('w:val'),'true'); pr.append(x)

def clean_unused_images(path):
    relns={'pr':'http://schemas.openxmlformats.org/package/2006/relationships'}; docns={'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    with zipfile.ZipFile(path,'r') as zin: members={n:zin.read(n) for n in zin.namelist()}
    docxml=etree.fromstring(members['word/document.xml']); used=set(docxml.xpath('//@r:embed | //@r:link | //@r:id',namespaces=docns))
    relxml=etree.fromstring(members['word/_rels/document.xml.rels']); removed=[]
    for r in list(relxml.xpath('//pr:Relationship',namespaces=relns)):
        if r.get('Type','').endswith('/image') and r.get('Id') not in used: removed.append(r.get('Target')); r.getparent().remove(r)
    members['word/_rels/document.xml.rels']=etree.tostring(relxml,xml_declaration=True,encoding='UTF-8',standalone='yes')
    for t in removed:
        if t.startswith('media/'): members.pop('word/'+t,None)
    tmp=path.with_suffix('.cleaning.docx')
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in members.items():z.writestr(n,b)
    tmp.replace(path)

def main():
    parsed=parse_md(); doc=Document(TEMPLATE); proto=copy.deepcopy(doc.tables[2]._element)
    for t in list(doc.tables[2:]):remove_table(t)
    # remove all old lesson body content after template notes; retain cover/basic-info structure
    body=doc._element.body
    for ch in list(body)[26:]:
        if ch.tag.endswith('}sectPr'): continue
        body.remove(ch)
    cover=doc.tables[0]
    vals=['机械与动力工程学院','智能制造工程','智能工厂集成技术','（按实际填写）','（按实际填写）','第7学期','智能制造工程']
    for i,v in enumerate(vals):set_text(cover.cell(i,1),v,12.5,align=WD_ALIGN_PARAGRAPH.CENTER)
    ci=doc.tables[1]
    set_text(ci.cell(0,1),COURSE['name'],9.5); set_text(ci.cell(1,1),COURSE['english'],9.5); set_text(ci.cell(2,1),COURSE['category'],9.5); set_text(ci.cell(2,3),COURSE['nature'],9.5); set_text(ci.cell(2,5),COURSE['language'],9.5); set_text(ci.cell(3,1),COURSE['semester'],9.5); set_text(ci.cell(3,5),COURSE['credits'],9.5)
    for c,v in zip(range(1,6),['32','32','0','0','0']):set_text(ci.cell(5,c),v,9.5,align=WD_ALIGN_PARAGRAPH.CENTER)
    set_text(ci.cell(6,1),COURSE['major'],9.5); set_text(ci.cell(7,1),COURSE['textbook'],8.8); set_text(ci.cell(8,1),COURSE['college'],9.5); set_text(ci.cell(9,1),'现行课程教学大纲未列明',9.2); set_text(ci.cell(10,1),'现行课程教学大纲未列明',9.2)
    set_text(ci.cell(11,1),'考试课（ ）；考查课（√）',9.5); set_text(ci.cell(12,1),'闭卷（ ）；开卷（ ）；笔试（ ）；机试（ ）；口试（√，答辩/个人核验）；其它（√，智能工厂集成综合方案报告）',8.6); set_text(ci.cell(13,1),'平时成绩（√）；课程作业（√）；综合方案报告（√）；答辩/个人核验（√）',8.8); set_text(ci.cell(14,1),COURSE['assessment'],9.0)
    intro='课程基本定位：本课程是智能制造工程专业的专业选修课，面向设备、单元、产线、车间与企业信息系统之间的集成问题，以智能工厂总体架构为主线。\n核心学习结果：学生能够识别智能工厂各层对象与接口，理解典型工业通信协议、制造单元协同、MES与企业系统、工业互联网平台和数字孪生，并形成包含架构、接口、数据、测试、安全与运维的综合方案。\n教学组织：培养方案规定32学时全部为理论学时，不另行虚构实验、上机或项目式学时；通过系统框图、接口表、数据流、时序图、测试矩阵和案例评审训练集成思维。'
    set_text(ci.cell(15,1),intro,8.8,bold_prefixes=['课程基本定位','核心学习结果','教学组织'])
    for p in parsed:
        sep=doc.add_paragraph(); sep.add_run().add_break(WD_BREAK.PAGE); new=copy.deepcopy(proto); sep._p.addnext(new); t=doc.tables[-1]; f=p['fields']; li=info(f['课次信息'])
        set_text(t.cell(0,1),li.get('章节/单元',''),9.2); set_text(t.cell(0,3),p['topic'],9.2); set_text(t.cell(1,1),'按实际课表填写',9.2); set_text(t.cell(1,3),'2课时（100 min）',9.2)
        c=t.cell(2,1); c.text=''; pp=c.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.add_run().add_picture(str(HERE/f['image']),width=Cm(13.2))
        set_text(t.cell(3,1),clean_md(f['教学目标']),9.0,bold_prefixes=['知识目标','能力目标','价值目标']); set_text(t.cell(4,1),clean_md(f['课程思政']),8.9); set_text(t.cell(5,1),clean_md(f['专创融合']),8.9); set_text(t.cell(6,1),clean_md(f['教学重难点']),8.9,bold_prefixes=['教学重点','教学难点'])
        method=clean_md(f['教学方法与用具']); mm=re.search(r'教学方法：(.*?)(?:\n\n|\n)教学用具：(.*)',method,re.S)
        if mm:set_text(t.cell(7,1),mm.group(1).strip(),8.8);set_text(t.cell(8,1),mm.group(2).strip(),8.8)
        else:set_text(t.cell(7,1),method,8.8);set_text(t.cell(8,1),'电脑、投影仪、多媒体课件、教材。',8.8)
        set_text(t.cell(9,1),clean_md(f['教学设计']),8.8); set_text(t.cell(10,1),'教学步骤及主要教学内容',10.0); set_text(t.cell(10,4),'设计意图',10.0)
        set_text(t.cell(11,1),clean_md(f['课前任务']),8.8,bold_prefixes=['【教师】','【学生】']); set_text(t.cell(11,4),'通过预读架构/接口/案例建立先备认知，让学生带着系统边界和接口问题进入课堂。',8.5)
        combo='【互动导入 15 min】\n'+clean_md(f['互动导入'])+'\n\n【传授新知与案例训练 70 min】\n'+clean_md(f['传授新知与案例训练'])
        set_text(t.cell(12,0),'互动导入\n（15 min）\n\n传授新知与案例训练\n（70 min）',9.0); set_text(t.cell(12,1),combo,8.5,bold_prefixes=['【互动导入','【传授','课堂学习产物','学习证据','评价映射']); set_text(t.cell(12,4),'采用“场景/异常 → 架构与接口讲解 → 学生图表产出 → 同伴评审 → 证据固化”的理论课堂闭环，不虚构实验学时。',8.4)
        set_text(t.cell(13,1),clean_md(f['过关检测']),8.7); set_text(t.cell(13,4),'检测概念、边界、接口和异常判断，要求说明系统或数据依据。',8.4)
        set_text(t.cell(14,1),clean_md(f['课堂小结']),8.7); set_text(t.cell(14,4),'回到知识脉络图，将对象、接口、数据、流程、异常和证据组织为系统结构。',8.4)
        set_text(t.cell(15,1),clean_md(f['作业布置']),8.7); set_text(t.cell(15,4),'将课堂图表沉淀为课程作业阶段证据，保持版本、来源和修改依据。',8.4)
        set_text(t.cell(16,1),clean_md(f['考勤']),8.7); set_text(t.cell(16,4),'保留学校模板考勤环节，不预填实际出勤事实。',8.4)
        set_text(t.cell(17,1),clean_md(f['课后教学反思']),8.5); set_text(t.cell(17,4),'教学反思仅依据真实课堂证据课后填写。',8.3)
        set_text(t.cell(18,0),'本章节参考文献',8.2,align=WD_ALIGN_PARAGRAPH.CENTER); set_text(t.cell(18,1),clean_md(f['本章节参考文献']),8.2)
        for row in t.rows:
            for cc in row.cells: cc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; margins(cc)
        repeat_header(t.rows[0])
    cp=doc.core_properties; cp.title='智能工厂集成技术 教学设计（16次课）'; cp.subject='智能制造工程；32学时理论课'; cp.keywords='智能工厂集成技术,智能制造,MES,工业互联网,数字孪生,教案'; cp.comments='Markdown为内容主源，Word为学校格式编译产物。'
    doc.save(OUT)
    scrub=Path('/home/oai/skills/docx/scripts/privacy_scrub.py'); tmp=OUT.with_name(OUT.stem+'_scrub.docx'); subprocess.run(['python',str(scrub),str(OUT),'--out',str(tmp)],check=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True); tmp.replace(OUT); clean_unused_images(OUT)
if __name__=='__main__':main()
