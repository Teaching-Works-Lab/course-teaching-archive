from pathlib import Path
from graphviz import Digraph
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy
import json, zipfile, hashlib, shutil, re, subprocess, os

BASE=Path('/mnt/data')
TEMPLATE=BASE/'工业机器人技术_教案_16次课_最终版.docx'
COURSE_DIR=BASE/'机器视觉技术_教案_源码包'
IMG_DIR=COURSE_DIR/'images'
FULL_DIR=BASE/'机器视觉技术_教案_完整交付包'
FINAL_MD=BASE/'机器视觉技术_教案_16次课_最终版.md'
FINAL_DOCX=BASE/'机器视觉技术_教案_16次课_最终版.docx'
SRC_ZIP=BASE/'机器视觉技术_教案_16次课_源码包.zip'
FULL_ZIP=BASE/'机器视觉技术_教案_完整交付包.zip'
for d in [COURSE_DIR,FULL_DIR]:
    if d.exists(): shutil.rmtree(d)
COURSE_DIR.mkdir(); IMG_DIR.mkdir()

COURSE={
'name':'机器视觉技术','english':'Machine Vision Technology','code':'25JD32405','category':'专业类','nature':'选修','language':'中文','semester':'第7学期','credits':'2',
'hours_total':32,'hours_theory':32,'hours_exp':0,'hours_lab':0,'hours_other':0,'major':'智能制造工程','college':'机械与动力工程学院',
'prereq':'程序设计基础（C&C++）Ⅱ、工程建模与科学计算可视化基础（Python）','followup':'毕业设计（论文）',
'textbook':'何文辉、葛大伟.《机器视觉技术与应用》[M]. 北京：机械工业出版社，2024. ISBN 9787111763048',
'assessment':'最终成绩100%=平时成绩20%＋课程作业20%＋期末考核60%；期末考核采用机器视觉综合方案/项目报告（或等效作品）＋现场答辩/个人核验，不另设笔试或机考。',
'syllabus_path':'course/机器视觉技术/2025/大纲/25JD32405-机器视觉技术-课程教学大纲.md','syllabus_sha':'cd100bf9dd06b85abdc5cbc10394f930d5973934',
'note_path':'course/机器视觉技术/2025/说明/25JD32405-机器视觉技术-依据与实施说明.md','note_sha':'15a237b72262f2deb3e32158a250e41575c962b9',
'skill_repo':'Teaching-Works-Lab/lesson-plan-compiler','skill_sha':'23bc5a10e7881e8863ad993c74dabdc9faf24a4d','updated':'2026年9月'
}
REF1='[1] 何文辉、葛大伟.《机器视觉技术与应用》[M]. 北京：机械工业出版社，2024.'
REF2='[2] 李宁.《Python OpenCV从菜鸟到高手》[M]. 北京：清华大学出版社，2024.'

# milestone assignment weights inside course-work 20%
UNIT_ASSESS={
'第一单元 机器视觉系统与成像基础':'课程作业内部权重15%；目标1.1、2.1、3.1、3.3',
'第二单元 Python/OpenCV图像表示与预处理':'课程作业内部权重20%；目标1.2、2.2、3.2、3.5',
'第三单元 分割、边缘与数学形态学':'课程作业内部权重20%；目标1.3、2.2、3.2、3.4',
'第四单元 轮廓、连通域、测量与定位':'课程作业内部权重20%；目标1.4、2.3、3.3、3.4',
'第五单元 模板匹配、特征与工业识别':'课程作业内部权重10%；目标1.5、2.4、3.4、3.6',
'第六单元 深度学习视觉入门与智能制造综合案例':'课程作业内部权重15%；目标1.6、2.5、2.6、3.1—3.6',
}

def L(no,unit,topic,map_,objk,obja,objv,focus,diff,intro,blocks,checks,homework,ideology,innovation,tools,refs,milestone=False):
    return dict(no=no,unit=unit,topic=topic,map=map_,objk=objk,obja=obja,objv=objv,focus=focus,diff=diff,intro=intro,blocks=blocks,checks=checks,homework=homework,ideology=ideology,innovation=innovation,tools=tools,refs=refs,milestone=milestone)

lessons=[
L(1,'第一单元 机器视觉系统与成像基础','机器视觉系统、工业任务与“成像—处理—判定”工作链',
 {'任务定义':['检测/测量','定位/识别/引导'],'系统组成':['相机/镜头/光源','计算与I/O'],'视觉工作链':['成像','处理/特征','判定/评价'],'工程约束':['精度/速度','环境/安全']},
 ['理解机器视觉的定义、典型工业任务和系统基本组成。','理解“成像条件—图像处理—特征与测量—识别判定—性能评价”的完整工作链。'],
 ['能够把零件检测或装配检查任务拆解为对象、输入图像、输出结果和评价指标。','能够识别系统中相机、镜头、光源、计算平台与外围接口的职责。'],
 ['形成先定义检测目标和工程约束、再选择算法的系统工程意识。','理解误检、漏检和错误引导可能带来的质量与安全责任。'],
 '机器视觉任务类型、系统组成与视觉工作链。','把“识别算法”放回完整系统中理解，避免只关注后端算法而忽视成像条件和工程判据。',
 '展示同一个零件在良好照明与强反光条件下的两组图像，提问：算法不变时为什么结果可能完全不同？由此引出完整视觉链。',
 [('机器视觉任务与应用',18,'【教师】用质量检测、尺寸测量、工件定位、装配完整性和视觉引导案例区分不同任务输出。','【学生】为5个案例判断任务类型，并写出最小输出结果。'),('系统组成与职责',18,'【教师】拆解相机、镜头、光源、计算单元、I/O和执行机构的功能边界。','【学生】完成“模块—输入—输出—失效影响”表。'),('视觉工作链',18,'【教师】用成像→预处理→分割/特征→测量/识别→判定→评价串联后续课程。','【学生】为一个装配检查案例画流程图。'),('工程约束与评价',16,'【教师】说明精度、速度、鲁棒性、误检/漏检和安全约束，强调“单张成功图”不能代表工程可用。','【学生】为案例补充3个验收指标与1个失败工况。')],
 ['【选择题】机器视觉系统中最直接决定原始图像质量的环节通常包括（ ）。A. 相机/镜头/光源 B. 排序算法 C. 数据库 D. PLC程序','【判断题】只要识别算法准确率高，成像条件变化通常无需重新验证。（ ）','【简答题】一个可验收的工业视觉任务至少应明确哪些输入、输出和评价指标？'],
 ['选择一个机械零件、装配或FDM打印场景，完成“视觉任务定义卡”：对象、任务、输入、输出、精度/速度和失败风险。'],
 '结合关键缺陷漏检和错误定位案例，强调视觉结论会进入质量判定甚至设备动作链，必须以可复核证据承担工程责任。',
 '围绕“人工目检→机器视觉”的替代条件，比较一致性、节拍、可追溯和改造成本，提出一个可行的小型视觉改造方案。',
 '电脑、投影仪、多媒体课件、教材、工业视觉系统案例图、任务定义卡。',REF1+'；'+REF2),
L(2,'第一单元 机器视觉系统与成像基础','相机、镜头、光源与视场/分辨率/景深/曝光设计',
 {'相机':['传感器/像素','曝光/噪声'],'镜头':['焦距/视场','工作距离/景深'],'光源':['明场/暗场','背光/同轴等思路'],'方案约束':['分辨率','运动模糊','稳定性']},
 ['理解相机像素、曝光、噪声，镜头焦距、视场、工作距离、景深以及光源方式的工程含义。','理解视场、目标尺寸、测量精度和像素分辨率之间的关系。'],
 ['能够根据目标尺寸和最小特征估算基本成像分辨率需求。','能够针对反光、轮廓、表面缺陷等场景选择合理照明思路并说明理由。'],
 ['形成“先把图像拍稳定，再谈算法”的工程习惯。','树立参数变更留痕和多工况复核意识。'],
 '成像参数与光照方案；从检测精度反推视场/像素需求。','将参数计算与现场光照、运动和景深约束结合，避免只做理想条件下的静态计算。',
 '给出同一金属件的正面明场、侧向暗场和背光图，要求学生先判断哪一种更适合表面划痕、轮廓尺寸和孔洞检测。',
 [('相机与曝光',18,'【教师】讲解像素、分辨率、曝光、增益、噪声与运动模糊的基本关系。','【学生】判断提高曝光时间可能带来的收益与风险。'),('镜头与几何参数',20,'【教师】讲解视场、工作距离、焦距、景深和畸变的工程意义，完成最小像素需求示例。','【学生】根据零件尺寸和最小缺陷尺寸完成一组分辨率估算。'),('光源与成像对比',18,'【教师】比较明场、暗场、背光、同轴/漫射等照明思路及其适用对象。','【学生】为三种检测任务匹配照明方案并说明理由。'),('方案评审',14,'【教师】组织“精度—视场—景深—曝光—速度”约束评审。','【学生】完成一页成像方案卡，指出2个最可能失效的工况。')],
 ['【选择题】要突出零件外轮廓并弱化表面纹理，通常更适合优先考虑（ ）。A. 背光 B. 随机环境光 C. 关闭光源 D. 任意颜色屏幕','【判断题】曝光时间越长，所有运动场景图像都越清晰。（ ）','【简答题】为什么“像素很多”并不自动等于“测量精度一定高”？'],
 ['完成一个零件尺寸检测成像方案：视场、最小特征、像素需求、工作距离、照明思路与两类风险。'],
 '以错误照明导致缺陷漏检的案例说明成像条件也是质量控制的一部分，不能把全部责任推给后端算法。',
 '比较“更换算法”和“优化光学/照明”两种改造路径的成本与收益，培养跨光机电软协同的方案意识。',
 '电脑、投影仪、多媒体课件、教材、相机/镜头参数样本、照明对比图、计算器。',REF1,milestone=True),
L(3,'第二单元 Python/OpenCV图像表示与预处理','数字图像、NumPy像素/ROI与BGR-RGB-HSV-灰度表示',
 {'数字图像':['像素/坐标','数据类型/范围'],'NumPy表示':['shape','索引/切片/ROI'],'颜色空间':['BGR/RGB','HSV','灰度'],'可复现输入':['路径/读取','原图保存']},
 ['理解数字图像的像素、坐标、通道、shape和数据类型。','掌握BGR/RGB、HSV和灰度表示及其基本用途。'],
 ['能够使用NumPy/OpenCV读取图像、访问像素、提取ROI并转换颜色空间。','能够识别通道顺序、坐标和数据类型错误造成的异常结果。'],
 ['形成原始图像不覆盖、参数和代码版本可追溯的规范习惯。'],
 '图像矩阵表示、ROI、颜色空间转换与数据类型。','学生容易混淆行列坐标、x/y坐标和BGR/RGB通道顺序，需要通过最小代码和结果图对应验证。',
 '展示一张“颜色完全错误”的OpenCV→Matplotlib显示结果，让学生判断是相机坏了还是通道顺序出了问题。',
 [('图像矩阵与坐标',18,'【教师】用小矩阵说明高×宽×通道、uint8范围、行列索引和图像坐标。','【学生】手算3个像素位置并预测切片结果。'),('NumPy与ROI',18,'【教师】演示读取、shape、切片、复制与ROI，说明视图/复制和越界风险。','【学生】写/补全最小代码提取指定ROI并检查尺寸。'),('颜色空间',20,'【教师】比较BGR/RGB、HSV和灰度的表示与适用场景。','【学生】对同一图像生成三种表示并描述信息变化。'),('规范输入与证据',14,'【教师】示范保留原图、输出目录和参数日志。','【学生】建立“原图—中间图—结果图”的文件命名规则。')],
 ['【选择题】OpenCV默认读取彩色图像的通道顺序通常是（ ）。A. RGB B. BGR C. HSV D. CMYK','【判断题】图像数组的第一维通常直接表示x坐标。（ ）','【简答题】HSV相比RGB/BGR在颜色阈值分割中有什么潜在优势？'],
 ['用一张课程图像完成读取、ROI、BGR→RGB、BGR→HSV和灰度转换，保留原图与中间结果。'],
 '强调原始图像、通道和数据类型属于实验事实，禁止为了“结果好看”覆盖原始数据或省略处理步骤。',
 '将ROI和颜色空间转换封装为可复用预处理模块，讨论模块接口对后续检测流程复用的价值。',
 '电脑、投影仪、多媒体课件、教材、Python、NumPy、OpenCV、Jupyter/VS Code、示例图像。',REF1+'；'+REF2),
L(4,'第二单元 Python/OpenCV图像表示与预处理','直方图、亮度/对比度与成像质量对比',
 {'灰度统计':['直方图','均值/范围'],'亮度':['偏移','饱和/截断'],'对比度':['动态范围','增强与失真'],'质量比较':['同场景对照','参数记录']},
 ['理解直方图反映的灰度分布及亮度、对比度的基本含义。','理解像素线性变换、截断/饱和对图像信息的影响。'],
 ['能够用OpenCV/NumPy计算并绘制直方图，比较不同曝光/亮度/对比度条件。','能够用受控变量方式评价预处理是否改善了后续可分性。'],
 ['形成不凭“视觉更漂亮”评价算法、而以任务效果和参数证据判断的习惯。'],
 '直方图、亮度/对比度调整与质量评价。','区分主观观感改善与机器视觉任务可分性改善，避免把图像增强等同于检测性能提升。',
 '展示两张“肉眼看起来更亮”的图像，其中一张细节被过曝截断，提问哪张更适合阈值检测并说明证据。',
 [('直方图与统计',18,'【教师】解释直方图、灰度范围、均值和峰形，演示同场景不同曝光的统计变化。','【学生】根据3条直方图判断欠曝、正常和过曝可能性。'),('亮度调整',16,'【教师】演示加减偏移和uint8截断问题。','【学生】预测不同偏移对暗区/亮区信息的影响。'),('对比度调整',18,'【教师】说明线性缩放和动态范围，强调增强可能放大噪声。','【学生】对一张低对比图做2组参数对比。'),('任务导向评价',18,'【教师】把处理结果送入简单阈值/边缘观察，说明增强必须服务后续任务。','【学生】记录“参数—直方图—后续效果”三列证据。')],
 ['【选择题】图像大量像素堆积在灰度高端并发生截断，最可能说明（ ）。A. 过曝 B. 欠曝 C. 坐标系错误 D. 镜头焦距为0','【判断题】肉眼观感更强的对比度一定会提高机器视觉检测准确性。（ ）','【简答题】为什么比较亮度/对比度算法时必须固定原始图像和其他条件？'],
 ['完成“原图＋两组亮度/对比度参数”的直方图和结果对比，写出哪组更适合后续分割及依据。'],
 '通过过曝和噪声放大案例强调参数调整必须诚实报告，不选择性隐藏失败图像。',
 '围绕“自动曝光/自适应参数”的工程需求，提出一种基于统计量的简单参数调节思路。',
 '电脑、投影仪、多媒体课件、教材、Python、NumPy、OpenCV、Matplotlib。',REF1+'；'+REF2),
L(5,'第二单元 Python/OpenCV图像表示与预处理','缩放/旋转/透视变换与均值/高斯/中值滤波',
 {'几何变换':['缩放/旋转','透视变换'],'插值':['最近邻','线性等概念'],'噪声处理':['均值/高斯','中值'],'对比评价':['细节/噪声','参数/失败样例']},
 ['掌握缩放、旋转、透视变换的用途和基本参数。','理解均值、高斯、中值滤波的核心差异及典型噪声适用性。'],
 ['能够用OpenCV完成常见几何变换和滤波，并保存对比结果。','能够针对椒盐噪声、随机噪声和细节保留要求选择基本滤波方案。'],
 ['形成受控变量和失败样例记录意识，避免“参数试到好看为止”。'],
 '几何变换与三类基本滤波；参数对图像信息的影响。','既要降低噪声又要保留边缘/几何信息，学生需理解滤波尺度与后续任务之间的权衡。',
 '给出“中值滤波去椒盐噪声”和“高斯滤波去椒盐噪声”两组结果，让学生先观察细节和残留噪声再解释原因。',
 [('几何变换',18,'【教师】演示缩放、旋转、透视变换的矩阵/映射思想和OpenCV接口。','【学生】为倾斜拍摄的矩形标签选择校正方案。'),('插值与信息变化',12,'【教师】说明缩放并非凭空增加真实细节，比较最近邻/线性插值现象。','【学生】观察放大后边缘差异并写出限制。'),('三类滤波',22,'【教师】比较均值、高斯和中值滤波及核大小影响。','【学生】对两类噪声分别选择滤波器并说明依据。'),('参数AB比较',18,'【教师】要求固定原图，比较2—3组核大小及边缘保留。','【学生】提交参数表和失败样例，不能只保留最好结果。')],
 ['【选择题】对椒盐噪声通常优先尝试（ ）。A. 中值滤波 B. 任意旋转 C. 直方图统计 D. 透视变换','【判断题】图像放大后生成的更多像素等同于获得了更多真实空间细节。（ ）','【简答题】滤波核过大为什么可能降低后续边缘定位精度？'],
 ['选择一张噪声图，比较均值/高斯/中值滤波至少3种配置，并记录“噪声抑制—边缘保留—后续任务”结论。'],
 '以参数过度平滑造成关键缺陷消失的案例强调图像处理不能只追求视觉平滑，必须对检测后果负责。',
 '把几何校正和滤波组合为“可配置预处理流水线”，思考如何支持不同产品快速换型。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、NumPy、Matplotlib、噪声样例。',REF1+'；'+REF2,milestone=True),
L(6,'第三单元 分割、边缘与数学形态学','固定阈值、自适应阈值与Otsu分割',
 {'阈值分割':['前景/背景','固定阈值'],'自适应':['局部阈值','窗口/常数'],'Otsu':['类间方差思想','适用分布'],'失效分析':['光照不均','纹理/噪声']},
 ['掌握固定阈值、自适应阈值和Otsu阈值的基本思想与适用条件。','理解二值结果中前景/背景定义和阈值方向。'],
 ['能够对同一组图像比较三类阈值方法，并记录阈值/窗口等参数。','能够从光照不均、直方图分布和噪声解释分割失败。'],
 ['形成跨多图像验证参数、不以单张成功截图作为结论的实证意识。'],
 '三类阈值分割方法及光照/噪声条件下的失败分析。','参数并非越“自动”越可靠，需要从图像分布和局部光照解释方法边界。',
 '展示一张中心明亮、边缘偏暗的零件图：固定阈值在中心成功、边缘失败；要求学生先判断“换一个全局阈值能否根治”。',
 [('固定阈值',16,'【教师】讲解二值化、阈值方向和阈值选择，结合直方图解释。','【学生】根据灰度分布预测两个阈值的分割差异。'),('自适应阈值',18,'【教师】说明局部窗口与常数项的作用，演示光照不均场景。','【学生】比较两组窗口大小并观察局部噪声。'),('Otsu方法',18,'【教师】以类间可分性直观解释Otsu自动阈值，说明其对分布的依赖。','【学生】对双峰/非双峰样例判断是否适合。'),('AB验证与失败归因',18,'【教师】组织固定/自适应/Otsu多图像对照。','【学生】完成“方法—参数—成功条件—失败样例”表。')],
 ['【选择题】当图像存在明显不均匀光照时，相比单一固定阈值更值得尝试（ ）。A. 自适应阈值 B. 图像旋转 C. 只看RGB蓝通道 D. 不做验证','【判断题】Otsu产生一个自动阈值，因此在所有分布上都比人工阈值更可靠。（ ）','【简答题】阈值分割实验为什么必须明确前景是白还是黑？'],
 ['使用不少于5张不同光照图像比较固定/自适应/Otsu，统计失败图像并说明根因。'],
 '通过“只挑成功图片”会夸大方案能力的案例，强调评价必须包含失败样例和真实参数。',
 '把阈值方法选择设计成可配置策略，探索根据光照统计自动切换分割方案的思路。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、直方图与多光照图像。',REF1+'；'+REF2),
L(7,'第三单元 分割、边缘与数学形态学','Canny边缘检测、梯度阈值与边缘质量评价',
 {'边缘基础':['强度变化','梯度概念'],'Canny流程':['平滑/梯度','非极大抑制','双阈值/连接'],'参数':['低/高阈值','滤波尺度'],'质量':['断裂/伪边缘','后续轮廓']},
 ['理解边缘与灰度梯度的关系，掌握Canny流程的基本组成。','理解双阈值与边缘连接对噪声和弱边缘的影响。'],
 ['能够使用OpenCV Canny并系统比较阈值参数。','能够用断裂、伪边缘和后续轮廓完整性评价边缘结果。'],
 ['形成参数选择可解释、结果可复核的算法使用习惯。'],
 'Canny基本流程、双阈值参数和边缘质量评价。','学生容易只调阈值看“边缘多不多”，难点是把边缘质量与后续轮廓/测量任务关联。',
 '给出两张Canny结果：一张边缘很多但噪声严重，一张边缘较少但目标轮廓完整，要求学生先定义“哪个更好”的任务依据。',
 [('边缘与梯度',14,'【教师】用一维灰度剖面说明强度变化和梯度，建立边缘概念。','【学生】在剖面图上标出可能边缘。'),('Canny流程',22,'【教师】讲解平滑、梯度、非极大抑制、双阈值与连接的逻辑。','【学生】用流程卡排序Canny步骤并解释每步作用。'),('阈值参数实验',18,'【教师】固定图像比较3组低/高阈值。','【学生】记录伪边缘、断裂和目标边界完整度。'),('任务评价',16,'【教师】把边缘结果用于轮廓/直线检测示例，说明边缘是中间证据。','【学生】选择一组参数并用后续任务结果证明选择。')],
 ['【选择题】Canny中的非极大抑制主要用于（ ）。A. 细化边缘 B. 转换颜色空间 C. 计算直方图 D. 增大图像尺寸','【判断题】Canny阈值越低，边缘越多，因此检测效果一定越好。（ ）','【简答题】为什么评价边缘结果时应同时观察后续轮廓或几何检测？'],
 ['对同一图像比较至少3组Canny阈值，附原图、边缘图和后续轮廓/直线结果。'],
 '强调参数结果需经后续任务验证，避免为了截图“好看”而忽略伪边缘带来的误判风险。',
 '讨论将Canny参数与不同材质/光照产品配置绑定，形成可切换工艺参数表。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、边缘/轮廓样例。',REF1+'；'+REF2),
L(8,'第三单元 分割、边缘与数学形态学','腐蚀/膨胀、开闭运算与粘连/断裂/孔洞修复',
 {'基础操作':['腐蚀','膨胀'],'组合操作':['开运算','闭运算'],'结构元素':['形状','尺寸'],'问题修复':['噪点','粘连','断裂/孔洞']},
 ['掌握腐蚀、膨胀、开运算和闭运算的基本作用。','理解结构元素形状/尺寸与目标几何之间的关系。'],
 ['能够根据噪点、粘连、断裂、孔洞问题选择形态学组合。','能够比较不同结构元素并分析过处理导致的尺寸/拓扑改变。'],
 ['形成“结构修复不能破坏真实缺陷”的质量意识。'],
 '形态学操作、结构元素与二值结构修复。','难点是选择操作顺序和结构元素尺度，使噪声得到处理而目标几何不被过度改变。',
 '展示一张含小噪点、细小断裂和相邻粘连的二值图，要求学生指出一个操作不可能同时解决全部问题。',
 [('腐蚀与膨胀',16,'【教师】用集合/邻域直观说明腐蚀和膨胀对前景边界的作用。','【学生】预测同一小孔/细线经过两操作后的变化。'),('开闭运算',18,'【教师】比较开运算去小前景噪点、闭运算填小孔/连接断裂的典型效果。','【学生】为四类问题选择开/闭及顺序。'),('结构元素',18,'【教师】演示形状、大小对细长目标、圆孔和间隙的影响。','【学生】比较两种核并解释为何可能破坏尺寸。'),('组合验证',18,'【教师】提供多工况二值图进行流程组合。','【学生】提交“原始二值→形态学→轮廓”的前后证据和失败案例。')],
 ['【选择题】去除孤立的小白色噪点通常优先考虑（ ）。A. 开运算 B. 闭运算 C. 透视变换 D. 颜色空间转换','【判断题】结构元素越大，形态学结果通常越可靠。（ ）','【简答题】为什么测量任务中形态学处理必须评估几何尺寸偏差？'],
 ['围绕“噪点、粘连、断裂、孔洞”至少选择两类问题设计形态学处理，并报告结构元素与失败样例。'],
 '通过过度形态学把真实小缺陷抹掉的案例强调处理算法不能为了结果整洁牺牲质量事实。',
 '将分割+形态学流程封装成“产品参数配方”，比较不同产品换型时参数管理方法。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、二值图和结构元素示例。',REF1+'；'+REF2,milestone=True),
L(9,'第四单元 轮廓、连通域、测量与定位','轮廓、层级、连通域与面积/周长/质心统计',
 {'轮廓':['边界点','层级/孔洞'],'连通域':['标签','数量/面积'],'几何统计':['面积/周长','质心/矩'],'工业计数':['筛选','数量/异常']},
 ['掌握轮廓、层级与连通域的基本概念。','理解面积、周长、质心和矩等几何统计量。'],
 ['能够使用OpenCV提取轮廓/连通域并完成目标计数与筛选。','能够通过面积、位置等条件排除明显噪声并解释筛选依据。'],
 ['形成统计规则和阈值来源可解释的质量判定意识。'],
 '轮廓/连通域提取及面积、周长、质心统计。','轮廓层级、孔洞和噪声目标会影响计数，需要从二值结构和工业判据共同解释。',
 '给出一个带孔零件和一个实心零件的二值图，使用不同轮廓检索模式得到不同数量，要求学生解释“数量为什么变了”。',
 [('轮廓与检索模式',18,'【教师】讲解轮廓边界、层级和孔洞关系，比较常见检索模式。','【学生】根据层级树判断外轮廓与孔轮廓。'),('连通域',16,'【教师】讲解连通域标签、面积、包围盒与质心。','【学生】对颗粒/零件图完成数量统计思路。'),('几何统计',18,'【教师】说明面积、周长、矩和质心的计算用途。','【学生】根据特征表筛选过小噪声和目标。'),('计数案例',18,'【教师】用工件计数案例串联分割→形态学→连通域/轮廓→筛选。','【学生】形成参数和数量结果表，并保留误计案例。')],
 ['【选择题】需要同时获取每个独立前景区域的标签和面积时，常用思路是（ ）。A. 连通域分析 B. 图像旋转 C. 颜色显示 D. 文件压缩','【判断题】轮廓数量等于真实零件数量，不需要考虑孔洞或噪声。（ ）','【简答题】面积阈值用于剔除噪声时，阈值应如何获得工程依据？'],
 ['使用一组零件图完成轮廓/连通域计数，报告面积筛选阈值、误计样例与修改依据。'],
 '强调计数和统计阈值必须有来源，避免为了达到预期数量随意调参。',
 '把轮廓/连通域统计做成生产计数模块，思考如何输出可对接MES/PLC的结构化结果。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、计数与孔洞样例。',REF1+'；'+REF2),
L(10,'第四单元 轮廓、连通域、测量与定位','外接几何、直线/圆检测与尺寸/位置/角度提取',
 {'外接几何':['矩形/最小矩形','圆/椭圆'],'线圆检测':['直线','圆'],'位姿特征':['中心位置','方向角'],'尺寸信息':['像素长度','几何特征']},
 ['掌握外接矩形、最小外接矩形、圆/椭圆拟合及直线/圆检测的基本用途。','理解像素几何特征与零件位置、角度、尺寸之间的关系。'],
 ['能够从轮廓提取中心、宽高、角度等特征。','能够根据对象几何选择轮廓拟合或直线/圆检测方法。'],
 ['形成“几何结果必须对应真实基准和工艺判据”的测量意识。'],
 '外接几何、直线/圆检测、位置/角度与像素尺寸。','学生需区分“算法输出几何参数”和“工程尺寸/姿态”之间仍需要基准、尺度和坐标定义。',
 '展示倾斜矩形工件：普通外接矩形给出轴对齐宽高，而最小外接矩形给出方向角，要求学生判断哪组更适合姿态估计。',
 [('外接几何',18,'【教师】比较boundingRect、minAreaRect、minEnclosingCircle、fitEllipse的输出与条件。','【学生】为矩形、圆孔、椭圆件分别选择特征。'),('直线与圆',16,'【教师】讲解边缘到直线/圆检测的基本思路和参数敏感性。','【学生】判断复杂背景中为何先做好边缘/ROI更重要。'),('位姿特征',18,'【教师】用最小外接矩形中心和角度说明二维定位。','【学生】计算/读取中心与角度，并说明坐标参考。'),('像素尺寸与判定',18,'【教师】串联像素长度、尺度换算的下一步需求。','【学生】形成“像素量—工程量—需要的标定信息”表。')],
 ['【选择题】需要获得旋转矩形工件的方向角时，更适合采用（ ）。A. 最小外接矩形 B. 只看图像宽度 C. 直方图均值 D. 文件名','【判断题】检测到100像素长度即可直接得出100毫米实际尺寸。（ ）','【简答题】直线/圆检测前为什么常需要ROI、滤波或边缘预处理？'],
 ['选择一个矩形/孔类零件，提取中心、尺寸和方向角，说明每个输出的坐标/尺度含义。'],
 '通过错误坐标或尺度导致机器人错误抓取/质量误判的案例强调几何量必须有基准。',
 '将二维位置/角度输出设计成可供机器人或工控系统调用的结果接口，讨论单位和坐标契约。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、几何零件图、坐标示意。',REF1+'；'+REF2),
L(11,'第四单元 轮廓、连通域、测量与定位','尺度换算、标定概念、容差判定与重复性/误差分析',
 {'尺度基准':['像素→工程量','参考尺寸/标定概念'],'测量流程':['特征提取','尺寸/位置/角度'],'误差':['重复性','偏差/分辨率'],'判定':['容差','复测/不确定性']},
 ['理解尺度换算、参考基准和标定概念。','理解重复性、偏差、分辨率与容差判定的基本关系。'],
 ['能够依据已知尺度基准把像素量转换为工程量，并完成简单合格判定。','能够通过重复图像测量计算均值/极差或标准差等基本统计并分析误差来源。'],
 ['树立测量结果必须有尺度基准、重复性和误差说明的计量意识。'],
 '像素-工程量换算、容差判定、重复测量和误差分析。','难点是避免把单次算法输出当作真实值，需要从标定、视角、畸变、分割边界和重复性分析误差。',
 '给出同一标准件连续5次测量结果，均在公差附近波动；要求学生判断是否仅凭一次“合格”就能宣布方案可靠。',
 [('尺度换算与标定概念',18,'【教师】讲解已知基准尺寸、像素/毫米换算和更一般相机标定的作用边界。','【学生】用一组参考尺寸完成像素→毫米换算。'),('测量链与误差源',18,'【教师】梳理成像、分割、轮廓、尺度、透视/畸变等误差来源。','【学生】建立误差鱼骨/清单。'),('重复性分析',18,'【教师】演示多次测量均值、极差/标准差和异常值检查。','【学生】计算5次结果并判断稳定性。'),('容差与复核',16,'【教师】说明公差边界、灰区复测和安全判定。','【学生】设计“合格/不合格/需复核”三状态判定规则。')],
 ['【选择题】把像素长度换算成毫米，至少需要（ ）。A. 有物理意义的尺度基准 B. 只需图像文件大小 C. 只需颜色空间 D. 只需阈值','【判断题】一次测量落在公差内即可证明视觉测量系统具有良好重复性。（ ）','【简答题】列举至少三类会影响视觉尺寸测量的误差来源。'],
 ['用同一对象至少5次图像/测量结果完成尺度换算、重复性统计和公差判定，并写出主要误差来源。'],
 '以尺寸临界判定和误差责任为例强调不能隐去不确定性，对临界结果应保留复测/人工确认机制。',
 '将“合格/不合格/需复核”三状态输出用于质量系统，探索减少误判成本的产品化规则。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、测量数据表、标准尺寸样例。',REF1+'；'+REF2,milestone=True),
L(12,'第五单元 模板匹配、特征与工业识别','模板匹配、相似度图与尺度/旋转/遮挡鲁棒性',
 {'模板匹配':['模板/搜索图','相似度'],'定位':['峰值','阈值/多目标'],'敏感性':['尺度','旋转','遮挡','光照'],'验证':['误匹配','受控变量']},
 ['掌握模板匹配的基本思想和相似度响应图概念。','理解模板方法对尺度、旋转、遮挡和光照变化的敏感性。'],
 ['能够使用OpenCV完成基本模板定位并设置合理阈值。','能够设计受控变量实验分析误匹配和漏匹配。'],
 ['形成阈值必须有样本证据、识别结果必须报告误匹配的实证意识。'],
 '模板匹配流程、阈值与尺度/旋转/遮挡影响。','模板在理想条件下简单有效，但鲁棒性边界明显；学生需从受控变量实验而非单张图判断适用性。',
 '先用正向模板在原图准确定位，再把目标旋转15°，观察匹配值骤降；提问问题发生在“代码”还是“方法假设”。',
 [('模板匹配流程',18,'【教师】说明模板、搜索区域、滑动比较和相似度图，演示单目标定位。','【学生】读取响应图峰值并验证目标位置。'),('阈值与多目标',14,'【教师】讲解阈值对漏检/误检的影响。','【学生】比较两种阈值并记录匹配数量。'),('受控变量实验',22,'【教师】固定其他条件，改变尺度、旋转、遮挡、亮度。','【学生】建立“变量—匹配分数—是否成功”表。'),('适用边界',16,'【教师】引导比较标准化工位和自由姿态场景。','【学生】写出模板匹配适合/不适合的各2类场景。')],
 ['【选择题】模板匹配最容易受到下列哪类变化影响（ ）。A. 尺度和旋转变化 B. 文件扩展名 C. Python变量名 D. 显示窗口标题','【判断题】只要把匹配阈值降低，就可以同时消除漏检和误检。（ ）','【简答题】怎样设计一个公平实验比较光照变化对模板匹配的影响？'],
 ['对一目标完成原始、旋转、缩放、遮挡、亮度变化至少4类受控实验，形成匹配分数和误检/漏检表。'],
 '以误匹配造成错误抓取/分拣的案例强调阈值与环境假设必须通过数据验证。',
 '讨论通过治具约束、标准光照和ROI限制降低算法复杂度的工程优化思路。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、模板与受控变量图像。',REF1+'；'+REF2),
L(13,'第五单元 模板匹配、特征与工业识别','角点、ORB/SIFT局部特征、描述子匹配与几何验证',
 {'局部特征':['角点/关键点','尺度/方向概念'],'描述子':['局部描述','ORB/SIFT思想'],'匹配':['最近邻/距离','筛选'],'几何验证':['对应关系','单应性/内点概念']},
 ['理解角点/关键点、局部描述子与特征匹配的基本思想。','理解ORB/SIFT相对简单模板匹配在尺度/旋转变化下的优势与成本。'],
 ['能够调用OpenCV提取局部特征并完成基本描述子匹配。','能够通过匹配筛选与几何验证减少明显误匹配。'],
 ['形成对开源算法与现成接口“会调用更要会验证”的工具边界意识。'],
 '局部特征、描述子匹配与几何验证；与模板方法比较。','匹配点多不等于定位正确，需要用几何一致性/内点等证据进行二次验证。',
 '展示一张含大量错误连线的原始特征匹配图，让学生判断“匹配数量很多”为何反而不可信。',
 [('关键点与描述子',18,'【教师】直观说明角点/局部结构、尺度/方向和描述子作用。','【学生】比较平坦区域与角点区域为何可辨识性不同。'),('ORB/SIFT思想',16,'【教师】介绍两类局部特征的应用思路，不展开算法细节推导。','【学生】根据速度/鲁棒性需求选择方案。'),('描述子匹配',18,'【教师】演示匹配距离、筛选和误匹配现象。','【学生】比较筛选前后连线质量。'),('几何验证',18,'【教师】说明单应性/内点等几何一致性验证思想。','【学生】用“有效内点比例/定位框”评价匹配是否可信。')],
 ['【选择题】局部特征匹配中，为减少明显误匹配通常需要（ ）。A. 匹配筛选和几何验证 B. 只增加连线数量 C. 关闭图像 D. 改变量名','【判断题】匹配点数量越多，就一定说明目标定位越准确。（ ）','【简答题】与简单模板匹配相比，局部特征方法通常在哪些变化条件下更有优势？'],
 ['完成同一目标的模板匹配与ORB/SIFT之一的对比，报告尺度/旋转条件下的成功与失败。'],
 '强调“现成算法接口不等于可靠结论”，开源代码、参数和结果仍需独立核验。',
 '比较模板、ORB/SIFT与治具约束三种方案的开发成本、速度、鲁棒性和维护性，提出场景化选择。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、特征匹配图像。',REF1+'；'+REF2,milestone=True),
L(14,'第六单元 深度学习视觉入门与智能制造综合案例','分类/目标检测/缺陷检测、数据标注划分与预训练模型推理',
 {'任务类型':['分类','目标检测','缺陷检测'],'数据':['标注','训练/验证/测试'],'预训练模型':['输入预处理','推理输出'],'工程边界':['数据分布','算力/部署']},
 ['理解分类、目标检测、缺陷检测的输出差异。','理解数据标注、训练/验证/测试划分和预训练模型推理的基本流程。'],
 ['能够调用经过验证的预训练视觉模型完成基础推理并读取主要输出。','能够检查输入尺寸、类别标签、置信度等推理配置。'],
 ['形成数据质量、模型来源和推理配置可追溯的工程规范。'],
 '深度学习视觉任务、数据划分与预训练模型推理。','避免把“调用模型成功”误认为“模型适合本场景”，需要理解数据分布、类别和输入前处理边界。',
 '对一张制造缺陷图分别提出“分类”和“目标检测”任务，要求学生说明两者标注成本与输出差异。',
 [('视觉任务类型',18,'【教师】比较分类、目标检测和缺陷检测/异常检测的输入输出。','【学生】为4个制造案例选择任务类型。'),('数据标注与划分',18,'【教师】说明训练/验证/测试分工、数据泄漏和类别分布。','【学生】识别一个“同一批次图像被分到训练和测试”的泄漏问题。'),('预训练模型推理',20,'【教师】演示加载经过验证的预训练模型、输入预处理与输出解析。','【学生】记录模型来源、版本、输入尺寸和输出字段。'),('场景适用性',14,'【教师】讨论数据域差异、算力和部署边界。','【学生】写出“可以试用”与“可以上线”之间还缺的证据。')],
 ['【选择题】目标检测相比图像分类通常额外需要输出（ ）。A. 目标位置/框 B. 文件名 C. 显示器尺寸 D. Python版本号','【判断题】使用预训练模型推理成功即可证明它适合本企业全部缺陷类型。（ ）','【简答题】为什么训练集、验证集和测试集应承担不同角色？'],
 ['选取一个制造视觉任务，说明分类/检测任务定义、数据标注要求、预训练模型来源和预期输出。'],
 '强调模型来源、数据质量和外部代码/AI工具使用必须披露并核验，避免“模型权威化”。',
 '比较传统视觉与预训练模型在开发周期、数据需求、算力和维护成本上的产品化差异。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、经过验证的预训练视觉模型及小型示例数据。',REF1+'；'+REF2),
L(15,'第六单元 深度学习视觉入门与智能制造综合案例','置信度、IoU、混淆矩阵、精确率/召回率与传统/学习型视觉比较',
 {'检测指标':['置信度','IoU'],'分类统计':['TP/FP/FN','混淆矩阵'],'性能指标':['Precision','Recall'],'方案比较':['传统视觉','深度学习','速度/鲁棒性/成本']},
 ['理解置信度、IoU、TP/FP/FN、混淆矩阵、精确率和召回率。','理解传统视觉与学习型视觉在数据、解释性、速度、鲁棒性和维护成本上的差异。'],
 ['能够根据小型结果表计算Precision/Recall并解释误检/漏检。','能够针对质量检测场景说明为何不同指标权重可能不同。'],
 ['树立不以单一准确率或单张结果评价系统的实证意识。','理解关键缺陷漏检的质量与安全风险。'],
 'IoU、混淆矩阵、Precision/Recall和方案比较。','指标没有脱离业务代价的绝对优劣，需要把误检/漏检与质量、安全、节拍成本联系起来。',
 '给出两个模型：A精确率高但召回率低，B召回率高但误检多；提问在关键安全缺陷筛查中应优先关注什么。',
 [('置信度与IoU',16,'【教师】解释检测框置信度和IoU，演示阈值变化对保留框的影响。','【学生】手算/判断3个预测框是否满足IoU条件。'),('混淆矩阵与基本指标',20,'【教师】从TP/FP/FN推导Precision/Recall直观含义。','【学生】根据小表计算并解释误检/漏检。'),('阈值与业务代价',18,'【教师】改变置信阈值，展示Precision/Recall变化。','【学生】针对关键缺陷和普通外观缺陷分别选择偏好。'),('传统/学习型比较',16,'【教师】从数据、解释性、速度、部署和维护比较两类方案。','【学生】形成“任务条件→方案选择”的决策表。')],
 ['【选择题】关键缺陷“漏掉一个可能造成严重后果”时，通常需要重点关注（ ）。A. 召回率 B. 文件大小 C. 代码行数 D. 图像颜色','【判断题】精确率和召回率可以脱离具体阈值与场景成本单独判断哪个模型永远更好。（ ）','【简答题】传统视觉在什么条件下可能比深度学习方案更合适？'],
 ['用给定/自建小型预测结果计算混淆矩阵、Precision和Recall，并讨论提高/降低阈值的代价。'],
 '以关键缺陷漏检和误报停线的双重成本说明评价指标需要服务真实质量责任，而不是追求单一数字。',
 '把指标阈值与产品质量等级、人工复检成本关联，设计分层判定/复核策略。',
 '电脑、投影仪、多媒体课件、教材、指标练习表、预训练模型推理结果。',REF1+'；'+REF2),
L(16,'第六单元 深度学习视觉入门与智能制造综合案例','智能制造机器视觉综合方案：零件/装配/FDM打印场景与期末验收',
 {'方案定义':['对象/任务','精度/速度/环境'],'成像与算法':['相机/镜头/光源','传统/学习型流程'],'评价证据':['误差/混淆','速度/鲁棒性','失败样例'],'工程交付':['原图/参数/代码','报告/答辩/复核']},
 ['综合理解成像、预处理、分割/特征/测量、识别与评价之间的关系。','理解机器视觉综合方案和期末考核的工程交付要求。'],
 ['能够针对零件尺寸/孔位、装配完整性、传送目标或FDM打印异常设计完整机器视觉方案。','能够从原图、参数、代码、指标、失败样例和改进计划形成可复核技术报告并进行个人解释。'],
 ['形成持续学习、工具边界、数据诚信和对视觉结论独立核验的职业意识。','理解视觉系统输出进入质量与自动化决策后的人类最终责任。'],
 '端到端视觉方案、指标/失败样例、工程文档与答辩核验。','综合方案需要在成像、算法、数据、指标、速度和成本之间做有证据的权衡，不能简单堆叠算法。',
 '给出一个“算法在实验图片上100%成功、现场换光照后大量漏检”的项目复盘，要求学生从成像、数据、算法、阈值和验收集五个层面提出整改。',
 [('需求与成像方案',18,'【教师】用四类综合案例示范从对象、精度、节拍、环境反推相机/镜头/光源与采集约束。','【学生】选择一个场景写任务/成像方案卡。'),('算法链与方案选择',18,'【教师】组织传统预处理/分割/测量/匹配与预训练模型方案对比。','【学生】画端到端流程，解释每个模块输入输出和替代方案。'),('评价、失败与改进',18,'【教师】要求方案包含测量误差或Precision/Recall、速度、鲁棒性与失败样例。','【学生】为至少3类失败工况设计测试和改进。'),('工程交付与期末核验',16,'【教师】说明综合报告/等效作品+现场答辩/个人核验要求，强调不另设笔试/机考。','【学生】完成交付清单：原图、参数、代码、图表、结论、版本、失败案例和个人可解释内容。')],
 ['【选择题】完整机器视觉方案中，哪一项不能被“单张成功截图”替代（ ）。A. 多工况评价与失败样例 B. 文件名 C. 幻灯片背景 D. 代码缩进颜色','【判断题】综合方案只要算法指标高，就可以忽略光源、相机和现场部署条件。（ ）','【简答题】期末综合视觉方案至少应提供哪些可追溯工程证据？'],
 ['完成期末综合方案/项目报告准备：任务定义、成像条件、算法链、关键参数、评价指标、失败样例、改进计划与个人解释提纲。'],
 '以质量、安全、数据诚信和可追溯为课程最终底线，要求学生对开源代码、预训练模型和AI辅助结果保持独立核验。',
 '将综合视觉方案视为智能制造现场可落地的小型技术产品，训练从技术效果、部署成本、维护性和换型能力综合评审。',
 '电脑、投影仪、多媒体课件、教材、Python、OpenCV、预训练模型示例、综合项目验收表。',REF1+'；'+REF2,milestone=True),
]
assert len(lessons)==16

COURSE_MAP={'成像基础':['系统组成','相机/镜头/光源'],'图像预处理':['像素/颜色','几何变换/滤波'],'分割与结构':['阈值/Canny','形态学'],'测量定位':['轮廓/连通域','尺寸/位置/误差'],'定位识别':['模板匹配','ORB/SIFT特征'],'学习与评价':['预训练视觉','IoU/Precision/Recall','综合方案']}

# ---------------- maps ----------------
def render_map(path,title,branches):
    dot=Digraph('G',format='png')
    dot.attr(rankdir='TB',bgcolor='white',margin='0.02',pad='0.05',nodesep='0.20',ranksep='0.36',dpi='190')
    dot.attr('node',shape='box',style='rounded,filled',fontname='Noto Sans CJK SC',fontsize='10.5',color='#6B7280',penwidth='1.0',fillcolor='#F8FAFC',margin='0.09,0.055')
    dot.attr('edge',color='#94A3B8',arrowsize='0.5',penwidth='0.9')
    dot.node('root',title,fillcolor='#E7F0F8',color='#4D7894',fontsize='13.5',penwidth='1.3')
    for i,(b,leaves) in enumerate(branches.items(),1):
        bid=f'b{i}'; dot.node(bid,b,fillcolor='#EEF5F0',color='#6A8D75',fontsize='11.5'); dot.edge('root',bid)
        for j,leaf in enumerate(leaves,1):
            nid=f'{bid}_{j}'; dot.node(nid,leaf,fillcolor='white',color='#A7B5C1',fontsize='9.5'); dot.edge(bid,nid)
    out=str(path.with_suffix('')); dot.render(out,cleanup=True)
    gen=Path(out+'.png')
    if gen!=path: shutil.move(gen,path)

render_map(IMG_DIR/'course-knowledge-map.png','机器视觉技术课程知识链',COURSE_MAP)
for l in lessons: render_map(IMG_DIR/f'lesson-{l["no"]:02d}-knowledge-map.png',l['topic'],l['map'])

# ---------------- markdown ----------------
def goals(l):
    return ('**知识目标：**\n'+'\n'.join(f'（{i+1}）{x}' for i,x in enumerate(l['objk']))+'\n\n'
            '**能力目标：**\n'+'\n'.join(f'（{i+1}）{x}' for i,x in enumerate(l['obja']))+'\n\n'
            '**价值目标：**\n'+'\n'.join(f'（{i+1}）{x}' for i,x in enumerate(l['objv'])))

def process_rows(l):
    rows=[]
    pre=('【教师】1. 发布本课主题、教材/资料范围和一个成像/算法最小问题；<br>2. 明确课堂代码仅作理论课中的原理演示和案例复现，不改变课程32学时全理论的事实。<br>【学生】1. 预读关键概念/代码片段；<br>2. 记录至少1个预测和1个疑问，带着问题进入课堂。')
    rows.append(('课前任务',pre,'通过预读和最小问题建立先备认知，同时明确理论课中的代码演示属于案例教学活动。'))
    intro=f'【教师】{l["intro"]}<br>【学生】先独立判断，再与同伴交换依据；教师收集典型分歧后进入本课。'
    rows.append(('互动导入（10 min）',intro,'从真实成像/检测现象切入，使学生先暴露直觉判断，再建立本课问题主线。'))
    teach=[]
    for idx,(title,mins,teacher,student) in enumerate(l['blocks'],1):
        teach.append(f'{idx}、{title}（约{mins} min）<br>【教师】{teacher}<br>【学生】{student}')
    teach.append(f'【课堂强化】形成可检查的课堂产物：{l["homework"][:60]}；课堂中至少保留参数/图像/计算或分析证据。')
    rows.append(('传授新知（共70 min）','<br>'.join(teach),'按“现象—原理—Python/OpenCV最小实现—参数比较—结果验证”推进，把教师讲解转化为可检查的分析或代码证据。'))
    q='<br>'.join(f'{i+1}.{q}' for i,q in enumerate(l['checks']))
    rows.append(('过关检测（10 min）',f'【教师】组织课堂小测：<br>{q}<br>【学生】独立作答/运行或读图验证；同伴互查后说明依据。','覆盖本课重点和典型误区，以即时证据发现“会看结果但不会解释”的问题。'))
    rows.append(('课堂小结（5 min）',f'【教师】1. 本次课解决的关键问题是“{l["topic"]}”；<br>2. 需要重点掌握：{l["focus"]}<br>3. 最值得带走的方法是先固定条件、再做参数/方案对比；<br>4. 需要警惕的易错点是：{l["diff"]}','回到知识脉络图压缩信息，形成“概念—条件—参数—证据—边界”的完整认知。'))
    rows.append(('作业布置（3 min）',f'【教师】布置课后任务：{l["homework"]}<br>【要求】不只提交结果，还需保留原图/参数/中间结果/失败样例或判断依据。','固化课堂成果，形成可复核学习证据，并为下一课提供真实输入。'))
    rows.append(('考勤（2 min）','【教师】利用学校/课程实际使用的平台或课堂点名完成签到。','掌握出勤情况，保障教学秩序；不在教案中预填任何实际出勤事实。'))
    return rows

lines=['---','schema_version: 1','lesson_plan_version: 1','document_state: final','course_name: 机器视觉技术','semester: 第7学期','canonical_source: true','---','','# 编制信息','','## 当前任务与边界','- 用户需要成果：完整教案 Markdown、知识脉络图、学校格式 Word、源码包与完整交付包。','- 正常工作流：官方大纲 → canonical Markdown/图片 → 内容门禁 → Word编译 → 全页视觉QA。','- 课程硬边界：32学时全部为理论；Python/NumPy/OpenCV代码演示与案例复现属于理论课堂教学活动，不虚构上机学时。','','## 来源清单','','| source_id | 材料或 Skill | 证据角色 | 版本或日期 | 定位 |','|---|---|---|---|---|',f'| S1 | `{COURSE["syllabus_path"]}` | official_syllabus | {COURSE["updated"]}；SHA `{COURSE["syllabus_sha"]}` | 课程事实、目标、教学单元、评价 |',f'| S2 | `{COURSE["note_path"]}` | official_syllabus | SHA `{COURSE["note_sha"]}` | 教材与Python/OpenCV实施说明 |',f'| S3 | `{COURSE["skill_repo"]}` / lesson-plan-compiler | reference_only | SHA `{COURSE["skill_sha"]}` | 教案字段、语言、Markdown审查与Word版式规则 |','','## 缺项与冲突','','| issue_id | 字段或主题 | 当前情况 | 影响 | 处理 |','|---|---|---|---|---|','| I1 | 主讲教师 | 用户未提供 | 仅影响封面 | Word留空，不虚构 |','| I2 | 职称 | 用户未提供 | 仅影响封面 | Word留空，不虚构 |','| I3 | 制定日期 | 用户未提供 | 仅影响封面 | Word留空，不把大纲定稿日期冒充教案制定日期 |','| I4 | 具体校历周次 | 用户未提供课表 | 课次表周次栏 | Word留空，不推测 |','','# 正式教案内容','','## 课程基本信息','','| 字段 | 内容 | 状态 | source_id |','|---|---|---|---|',
'| 学院 | 机械与动力工程学院 | official | S1 |','| 专业（教研室） | 智能制造工程 | official | S1 |','| 课程名称 | 机器视觉技术 | official | S1 |','| 课程名称（英文） | Machine Vision Technology | official | S1 |','| 主讲教师 |  | to_confirm |  |','| 职称 |  | to_confirm |  |','| 授课专业 | 智能制造工程 | official | S1 |','| 授课学期 | 第7学期 | official | S1 |','| 制定日期 |  | to_confirm |  |','| 课程类别 | 专业类 | official | S1 |','| 课程性质 | 选修 | official | S1 |','| 授课语言 | 中文 | official | S1 |','| 学分 | 2 | official | S1 |','| 总学时 | 32 | official | S1 |','| 理论学时 | 32 | official | S1 |','| 实验/上机学时 | 0 | official | S1 |','| 实习学时 | 0 | official | S1 |','| 其他学时 | 0 | official | S1 |',f'| 教材 | {COURSE["textbook"]} | official | S2 |',f'| 授课学院 | {COURSE["college"]} | official | S1 |',f'| 先修课程 | {COURSE["prereq"]} | official | S1 |',f'| 后续课程 | {COURSE["followup"]} | official | S1 |','| 考核类型 | 考查课 | official | S2 |','| 考核形式 | 机器视觉综合方案/项目报告（或等效作品）＋现场答辩/个人核验；不另设笔试或机考 | official | S1 |','| 考核方式 | 平时成绩、课程作业、期末综合方案/作品与个人核验 | official | S1 |',f'| 总评成绩比例 | {COURSE["assessment"]} | official | S1 |','','#### 课程简介','','**课程基本定位**：本课程是智能制造工程专业的一门专业选修课，面向工业检测、测量、定位、识别与视觉引导等典型任务，以Python为统一实现语言，以NumPy和OpenCV为核心工具，建立“成像条件—图像处理—特征与测量—识别判定—性能评价”的机器视觉工作链。','','**学时组织**：总学时32，全部为理论学时，按16次课×2课时组织。课堂包含Python/OpenCV最小实现、参数对比和案例复现，但这些活动属于理论课堂中的案例教学，不作为独立上机课时。','','**核心学习结果**：学生能够分析相机、镜头、光源与成像参数，使用Python/OpenCV完成图像预处理、分割、形态学、轮廓与几何测量、模板与特征匹配，理解深度学习视觉基本流程，并针对零件、装配、输送或FDM 3D打印场景形成可复现视觉方案，以测量误差、误检/漏检、精确率/召回率、速度和鲁棒性评价结果。','','**主要教学方法**：采用“成像现象—算法原理—Python最小实现—OpenCV工程实现—参数对比—结果验证”的案例驱动方式；所有案例保留原始图像、参数、中间结果、评价指标和失败样例。','','![课程知识脉络图](images/course-knowledge-map.png)','','## 教学单元','']
for l in lessons:
    methods='讲授法、问答法、案例分析法、结构图解法、Python/OpenCV代码演示法、参数对比法、课堂检测法'
    tools=l['tools']
    lines += [f'### lesson-{l["no"]:02d}','', '| 字段 | 内容 | 状态 | source_id |','|---|---|---|---|',f'| 章节 | {l["unit"]} | derived | S1 |',f'| 授课题目 | {l["topic"]} | derived | S1 |','| 周次 |  | to_confirm |  |','| 课时安排 | 2课时（100 min）；类型：理论2 | official | S1 |',f'| 知识脉络图 | images/lesson-{l["no"]:02d}-knowledge-map.png | derived | S1 |',f'| 教学方法 | {methods} | derived | S3 |',f'| 教学用具 | {tools} | derived | S1/S2 |','| 教学设计 | 课前任务→互动导入（10 min）→传授新知（70 min）→过关检测（10 min）→课堂小结（5 min）→作业布置（3 min）→考勤（2 min） | derived | S3 |',f'| 参考文献 | {l["refs"]} | provided | S2 |','',f'![知识脉络图](images/lesson-{l["no"]:02d}-knowledge-map.png)','','#### 教学目标','',goals(l),'','#### 课程思政','',f'（1）{l["ideology"]}\n\n（2）要求课堂代码演示、图像参数和评价结果保留原始证据，不以单张成功截图替代多工况验证，强化数据诚信和质量责任。','','#### 专创融合','',l['innovation'],'','#### 教学重难点','',f'**教学重点**：{l["focus"]}\n\n**教学难点**：{l["diff"]}','','#### 教学过程','','| 教学步骤 | 主要教学内容 | 设计意图 |','|---|---|---|']
    for step,content,intent in process_rows(l):
        lines.append('| '+step.replace('|','\\|')+' | '+content.replace('|','\\|')+' | '+intent.replace('|','\\|')+' |')
    lines += ['','#### 教学反思','','【教师】课后重点从以下方面进行教学反思：','','1. 教学流程：导入是否把学生带入真实视觉问题，70分钟主体是否按“现象—原理—实现—参数—验证”由浅入深。',f'2. 教学内容：重点记录“{l["diff"]}”是否讲清，学生是否能说明参数/方法适用边界。','3. 学生参与：仅依据实际课堂记录填写参与、代码阅读/参数分析、互审和表达情况，不补写推测性结论。','4. 教学改进：依据真实课堂证据决定是否增加成像对比、失败样例、最小代码或指标练习。','']
    if l['milestone']:
        lines += ['#### 单元课程作业节点','',f'- {UNIT_ASSESS[l["unit"]]}。','- 本单元作业必须包含原图/任务条件、关键参数、中间结果、评价或失败样例；不要求虚构独立上机课时。','']
lines += ['# 内部追踪区','','## 版本与哈希','','- lesson_plan_version：1','- 大纲SHA：`'+COURSE['syllabus_sha']+'`','- lesson-plan-compiler SHA：`'+COURSE['skill_sha']+'`','','## 变更原则','','- 内容修改必须先回到本Markdown及图片主源，再重新编译Word。','- Word阶段只处理字体、表格、分页、图片尺寸等呈现问题；不得在Word中独立改写课程语义。']
FINAL_MD.write_text('\n'.join(lines),encoding='utf-8')

# copy to source package before compile
shutil.copy2(FINAL_MD,COURSE_DIR/FINAL_MD.name)
(COURSE_DIR/'README.md').write_text(f'''# 《机器视觉技术》教案源码包\n\n本包采用 Markdown-first 工作流。\n\n- `{FINAL_MD.name}`：canonical内容主源；\n- `images/`：课程总知识图 + 16次课知识脉络图；\n- `manifest.json`：课程、大纲/Skill版本、课次与图片映射；\n- `build_docx.py`：从canonical Markdown内容和学校格式模板编译Word。\n\n课程32学时全部为理论，Python/OpenCV代码演示仅作为理论课堂案例活动。正式内容修改请先改Markdown再重新编译Word。\n''',encoding='utf-8')

# ---------------- content QA ----------------
md=FINAL_MD.read_text(encoding='utf-8')
assert len(re.findall(r'^### lesson-\d{2}$',md,re.M))==16
assert md.count('2课时（100 min）；类型：理论2')==16
assert '类型：上机' not in md and '上机一' not in md
assert 16*2==COURSE['hours_total']==32
assert COURSE['hours_theory']==32 and COURSE['hours_lab']==0
assert md.count('#### 教学目标')==16 and md.count('#### 教学过程')==16 and md.count('#### 教学反思')==16
refs=re.findall(r'!\[[^]]*\]\((images/[^)]+)\)',md)
assert len(refs)==17 and len(set(refs))==17
assert not [x for x in refs if not (COURSE_DIR/x).exists()]
for l in lessons:
    assert sum(x[1] for x in l['blocks'])==70
assert '平时成绩20%＋课程作业20%＋期末考核60%' in md

# ---------------- parse canonical MD ----------------
def md_section(text,start_heading,end_patterns):
    m=re.search(rf'^{re.escape(start_heading)}\s*$',text,re.M)
    if not m: return ''
    start=m.end(); ends=[]
    for pat in end_patterns:
        m2=re.search(pat,text[start:],re.M)
        if m2: ends.append(start+m2.start())
    return text[start:min(ends) if ends else len(text)].strip()

def clean_md(s):
    s=re.sub(r'!\[[^]]*\]\([^)]+\)','',s)
    s=s.replace('**','').replace('`','')
    s=re.sub(r'^\s*[-*]\s+','',s,flags=re.M)
    s=re.sub(r'\n{3,}','\n\n',s)
    return s.strip()

parts=re.split(r'(?=^### lesson-\d{2}$)',md,flags=re.M)
lesson_parts=[p for p in parts if re.match(r'^### lesson-\d{2}\n',p)]
parsed=[]
for p in lesson_parts:
    no=int(re.search(r'^### lesson-(\d{2})$',p,re.M).group(1))
    # table key/value
    meta={}
    for line in p.splitlines():
        if line.startswith('| ') and ' | ' in line and not line.startswith('|---') and not line.startswith('| 字段') and not line.startswith('| 教学步骤'):
            cols=[x.strip() for x in line.strip('|').split('|')]
            if len(cols)>=4 and cols[0] in ['章节','授课题目','周次','课时安排','知识脉络图','教学方法','教学用具','教学设计','参考文献']:
                meta[cols[0]]=cols[1]
    secs={}
    headings=['#### 教学目标','#### 课程思政','#### 专创融合','#### 教学重难点','#### 教学过程','#### 教学反思','#### 单元课程作业节点']
    for h in headings:
        secs[h[5:]]=md_section(p,h,[r'^#### ',r'^### lesson-',r'^# '])
    # parse process rows
    pro=[]
    for line in secs['教学过程'].splitlines():
        if line.startswith('|') and not line.startswith('|---') and '教学步骤' not in line:
            cols=[x.strip().replace('\\|','|') for x in line.strip('|').split('|')]
            if len(cols)>=3: pro.append(cols[:3])
    parsed.append({'no':no,'meta':meta,'secs':secs,'process':pro})
assert len(parsed)==16 and all(len(x['process'])==7 for x in parsed)

# ---------------- Word compiler ----------------
shutil.copy2(TEMPLATE,FINAL_DOCX)
doc=Document(FINAL_DOCX)
proto=deepcopy(doc.tables[2]._element)
# remove all lesson tables
for t in list(doc.tables[2:]):
    el=t._element; el.getparent().remove(el)
# remove all stale lesson separator paragraphs/content after the section properties marker.
# The inherited template was previously assembled with lesson content after sectPr; keep only
# cover/basic-info/notes + sectPr, then insert new lessons *before* sectPr.
body=doc._element.body
children=list(body)
sect=next((ch for ch in children if ch.tag==qn('w:sectPr')),None)
if sect is None:
    raise RuntimeError('template missing sectPr')
seen_sect=False
for ch in list(body):
    if ch is sect:
        seen_sect=True
        continue
    if seen_sect and ch.getparent() is body:
        body.remove(ch)
# Cover
cover=doc.tables[0]
cover_vals=[COURSE['college'],COURSE['major'],COURSE['name'],'','',COURSE['semester'],COURSE['major']]

def set_run_font(r,size=9,bold=None,font='宋体'):
    r.font.name=font; r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),font); r.font.size=Pt(size)
    if bold is not None: r.bold=bold

def clear_cell(cell):
    tc=cell._tc; tcPr=tc.tcPr
    for child in list(tc):
        if child is not tcPr: tc.remove(child)
    tc.append(OxmlElement('w:p'))

def set_cell(cell,text,size=9,bold=False,align=None,bold_prefixes=()):
    clear_cell(cell); lines=str(text or '').split('\n')
    for i,line in enumerate(lines):
        p=cell.paragraphs[0] if i==0 else cell.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.0
        if align is not None: p.alignment=align
        r=p.add_run(line); set_run_font(r,size,bold or any(line.strip().startswith(x) for x in bold_prefixes))
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_cell_margins(cell,top=50,start=65,bottom=50,end=65):
    tcPr=cell._tc.get_or_add_tcPr(); tcMar=tcPr.first_child_found_in('w:tcMar')
    if tcMar is None: tcMar=OxmlElement('w:tcMar'); tcPr.append(tcMar)
    for tag,val in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        node=tcMar.find(qn('w:'+tag))
        if node is None: node=OxmlElement('w:'+tag); tcMar.append(node)
        node.set(qn('w:w'),str(val)); node.set(qn('w:type'),'dxa')

def cant_split(row):
    trPr=row._tr.get_or_add_trPr(); el=OxmlElement('w:cantSplit'); trPr.append(el)

for i,v in enumerate(cover_vals): set_cell(cover.cell(i,1),v,12,align=WD_ALIGN_PARAGRAPH.CENTER)
# clear old date text rather than fabricate
for p in doc.paragraphs:
    if re.search(r'20\d{2}年\d{1,2}月',p.text):
        p.text=''; break
# course basic info
info=doc.tables[1]
set_cell(info.cell(0,1),COURSE['name'],9); set_cell(info.cell(1,1),COURSE['english'],9)
set_cell(info.cell(2,1),COURSE['category'],9,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(info.cell(2,3),COURSE['nature'],9,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(info.cell(2,5),COURSE['language'],9,align=WD_ALIGN_PARAGRAPH.CENTER)
set_cell(info.cell(3,1),COURSE['semester'],9,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(info.cell(3,5),COURSE['credits'],9,align=WD_ALIGN_PARAGRAPH.CENTER)
# template headers: total, theory, experiment, machine, other - keep source facts 32,32,0,0,0
for col,v in enumerate([32,32,0,0,0],1): set_cell(info.cell(5,col),str(v),9,align=WD_ALIGN_PARAGRAPH.CENTER)
set_cell(info.cell(6,1),COURSE['major'],9); set_cell(info.cell(7,1),COURSE['textbook'],8.7); set_cell(info.cell(8,1),COURSE['college'],9); set_cell(info.cell(9,1),COURSE['prereq'],8.7); set_cell(info.cell(10,1),COURSE['followup'],9)
set_cell(info.cell(11,1),'考试课（ ）；考查课（√）',9)
set_cell(info.cell(12,1),'闭卷（ ）；开卷（ ）；笔试（ ）；机试（ ）；口试（√，现场答辩/个人核验）；其它（√，机器视觉综合方案/项目报告或等效作品）',8.4)
set_cell(info.cell(13,1),'平时成绩（√）；课程作业（√）；综合方案/项目报告或等效作品（√）；现场答辩/个人核验（√）',8.5)
set_cell(info.cell(14,1),COURSE['assessment'],8.8)
intro=('课程基本定位：本课程是智能制造工程专业的一门专业选修课，面向工业检测、测量、定位、识别与视觉引导等典型任务，以Python为统一实现语言，以NumPy和OpenCV为核心工具，建立“成像条件—图像处理—特征与测量—识别判定—性能评价”的机器视觉工作链。\n'
       '学时组织：总学时32学时，全部为理论学时，按16次课×2课时组织。课堂Python/OpenCV代码演示、参数实验和案例复现属于理论课堂案例教学活动，不虚构为独立上机学时。\n'
       '核心学习结果：学生能够分析相机、镜头、光源与成像参数，完成图像预处理、分割、形态学、轮廓与几何测量、模板与特征匹配，理解深度学习视觉流程，并用误差、误检/漏检、Precision/Recall、速度和鲁棒性评价智能制造视觉方案。\n'
       '主要教学方法：成像现象—算法原理—Python最小实现—OpenCV工程实现—参数对比—结果验证；所有案例保留原图、参数、中间结果、评价指标和失败样例。')
set_cell(info.cell(15,1),intro,8.4,bold_prefixes=('课程基本定位','学时组织','核心学习结果','主要教学方法'))
# clean paragraphs after notes; retain original first 26 body children style structure. Determine current body after deletions.
# append 16 lessons with page breaks and cloned table
for P in parsed:
    pbreak=OxmlElement('w:p'); r=OxmlElement('w:r'); br=OxmlElement('w:br'); br.set(qn('w:type'),'page'); r.append(br); pbreak.append(r); body.insert(len(body)-1,pbreak)
    tbl=deepcopy(proto); body.insert(len(body)-1,tbl)
    t=doc.tables[-1]
    meta=P['meta']; sec=P['secs']; pro=P['process']
    set_cell(t.cell(0,1),meta['章节'],8.7,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(t.cell(0,3),meta['授课题目'],8.7,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell(t.cell(1,1),'',9,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(t.cell(1,3),meta['课时安排'],8.7,align=WD_ALIGN_PARAGRAPH.CENTER)
    clear_cell(t.cell(2,1)); pp=t.cell(2,1).paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.add_run().add_picture(str(COURSE_DIR/meta['知识脉络图']),width=Cm(13.0))
    set_cell(t.cell(3,1),clean_md(sec['教学目标']),8.7,bold_prefixes=('知识目标','能力目标','价值目标'))
    set_cell(t.cell(4,1),clean_md(sec['课程思政']),8.5); set_cell(t.cell(5,1),clean_md(sec['专创融合']),8.5); set_cell(t.cell(6,1),clean_md(sec['教学重难点']),8.5,bold_prefixes=('教学重点','教学难点'))
    set_cell(t.cell(7,1),meta['教学方法'],8.4); set_cell(t.cell(8,1),meta['教学用具'],8.25); set_cell(t.cell(9,1),meta['教学设计'],8.3)
    set_cell(t.cell(10,1),'教学步骤及主要教学内容',9.1,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(t.cell(10,4),'设计意图',9.1,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER)
    # template has 20 rows matching 7 process + reflection + refs
    target=[11,12,13,14,15,16,17]
    for row_idx,(step,content,intent) in zip(target,pro):
        label=step.replace('（10 min）','\n（10 min）').replace('（共70 min）','\n（70 min）').replace('（5 min）','\n（5 min）').replace('（3 min）','\n（3 min）').replace('（2 min）','\n（2 min）')
        set_cell(t.cell(row_idx,0),label,8.3,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(t.cell(row_idx,1),content.replace('<br>','\n'),8.0,bold_prefixes=('【教师】','【学生】','【要求】','【课堂强化】')); set_cell(t.cell(row_idx,4),intent,7.9)
    set_cell(t.cell(18,0),'课后\n教学反思',8.4,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(t.cell(18,1),clean_md(sec['教学反思']),8.0,bold_prefixes=('【教师】',))
    set_cell(t.cell(19,0),'本章节\n参考文献',8.4,bold=True,align=WD_ALIGN_PARAGRAPH.CENTER); set_cell(t.cell(19,1),meta['参考文献'],8.1)
    for row in t.rows:
        cant_split(row); seen=set()
        for c in row.cells:
            if id(c._tc) in seen: continue
            seen.add(id(c._tc)); set_cell_margins(c)

doc.save(FINAL_DOCX)
# privacy scrub
clean=BASE/'机器视觉技术_教案_16次课_最终版_clean.docx'
subprocess.run(['python','/home/oai/skills/docx/scripts/privacy_scrub.py',str(FINAL_DOCX),'--out',str(clean)],check=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
clean.replace(FINAL_DOCX)
# source package build script
shutil.copy2(Path(__file__),COURSE_DIR/'build_docx.py')

def sha(p):
    h=hashlib.sha256(); h.update(Path(p).read_bytes()); return h.hexdigest()
manifest={'course':COURSE,'document_state':'final_compiled','canonical_markdown':FINAL_MD.name,'compiled_docx':FINAL_DOCX.name,'lesson_count':16,'hours':{'total':32,'theory':32,'lab':0},'images':17,
'source':{'syllabus_path':COURSE['syllabus_path'],'syllabus_sha':COURSE['syllabus_sha'],'implementation_note_path':COURSE['note_path'],'implementation_note_sha':COURSE['note_sha'],'skill_repo':COURSE['skill_repo'],'skill_sha':COURSE['skill_sha']},
'unresolved_handling':{'主讲教师':'Word封面留空，未虚构','职称':'Word封面留空，未虚构','制定日期':'Word留空，未把大纲定稿日期冒充教案制定日期','具体周次':'Word各课次周次栏留空，未推测'},
'lesson_images':[{'lesson':l['no'],'title':l['topic'],'image':f'images/lesson-{l["no"]:02d}-knowledge-map.png'} for l in lessons],
'checksums':{FINAL_MD.name:sha(FINAL_MD),FINAL_DOCX.name:sha(FINAL_DOCX)}}
(COURSE_DIR/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
# complete dir
FULL_DIR.mkdir(); (FULL_DIR/'images').mkdir()
for p in COURSE_DIR.iterdir():
    if p.is_file(): shutil.copy2(p,FULL_DIR/p.name)
for p in IMG_DIR.glob('*.png'): shutil.copy2(p,FULL_DIR/'images'/p.name)
shutil.copy2(FINAL_DOCX,FULL_DIR/FINAL_DOCX.name)
(FULL_DIR/'README.md').write_text((COURSE_DIR/'README.md').read_text(encoding='utf-8')+'\n本完整包另含学校格式Word成品。\n',encoding='utf-8')
# zip
for zpath,root in [(SRC_ZIP,COURSE_DIR),(FULL_ZIP,FULL_DIR)]:
    if zpath.exists(): zpath.unlink()
    with zipfile.ZipFile(zpath,'w',zipfile.ZIP_DEFLATED) as z:
        for p in sorted(root.rglob('*')):
            if p.is_file(): z.write(p,arcname=f'{root.name}/{p.relative_to(root).as_posix()}')
print('built',FINAL_MD,FINAL_DOCX,SRC_ZIP,FULL_ZIP,sep='\n')
