/* ============================================================
   开物AI科学 · 站点数据（改完 Commit，约1分钟自动生效）
   数据版本：v2.0
   · 31个教学网页的分类/年级/单元/课时/领域/核心素养 全部对齐 Excel 总表
   · 提示词 10 条（原7条保留 + 新增命题/教学设计Skill/教学网页Skill）
   · 校准 w11–w31 下载链接的提取码后缀
   ============================================================ */

const CONFIG = {
  SECRET:     "kaiwuai-2026",
  ADMIN_USER: "vpilgrim",
  ADMIN_PASS: "kaiwuai2026",        // ← 如果你在别处已改过密码，这里填你自己的
  QR:         "assets/qr.png.jpg",  // 仓库里二维码实际文件名是 qr.png.jpg
};

const ABOUT = `开物AI科学是一个免费开放的中小学科学教学资源库。
这里的每一个互动实验、课堂工具，都由 AI 辅助生成、一线科学教师实测打磨。
我们相信：好的教学资源不该有界限——一部手机、一间教室，就能把太阳系搬到课桌上。
本站内容由 AI 辅助生成并经人工审核，仅供教学交流使用。`;

/* ---------- 教学网页 31（分类对齐Excel总表） ---------- */
const WEB_ITEMS = [
{id:"w01",t:"气温计",g:"三年级上",u:"第一单元 气象观测",ls:"1 测量气温",f:"地球与宇宙科学",c:"10.地球系统",ty:"仪器",d:"认读气温计刻度：随机出题、液面放大镜、分度值切换、连续数与记录。",lk:"https://wwdj.lanzoul.com/iWzQ48z4x0f"},
{id:"w02",t:"沉浸体验活字印刷",g:"三年级上",u:"第五单元 书籍的发展与演变",ls:"3 体验活字印刷",f:"技术与工程",c:"12.技术、工程与社会",ty:"体验",d:"逆反字字模、排版、刷墨印刷，分难度体验毕昇活字印刷古法流程。",lk:"https://wwdj.lanzoul.com/iCgA248z4wwb"},
{id:"w03",t:"探究条形磁铁磁性最强",g:"三年级下",u:"第一单元 磁针与南北",ls:"1 磁铁有两极",f:"物质科学",c:"3.物质的运动与相互作用",ty:"实验",d:"拖动磁铁各部位吸附回形针，对比磁力强弱，发现两端最强、中间最弱。",lk:"https://wwdj.lanzoul.com/i02w948z4wsh"},
{id:"w04",t:"磁针总是指南北吗",g:"三年级下",u:"第一单元 磁针与南北",ls:"2 两极指南北",f:"物质科学",c:"3.物质的运动与相互作用",ty:"AR场景",d:"AR 手势操控地磁场模型，解释小磁针为什么总是指示南北。",lk:"https://wwdj.lanzoul.com/if1aI48z4vmf"},
{id:"w05",t:"听话的小车",g:"三年级下",u:"第一单元 磁针与南北",ls:"4 两极相遇了",f:"物质科学",c:"3.物质的运动与相互作用",ty:"AR场景",d:"手势控制小车靠近或推开，用「隐形磁铁」揭示同极相斥、异极相吸。",lk:"https://wwdj.lanzoul.com/iLV6P48z4vsb",lv:"/r/magnetic-cart/",sh:"assets/magnetic-cart.png",pp:"p02"},
{id:"w06",t:"人体消化系统",g:"三年级下",u:"第三单元 食物与消化",ls:"2 一次神奇的旅行",f:"生命科学",c:"5.生命系统的构成层次",ty:"3D模型",d:"把8个消化器官拖入人体虚线位置拼装，演示食物消化的完整旅程。",lk:"https://wwdj.lanzoul.com/irq6Z48z4v7a",lv:"/r/digestive-system/",sh:"assets/digestive-system.png",pp:"p03"},
{id:"w07",t:"人体呼吸系统",g:"三年级下",u:"第四单元 空气与呼吸",ls:"2 我们的呼吸器官",f:"生命科学",c:"5.生命系统的构成层次",ty:"3D模型",d:"拖动器官完成呼吸系统拼装，含气体交换演示与各器官功能讲解。",lk:"https://wwdj.lanzoul.com/iCGw48z4v5i"},
{id:"w08",t:"寻找彩虹",g:"四年级上",u:"体验单元 寻找彩虹",ls:"体验·寻找彩虹",f:"物质科学",c:"3.物质的运动与相互作用",ty:"互动课件",d:"情境对话、喷水雾模拟、总结规律、检测挑战，六环节完整探究光色散。",lk:"https://wwdj.lanzoul.com/ieidq48z4vyh"},
{id:"w09",t:"空气受热体积变化",g:"四年级上",u:"第一单元 冷与热",ls:"1 乒乓球复原了",f:"物质科学",c:"1.物质的结构与性质",ty:"实验",d:"调水温观察针筒活塞位置变化，用记录表对比空气受热前后的体积。",lk:"https://wwdj.lanzoul.com/iUNn48z4wla"},
{id:"w10",t:"给水加热",g:"四年级上",u:"第一单元 冷与热",ls:"3 给水加热",f:"物质科学",c:"1.物质的结构与性质",ty:"实验",d:"酒精灯加热烧杯，采集温度—时间曲线，观察100℃沸腾并记录数据。",lk:"https://wwdj.lanzoul.com/iwSr348z4wfe",lv:"/r/water-heating/",sh:"assets/water-heating.png"},
{id:"w11",t:"小水珠从哪来",g:"四年级上",u:"第一单元 冷与热",ls:"4 小水珠从哪儿来",f:"物质科学",c:"1.物质的结构与性质",ty:"实验",d:"抱冰块、盖玻璃片，擦出水珠并得出结论：水蒸气遇冷凝结。",lk:"https://wwdj.lanzoul.com/iSpiF48z4wuj"},
{id:"w12",t:"水结冰",g:"四年级上",u:"第一单元 冷与热",ls:"5 水结冰了",f:"物质科学",c:"1.物质的结构与性质",ty:"实验",d:"碎冰加食盐降温，逐分钟记录温度，发现0℃开始结冰且体积膨胀。",lk:"https://wwdj.lanzoul.com/iqmle48z4wnc"},
{id:"w13",t:"水与金属的冷热变化",g:"四年级上",u:"第一单元 冷与热",ls:"拓展探究（热胀冷缩）",f:"物质科学",c:"1.物质的结构与性质",ty:"实验",d:"两个实验：液体的热胀冷缩、铜球穿铜环验证固体的热胀冷缩。",lk:"https://wwdj.lanzoul.com/i8lKi48z4wqf"},
{id:"w14",t:"阳光下的影子",g:"四年级上",u:"第二单元 太阳、地球和月球",ls:"1 阳光下的影子",f:"地球与宇宙科学",c:"9.宇宙中的地球",ty:"3D实验",d:"拖动时间轴，3D演示一天中太阳高度角与影长的变化规律。",lk:"https://wwdj.lanzoul.com/iKrdm48z4vha"},
{id:"w15",t:"地球的形状观察实验",g:"四年级上",u:"第二单元 太阳、地球和月球",ls:"2 地球的形状",f:"地球与宇宙科学",c:"9.宇宙中的地球",ty:"3D模型",d:"3D帆船进港实验：先见船帆后见船身，证明地球表面是球面。",lk:"https://wwdj.lanzoul.com/icWa48z4v4h"},
{id:"w16",t:"太阳与地球-两种学说",g:"四年级上",u:"第二单元 太阳、地球和月球",ls:"3 太阳与地球",f:"地球与宇宙科学",c:"9.宇宙中的地球",ty:"3D模型",d:"切换地心说与日心说模型，对比两种学说的运动关系与证据。",lk:"https://wwdj.lanzoul.com/iVqQa48z4vdg",pp:"p04"},
{id:"w17",t:"太阳系中的行星",g:"四年级上",u:"第二单元 太阳、地球和月球",ls:"4 太阳系中的行星",f:"地球与宇宙科学",c:"9.宇宙中的地球",ty:"3D模型",d:"3D八大行星按真实比例排列，点击行星查看资料与特征数据。",lk:"https://wwdj.lanzoul.com/iUpGq48z4vad",lv:"/r/solar-planets/",sh:"assets/solar-planets.png"},
{id:"w18",t:"太阳系（AR）",g:"四年级上",u:"第二单元 太阳、地球和月球",ls:"4 太阳系中的行星",f:"地球与宇宙科学",c:"9.宇宙中的地球",ty:"AR场景",d:"AR 手势操控太阳系广角视角，捏合拖拽查看行星位置与资料。",lk:"https://wwdj.lanzoul.com/ivNgr48z4vpi"},
{id:"w19",t:"月球的表面",g:"四年级上",u:"第二单元 太阳、地球和月球",ls:"5 月球的表面",f:"地球与宇宙科学",c:"9.宇宙中的地球",ty:"3D模型",d:"3D月面环形山与陨石撞击模拟，认识环形山、月海与月陆。",lk:"https://wwdj.lanzoul.com/iohGb48z4vfi"},
{id:"w20",t:"动物的感知本领",g:"四年级上",u:"第三单元 动物与环境",ls:"1 动物的感知本领",f:"生命科学",c:"5.生命系统的构成层次",ty:"匹配",d:"拖动「眼睛/触角/回声定位」等本领词，配到对应动物身上。",lk:"https://wwdj.lanzoul.com/itZ6U48z4v7g"},
{id:"w21",t:"寻访动物的家",g:"四年级上",u:"第三单元 动物与环境",ls:"2 寻访动物的家",f:"生命科学",c:"7.生物与环境的相互关系",ty:"匹配",d:"为喜鹊、野兔等匹配食物、位置、外形与筑巢材料，含双人匹配赛。",lk:"https://wwdj.lanzoul.com/ilbUI48z4waj"},
{id:"w22",t:"饲养小动物-小仓鼠",g:"四年级上",u:"第三单元 动物与环境",ls:"3 饲养小动物",f:"生命科学",c:"7.生物与环境的相互关系",ty:"游戏",d:"照料仓鼠的食物、饮水、居住、空气四项指标，制定饲养方案。",lk:"https://wwdj.lanzoul.com/iNnC48z4x6b"},
{id:"w23",t:"动物与四季变化",g:"四年级上",u:"第三单元 动物与环境",ls:"4 动物与四季变化",f:"生命科学",c:"7.生物与环境的相互关系",ty:"互动课件",d:"情境、四季转盘、温度实验、归纳配对多环节，认识动物过冬方式。",lk:"https://wwdj.lanzoul.com/iEZ8348z4ywf"},
{id:"w24",t:"影子的形成",g:"四年级上",u:"第四单元 光和声音",ls:"2 光会被挡住吗",f:"物质科学",c:"3.物质的运动与相互作用",ty:"3D实验",d:"调光源距离与角度、换遮挡物形状材质，观察影子的形成与变化。",lk:"https://wwdj.lanzoul.com/i4kwI48z4vkd"},
{id:"w25",t:"杠杆",g:"五年级上",u:"第一单元 简单机械",ls:"2 杠杆",f:"物质科学",c:"3.物质的运动与相互作用",ty:"实验",d:"在杠杆尺孔位挂一摞钩码，由两点距离的大小关系判断省力或费力。",lk:"https://wwdj.lanzoul.com/iiJjv48z4wcb"},
{id:"w26",t:"滑轮",g:"五年级上",u:"第一单元 简单机械",ls:"3 滑轮",f:"物质科学",c:"3.物质的运动与相互作用",ty:"实验",d:"直接提起、定滑轮、动滑轮三组对照，读测力计示数比较用力大小。",lk:"https://wwdj.lanzoul.com/iwaU48z4wji"},
{id:"w27",t:"毛线推演-食物网",g:"六年级上",u:"第四单元 食物与能量",ls:"4 食物链和食物网",f:"生命科学",c:"7.生物与环境的相互关系",ty:"体验",d:"以草原生态系统为背景连线，梳理生产者、消费者与食物链关系。",lk:"https://wwdj.lanzoul.com/ifFTn48z4wze"},
{id:"w28",t:"能量大转盘",g:"六年级上",u:"第四单元 食物与能量（拓展）",ls:"拓展·能量类型转盘",f:"物质科学",c:"4.能的转化与能量守恒",ty:"游戏",d:"转动转盘随机抽到核能、风能、机械能等能量类型，用于课堂抽题。",lk:"https://wwdj.lanzoul.com/i76W48z4x3i"},
{id:"w29",t:"近视的成因（版本一）",g:"八年级（初中物理）",u:"第五章 透镜及其应用",ls:"第4节 眼睛和眼镜",f:"物质科学",c:"3.物质的运动与相互作用",ty:"模型",d:"对比正常视力与屈光性、轴性近视模型，并演示凹透镜矫正。",lk:"https://wwdj.lanzoul.com/iAnN48z4w0j"},
{id:"w30",t:"近视的成因（版本二）",g:"八年级（初中物理）",u:"第五章 透镜及其应用",ls:"第4节 眼睛和眼镜",f:"物质科学",c:"3.物质的运动与相互作用",ty:"模型",d:"眼睛光学模型：调节眼轴与晶状体，观察光路会聚位置与矫正效果。",lk:"https://wwdj.lanzoul.com/i8DyO48z4w3c"},
{id:"w31",t:"火箭发射分离控制",g:"五年级上",u:"第三单元 浩瀚宇宙（拓展）",ls:"拓展·工程设计流程",f:"技术与工程",c:"13.工程设计与物化",ty:"仿真",d:"按点火升空、一级分离、抛整流罩、展开帆板四步操作火箭发射控制台。",lk:"https://wwdj.lanzoul.com/i42kk48z4vve"},
];

/* ---------- 课堂工具 17 ---------- */
const TOOLS = [
{id:"t01",t:"舰队积分指挥系统",g:"通用",ty:"小组积分",d:"以舰队航行为主题的小组积分系统，实时排名激励课堂竞争与合作。",lk:"https://wwdj.lanzoul.com/isbHM4z54uh"},
{id:"t02",t:"月相积分系统",g:"五年级上",u:"浩瀚宇宙",ty:"小组积分",d:"10个小组以月亮雅称为队名，答题推进月相演化，从新月长成满月。",lk:"https://wwdj.lanzoul.com/iHCVt48z551e",lv:"/r/moon-phases/",sh:"assets/moon-phases.png",pp:"p01"},
{id:"t03",t:"呼吸系统小组积分",g:"三年级下",u:"人体",ty:"小组积分",d:"呼吸系统主题小组积分，答对题目点亮肺部，适合单元复习课。",lk:"https://wwdj.lanzoul.com/iPJ6E48z54sf"},
{id:"t04",t:"消化系统积分",g:"三年级下",u:"人体",ty:"小组积分",d:"消化系统主题小组积分，食物一路通关推进积分进度。",lk:"https://wwdj.lanzoul.com/iOdsI48z550d"},
{id:"t05",t:"冷与热小组积分系统",g:"四年级上",u:"冷与热",ty:"小组积分",d:"冷与热主题小组积分，温度计式进度条直观展示各组表现。",lk:"https://wwdj.lanzoul.com/i09z348z54wi"},
{id:"t06",t:"哺乳动物积分系统",g:"三年级上",u:"动物分类",ty:"小组积分",d:"哺乳动物主题小组积分，集齐动物卡片式激励。",lk:"https://wwdj.lanzoul.com/iANm48z53vc"},
{id:"t07",t:"磁悬浮积分系统",g:"三年级下",u:"磁针与南北",ty:"小组积分",d:"磁悬浮主题小组积分，积分越高悬浮越高，科技感拉满。",lk:"https://wwdj.lanzoul.com/iCXwR48z53xe"},
{id:"t08",t:"运动积分系统",g:"三年级下",u:"第五单元",ty:"小组积分",d:"运动会竞速主题，10组赛道积分排名，课堂气氛神器。",lk:"https://wwdj.lanzoul.com/i4Iv048z53zg",lv:"/r/sports-points/",sh:"assets/sports-points.png"},
{id:"t09",t:"3年级随机点名系统",g:"三年级",ty:"随机点名",d:"按三年级名单随机点名，公平抽取，可排除已答过的同学。",lk:"https://wwdj.lanzoul.com/isYr48z54kh"},
{id:"t10",t:"AI随机点名系统",g:"通用",ty:"随机点名",d:"智能随机点名，兼顾均衡性，避免重复点名。",lk:"https://wwdj.lanzoul.com/iTr9t48z54na"},
{id:"t11",t:"乾坤随机点名",g:"通用",ty:"随机点名",d:"乾坤大挪移式转盘点名，仪式感十足的课堂抽取。",lk:"https://wwdj.lanzoul.com/iWFR48z54qd"},
{id:"t12",t:"磁极随机点名",g:"三年级下",u:"磁针与南北",ty:"随机点名",d:"磁极主题随机点名，与磁铁单元课堂情境无缝衔接。",lk:"https://wwdj.lanzoul.com/iFaJu48z54pc"},
{id:"t13",t:"3D粒子点名系统",g:"通用",ty:"随机点名",d:"3D粒子特效随机点名，视觉冲击力强，瞬间抓住全班注意力。",lk:"https://wwdj.lanzoul.com/i8JCm48z544b"},
{id:"t14",t:"计时器",g:"通用",ty:"计时工具",d:"大字号课堂计时器，全屏显示，活动环节时间管理。",lk:"https://wwdj.lanzoul.com/ijSsH48z54ah"},
{id:"t15",t:"倒计时器",g:"通用",ty:"计时工具",d:"分钟/秒钟双模式倒计时，最后3/5/10秒变色提醒，投屏清晰醒目。",lk:"https://wwdj.lanzoul.com/igBZ848z545c",lv:"/r/countdown-timer/",sh:"assets/countdown-timer.png"},
{id:"t16",t:"智慧课堂 专注森林",g:"通用",ty:"专注工具",d:"自习作业时间专注计时，坚持专注让森林成长。",lk:"https://wwdj.lanzoul.com/ioDui48z54bi"},
{id:"t17",t:"专注森林（升级版）",g:"通用",ty:"专注工具",d:"增强版专注训练，多种树木与成长动画，课堂常规养成工具。",lk:"https://wwdj.lanzoul.com/iTHDX48z54fc"},
];

/* ---------- 提示词 10（原7条保留 + 上传3条） ---------- */
const PROMPTS = [
{id:"p01",t:"月相积分系统提示词",plat:"DeepSeek",pr:"t02",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"月相积分系统"。

（限制）功能：页面显示10个小组积分卡片，队名用月亮的雅称：婵娟、玉盘、冰轮、玉弓、银钩、玉兔、金蟾、望舒、太阴、桂轮。每个组从新月🌑开始，老师点"+1月相演化"按钮，月相就按 🌑→🌒→🌓→🌔→🌕 的顺序长一格，最先变成满月🌕的组获胜。

交互：每张卡片上有大月相图和演化进度小图标；顶部有"重置全部"和"随机+1"按钮，方便课堂快节奏使用。

风格要像深空星空：紫色渐变背景、金色月亮，卡片圆角，可爱、教学感强。

月相用CSS/SVG绘制或使用emoji，要求清晰可辨。

（限制）输出格式为html。`},
{id:"p02",t:"听话的小车提示词",plat:"DeepSeek",pr:"w05",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"听话的小车"。

（限制）功能：打开摄像头，识别我的手是张开还是握拳。页面上有一辆画着N和S极的小车（红色和蓝色）。如果我张开手，小车就像被磁铁吸住一样，慢慢朝我的手移动。如果我握拳，小车就被推开，往相反方向跑。

交互：右下角有三个按钮。揭秘模式：打开后，手的位置会出现一块红蓝"隐形磁铁"，而且握拳的时候磁铁的N/S会翻转，这样就能解释小车为什么动。指南：滑出一个面板，告诉用户怎么玩。

AR模式：打开后背景摄像头画面变清楚，就像真的在现实里操作。

风格要可爱、教学感强，背景半透明能看到摄像头。

小车用SVG（可缩放矢量图）画。直接用MediaPipe（手势识别）。

（限制）输出格式为html。`},
{id:"p03",t:"3D消化系统提示词",plat:"DeepSeek",pr:"w06",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"消化系统实验室"。

（限制）功能：用Three.js做3D拼装。左侧散落8个消化道器官（口腔、食道、胃、肝脏、胰腺、小肠、大肠、肛门），中间是一个半透明线框人体轮廓。点击或拖动器官到正确位置，距离合适时强磁吸自动归位，并弹出该器官的功能介绍卡片。

交互：右侧控制面板：实验目标、"演示消化过程"按钮（一颗食物从口腔走到肛门）、难度切换（简单：线框常显；困难：隐藏轮廓）、重置实验、重置视角；底部显示拼装进度（0/8）。操作提示：鼠标拖动旋转、滚轮缩放、强磁吸轻松拼装。

风格：浅色实验室风格，器官用可爱的卡通材质配色，教学感强。

（限制）输出格式为html。`},
{id:"p04",t:"太阳与地球两种学说提示词",plat:"DeepSeek",pr:"w16",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"太阳与地球——两种学说"。

（限制）功能：页面并排放两个可切换的天体运行模型。地心说（托勒密）：地球在中心，太阳、月亮和行星绕地球转，行星带本轮结构。日心说（哥白尼）：太阳在中心，地球等行星绕太阳转。有"播放/暂停"按钮让模型运转，直观对比两种学说对同一片星空的不同解释。

交互：顶部按钮切换两种学说；下方简介面板介绍两种学说的历史，以及后来支持日心说的证据（金星相位、木星卫星）。

风格：古典科学史插画感，深蓝星空底色，字体有年代感但不失可爱。

（限制）输出格式为html。`},
{id:"p05",t:"课堂倒计时器提示词",plat:"DeepSeek",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"课堂倒计时器"。

（限制）功能：分钟、秒钟两组数字可点击加减设置；"开始/暂停"和"重置"按钮；时间到响铃提醒。最后3秒、5秒、10秒分别用不同颜色闪烁提醒，教室最后一排也能看清。

风格：大字号居中显示，深色背景高对比，适合教室投屏。

（限制）输出格式为html。`},
{id:"p06",t:"随机点名系统提示词",plat:"DeepSeek",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"随机点名系统"。

（限制）功能：内置可编辑的学生名单；点"开始"名字快速滚动，点"停止"随机抽中一位；可勾选"排除已答过的同学"，避免重复点名。

交互：抽中后大字展示并配音效；支持一键重置名单，方便下一个班使用。

风格：热闹但不花哨，适合课堂投屏。

（限制）输出格式为html。`},
{id:"p07",t:"专注森林提示词",plat:"DeepSeek",txt:`用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）

（明确）主题：我是一名小学科学教师，请你帮我做一个网页，叫"专注森林"。

（限制）功能：设定专注时长后开始计时，专注期间种下一棵树苗，坚持到结束树就长大；中途放弃树会枯萎。全班多棵树共同长成一片森林，累计专注时长可见。

风格：清新绿色系，树木成长有简单动画，治愈、激励感强。

（限制）输出格式为html。`},
{id:"p08",t:"命题设计提示词",plat:"DeepSeek",txt:`你是一名深谙小学科学命题标准的资深教研员。请严格依据以下要求，为命制一份三年级下学期的科学期末试卷。

【命题根本依据】
1.命题必须严格参照三年级下册科学教材，试题考察内容来源于教材，所有考点不得超出教材。
2.试题类型、难度和格式需参考往年的科学期末试卷，保持整体命题风格的一致性。

【考试范围与内容】
1.范围：三年级下册科学教材的全部单元。考试范围不得参考往年试卷。
2.内容：覆盖范围内所有知识点，同一知识点不得重复考查。

【命题难度与风格】
1.难度：简单或适中，符合三年级学生认知水平。
2.风格：题目必须植入真实情境，且情境是解题的必要信息载体而非可有可无的装饰。具体要求：①情境应简洁明了，一句话即可创设完整的问题场域，符合三年级学生的阅读水平。②填空题和选择题的题干应让学生"从情境中发现问题"，而非直接提问概念定义。③凡出现图片的题目，必须是"读图答题"型，即学生必须观察图片才能作答，图片不能仅作装饰。

【试卷整体结构】
1.满分100分。
2.卷内结构：一、填空题；二、选择题；三、连线题/画图题/看图填空题；四、阅读分析题；五、实验探究题，分为三道实验探究小题。第三大题、第四大题、以及第五大题的三道实验探究小题都不得重复考察同一教学单元。
3.试卷末尾附上"考点细目表"，包含题型、题号、能力层级目标（识记、理解、应用、分析、评价、创造）、难度预估（易、中、难）、核心素养（科学观念、科学思维、探究实践、态度责任）、教材对应单元/课、核心考点及分值。
4.试卷格式要求：①一级标题（一、二、三…）：宋体、四号、加粗。②二级标题（（一）、（二）、（三）…）：宋体、四号，用于实验探究题的内部小题。③正文：宋体、小四。④段落间距：单倍行距。⑤无得分栏，大题标题（如"一、填空题"）下方不要有横线或下划线，不要有任何灰色背景、蓝色边框或其他彩色装饰。

【各题型具体要求】
一、填空题（共8-10分，每空1分）
考查逻辑：在科学现象或实验结论的语境中，精准再现科学核心术语、概念或数据。
命题技巧：1.每空答案控制在1-4个字，确保答案的唯一性。2.可针对学生易混、易错字设计题目。3.可提供实验装置图或现象图，考查对结构名称或原理的识记。4.可考查实验仪器的读数。

二、选择题（共40分，每题2分，共20题）
考查逻辑：通过创设具体情境，精准辨析易混淆的科学概念。不得考察填空题已考察的知识点。
命题技巧：1.将知识点融入生活习惯、成语谚语、游戏活动等真实场景中。2.每个干扰项都需源自学生的典型错误概念或常见误区，使题目具有诊断功能。3.能力考察点包括：概念辨析、图表信息提取、简单推理预测、选择最优实验方案等。

三、连线题/画图题/看图填空题（共6-10分）
考查逻辑：建立科学概念间的逻辑关联，或对结构类知识进行精准识记。不得考察填空题、选择题已考察的知识点。
命题要求：1.连线题：明确要求"请使用铅笔和直尺连线"，考查"结构与功能"、"概念与实例"的匹配关系。2.看图填空题：提供带有多条引线的结构示意图，在引线末端设置填空题，让学生填写各部分的科学名称。图为答题的必要依据，引线序号应与填空序号一一对应。3.若为画图题，需提供明确的作图要求和必要的参考信息。

四、阅读分析题（共8-10分，每空2分，3-5个小题）
考查逻辑：模拟"阅读提取-推理-应用"的完整信息加工闭环。不得考察连线题已考察的教学单元。
材料要求：提供200-400字的科学短文或数据图表，需图文结合。
设问必须体现以下三个递进层次，且每一问都应与材料直接相关：1.信息提取：答案可直接从文中找到，材料中必须有明确的关键信息支撑。2.分析推理：需结合文中信息与已有知识进行逻辑判断，题干应暗示学生"回到文章中去寻找依据"。3.迁移评价：设计开放性问题，必须基于材料中的原理或方法进行延伸，材料应暗含答题的方向性提示。
题型可混合包含：填空题(1-2个)、选择题(1-2个)、简答题(0-1个，非必要不命题)。

五、实验探究题（共24-35分，由3个独立探究小题组成，每题10-12分）
总体要求：
1.通常从三个教学单元中各选取一个核心实验作为命题素材，每题对应一个单元。不得考察连线题、阅读分析题已考察的教学单元。
2.根据教材内容，可全部命制为"实验探究型"，也可选用"模拟实验型"或"工程实践型"，但必须确保所选类型在教材中有对应的探究活动。
3.逻辑递进要求：每道题的各个小问必须构成一个完整的探究故事线：问题→方法→现象→结论→应用。严禁在探究过程中插入与主线无关的知识点。同一探究问题内，前一个问题的基础应成为后一个问题的铺垫。
4.每道题可设置：2-3个选择题、2-3个填空题。三道探究小题一共最多出现1个简答题（比如：建议、计划等）。
5.三种可选类型及命题逻辑链：
类型一：实验探究型（最常用）。逻辑链：实验装置图/操作描述→数据记录表/现象图→分析数据现象→归纳结论→迁移应用。核心命题点：识别变量（保持不变的条件是什么？改变的条件是什么？）；从图表中读取数据，描述现象，或指出异常的记录数据；基于证据，得出实验结论；将结论应用于解释新的生活情境或解决新问题（非简答题）。
类型二：模拟实验型。逻辑链：模型装置图→识别模拟对象→预测操作后的现象→推测模型原理→联系实际应用。核心命题点：填空/选择：将模型各组件与真实对象一一对应；选择：预测对模型进行操作后将如何变化；简答/选择：推测模型变化所模拟的真实科学原理；迁移：将该原理应用于生活（非简答题）。
类型三：工程实践型。逻辑链：工程任务需求→设计图或材料清单→测试数据→评估效果→优化改进。核心命题点：选择：为实现特定功能选择最合适的材料；数据分析：根据测试数据评价作品性能；评价与改进：指出设计或制作中的不足，并提出具体的改进方案（非简答题）。

【输出格式要求】
以HTML格式输出完整试卷，提供可以下载word的按钮。`},
{id:"p09",t:"教学设计Skill提示词",plat:"Agent · Skill",txt:`请帮我创建一个"基于标准的教学评一致性教学设计Skill"。

1.收集信息：课题、年级、教材版本（大象版+上下册+出版年份）、单元、课时、设计者。
2.查课标：常规单元先查《教学指南》单元末"核心概念进阶关系图"，只取图中黑体条目；体验/反思单元直接查课标原文。内容要求与学业要求在课标中位置相邻，一一对应。学业要求必须一字不落抄写，禁止归纳改写。两者各自标页码"（第 X 页）"。课标原文括号用半角()，我们自注页码用全角（）。
3.生成内容，结构固定：
标题《课题》教学设计（居中）
基本信息行（不缩进）：教材来源/内容来源/主题/教学课时/授课对象/设计者
【课标要求】(一)学习内容及内容要求 (二)学业要求
【教材分析】2段，每段150-220字
【学情分析】2段，每段130-180字，必须含困难预判与应对
【教学目标】4条：科学观念/科学思维/探究实践/态度责任
【评价任务】4条，与目标一一对应，以"能""会"开头
【教学重点与难点】
【学习过程】四列表格：教学环节|教师活动|学生活动|评价与说明
课前准备
板书设计（1-3行，每行居中）
教学反思（2段，每段150-200字）
4.【硬规则】
- 目标行为动词取自课标附录2，可观察可测评。
- "评价与说明"列只写设计意图（40-60字）。
- 教学环节按教材板块：问题与假设→探究与结论→反思与评价→应用与迁移；"应用与迁移"几乎每课都有必须保留且用教材原文情境；"阅读与新知"教材才有写。
- 全文正文用中文引号""，禁止英文直引号（HTML属性除外）。
- 表格列宽固定：教学环节2.2cm/教师活动5.5cm/学生活动3.5cm/评价与说明3.3cm；Word里CSS宽度不生效，必须给每个th/td注入内联pt宽度（62.4/155.9/99.2/93.6pt）。
- 排版：标题宋体二号加粗居中；一级标题【】黑体三号；二级标题(一)(二)楷体三号；正文仿宋四号；表格仿宋小四；段前段后0，行距1.15。`},
{id:"p10",t:"教学网页Skill提示词",plat:"Agent · Skill",txt:`大致思路：
1.已有成果（教学网页）分类，比如按照科学四大领域分类。
2.教学网页分析，分析布局、交互方式、背景风格等维度。
3.教学网页匹配，与科学教材进行匹配。
4.生成教学网页skill，明确课题匹配网页类别的优先级。

提醒：在用skill批量生成之前，先让Agent判定课题的网页类别。

——极简版教学网页Skill提示词——
做一个固定一屏、左右分栏/全屏3D/多模块导航、触摸友好、支持拖拽拼装、点击拖动、手势识别、画线连线、拖拽填空、转盘随机、时间轴播放、全屏演示、音效语音、庆祝动画、毛玻璃卡片、圆角柔和阴影、字体10～80px、无下拉菜单的中文小学科学/语文互动教学网页模板。`},
];

/* ---------- 教学设计 9（大象版AI生成教案合集，对齐Excel） ---------- */
const DESIGNS = [
{id:"d01",t:"大象版一年级上册教学设计",g:"一年级上册",lk:"https://wwdj.lanzoul.com/b00g4hqpva",cd:"279x",d:"大象版科学一年级上册 AI 生成教学设计（教案）合集。"},
{id:"d02",t:"大象版二年级上册教学设计",g:"二年级上册",lk:"https://wwdj.lanzoul.com/b00g4hqpwb",cd:"50d9",d:"大象版科学二年级上册 AI 生成教学设计（教案）合集。"},
{id:"d03",t:"大象版三年级上册教学设计",g:"三年级上册",lk:"https://wwdj.lanzoul.com/b00g4hqoxc",cd:"audk",d:"大象版科学三年级上册 AI 生成教学设计（教案）合集。"},
{id:"d04",t:"大象版四年级上册教学设计",g:"四年级上册",lk:"https://wwdj.lanzoul.com/b00g4hqp0f",cd:"ct66",d:"大象版科学四年级上册 AI 生成教学设计（教案）合集。"},
{id:"d05",t:"大象版五年级上册教学设计",g:"五年级上册",lk:"https://wwdj.lanzoul.com/b00g4hqp2h",cd:"d9yr",d:"大象版科学五年级上册 AI 生成教学设计（教案）合集。"},
{id:"d06",t:"大象版六年级上册教学设计",g:"六年级上册",lk:"https://wwdj.lanzoul.com/b00g4hqp5a",cd:"bn8z",d:"大象版科学六年级上册 AI 生成教学设计（教案）合集。"},
{id:"d07",t:"大象版一年级下册教学设计",g:"一年级下册",lk:"https://wwdj.lanzoul.com/b00g4hqp7c",cd:"6fe1",d:"大象版科学一年级下册 AI 生成教学设计（教案）合集。"},
{id:"d08",t:"大象版二年级下册教学设计",g:"二年级下册",lk:"https://wwdj.lanzoul.com/b00g4hqpaf",cd:"f2mq",d:"大象版科学二年级下册 AI 生成教学设计（教案）合集。"},
{id:"d09",t:"大象版三年级下册教学设计",g:"三年级下册",lk:"https://wwdj.lanzoul.com/b00g4hqpch",cd:"5p2j",d:"大象版科学三年级下册 AI 生成教学设计（教案）合集。"},
];
/* ============================================================
   追加模块：AI实践经验 / AI教学案例 / AI科普视频 + 全站自动增强
   用法：粘贴到 data.js 最末尾，Commit 即生效，无需改 index.html
   填内容：把 lk:"TODO" 换成真实链接；img/cover 换成 assets/ 里的图片文件名
   ============================================================ */

/* ---------- AI实践经验（公众号文章 / 小红书图文） ---------- */
const ARTICLES = [
{id:"a01",t:"（示例）我是怎么用AI备出一节课的",src:"公众号文章",d:"从备课到课堂实施的完整AI备课流程记录。",lk:"TODO",img:""},
{id:"a02",t:"（示例）AI出题工作流",src:"小红书图文",d:"用AI按考点细目表命制期末卷的步骤拆解。",lk:"TODO",img:""},
];

/* ---------- AI教学案例（公开课课件 / 录课视频） ----------
   lk = 课件下载链接（走下载码验证）；lk2 = 录课视频在线观看链接 */
const CASES = [
{id:"c01",t:"（示例）《给水加热》公开课",ty:"公开课课件",g:"四年级上",d:"课件+教学设计+课堂实录，AI互动实验网页融入课堂的完整案例。",lk:"TODO",lk2:""},
{id:"c02",t:"（示例）《磁铁有两极》录课",ty:"录课视频",g:"三年级下",d:"课堂实录，含AR听话的小车互动环节。",lk:"TODO",lk2:""},
];

/* ---------- AI科普视频（B站 / 公众号） ---------- */
const VIDEOS = [
{id:"v01",t:"（示例）AI三分钟讲透光的色散",src:"B站",d:"AI生成的动画演示牛顿三棱镜实验。",lk:"TODO",cover:""},
{id:"v02",t:"（示例）太阳系有多大？",src:"公众号",d:"AI辅助制作的科普短视频。",lk:"TODO",cover:""},
];

/* ---------- 以下为自动增强程序，勿动 ---------- */
(function(){
function ready(fn){if(document.readyState!=='loading'){fn()}else{document.addEventListener('DOMContentLoaded',fn)}}
ready(function(){
var E=window.esc||function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})};
var IT=window.isTodo||function(v){return !v||String(v).indexOf('TODO')>-1};

if(!document.getElementById('kwExtraStyle')){
var st=document.createElement('style');st.id='kwExtraStyle';
st.textContent=['.qr img{width:auto!important;height:auto!important;max-width:100%!important;max-height:100%!important;object-fit:contain}',
'.ir1{animation-delay:.95s}.ir2{animation-delay:1.02s}.ir3{animation-delay:1.09s}.ir4{animation-delay:1.16s}.ir5{animation-delay:1.23s}.ir6{animation-delay:1.3s}',
'.hero-cards{display:grid;grid-template-columns:repeat(3,1fr);gap:.8rem;margin-top:1.4rem}',
'.hcard{display:flex;flex-direction:column;gap:.45rem;border:1px solid var(--line);background:var(--paper);padding:.7rem .8rem;text-decoration:none;color:var(--ink);transition:.25s}',
'.hcard:hover{border-color:var(--accent);box-shadow:7px 7px 0 rgba(200,69,27,.14);transform:translateY(-3px)}',
'.hcard b{font:700 13px var(--serif)}.hcard i{font:400 10px var(--mono);font-style:normal;letter-spacing:.06em;color:var(--ink2)}',
'.hc-img{display:block;height:64px;border:1px solid var(--line);overflow:hidden;background:var(--paper2)}',
'.hc-img img{width:100%;height:100%;object-fit:cover;display:block}',
'.hc-txt{display:grid;place-items:center;height:64px;padding:.4em;font:400 10.5px/1.7 var(--mono);color:var(--ink-soft);background:var(--paper2);border:1px solid var(--line);overflow:hidden;text-align:center}',
'.hc-doc{display:grid;place-items:center;height:64px;font:500 11px var(--mono);letter-spacing:.08em;color:var(--ink-soft);background:var(--paper2);border:1px solid var(--line)}',
'.kcover{width:100%;aspect-ratio:16/9;object-fit:cover;border:1px solid var(--line);background:var(--paper2);display:block;margin-bottom:.55rem}',
'.ksoon{font:400 10.5px var(--mono);color:var(--ink2);letter-spacing:.06em}',
'.kw-sub{font:500 11px var(--mono);letter-spacing:.24em;color:var(--accent);margin:1.6rem 0 .7rem}',
'.s-tab{position:fixed;right:0;top:50%;transform:translateY(-50%);z-index:55;writing-mode:vertical-rl;background:var(--ink);color:var(--paper);border:none;padding:1.35rem .78rem;font:500 12.5px var(--sans);letter-spacing:.4em;cursor:pointer;border-radius:4px 0 0 4px;transition:.25s}',
'.s-tab:hover{background:var(--accent)}',
'.s-fab{display:none;position:fixed;right:1.1rem;bottom:5.8rem;width:52px;height:52px;background:var(--accent);color:var(--paper);border:none;border-radius:50%;z-index:60;place-items:center;box-shadow:0 8px 24px rgba(25,23,20,.28);cursor:pointer;font-size:19px}',
'.s-mask{position:fixed;inset:0;background:rgba(25,23,20,.44);z-index:74;opacity:0;visibility:hidden;transition:.28s}',
'.s-mask.open{opacity:1;visibility:visible}',
'.s-panel{position:fixed;top:0;right:0;height:100vh;width:min(360px,92vw);z-index:75;background:var(--paper);border-left:1px solid var(--ink);box-shadow:-10px 0 0 rgba(25,23,20,.08);transform:translateX(103%);transition:.3s cubic-bezier(.2,.7,.2,1);display:flex;flex-direction:column}',
'.s-panel.open{transform:none}',
'.s-head{display:flex;gap:.6rem;align-items:center;padding:1rem;border-bottom:1px solid var(--line)}',
'.s-input{flex:1;min-width:0;border:1px solid var(--ink);background:transparent;padding:.7em .9em;font:500 14px var(--sans);outline:none;color:var(--ink)}',
'.s-input:focus{border-color:var(--accent)}',
'.s-close{width:34px;height:34px;background:none;border:none;font-size:24px;line-height:1;color:var(--ink2);cursor:pointer}',
'.s-close:hover{color:var(--accent)}',
'.s-body{flex:1;overflow:auto;padding:.8rem 1rem 2rem}',
'.s-empty{font:400 12px var(--mono);color:var(--ink2);line-height:2;padding:.6rem 0}',
'.s-grp{font:500 10.5px var(--mono);letter-spacing:.24em;color:var(--accent);margin:1.1rem 0 .45rem;display:flex;align-items:center;gap:.6em}',
'.s-grp i{font-style:normal;color:var(--ink2)}',
'.s-item{display:block;width:100%;text-align:left;background:var(--paper);border:1px solid var(--line);border-left:2px solid var(--line);padding:.6rem .8rem;margin-bottom:.45rem;cursor:pointer;transition:.2s;font-family:var(--sans)}',
'.s-item:hover{border-color:var(--accent);border-left-color:var(--accent);background:var(--paper2)}',
'.s-item b{display:block;font:700 13.5px var(--serif);color:var(--ink)}',
'.s-item span{display:block;margin-top:.15em;font-size:11.5px;color:var(--ink2)}',
'@media (max-width:860px){.s-tab{display:none}.s-fab{display:grid}}'].join('');
document.head.appendChild(st);
}

document.title='开物AI科学 · AI赋能教育教学全流程';
var eb=document.querySelector('.eyebrow');
if(eb&&/EST\. 2026/.test(eb.textContent))eb.textContent='KAIWU AI SCIENCE · AI FOR BASIC EDUCATION';
var lead=document.querySelector('.hero p.lead');
if(lead&&/全部家当/.test(lead.textContent))lead.textContent='我把真实备课、课堂实践和 AI 工具应用过程中产生的资源整理在这里。提示词、互动网页、教学设计，以及 AI 实践方法，免费分享给一线教师。';

function patchKw(){
document.querySelectorAll('.m-steps li b').forEach(function(b){if(b.textContent==='开物')b.textContent='下载码'});
document.querySelectorAll('.knote').forEach(function(el){if(el.innerHTML.indexOf('「开物」')>-1)el.innerHTML=el.innerHTML.replace(/「开物」/g,'「下载码」')});
}
patchKw();

var pr=document.querySelector('.prism');
if(pr&&!pr.dataset.kwFixed){
pr.dataset.kwFixed='1';
pr.setAttribute('aria-label','三棱镜色散示意图：白光经两次折射，棱镜内色散成光谱');
pr.innerHTML='<path class="pz" pathLength="1" d="M340 80 L430 380 L250 380 Z"/>'
+'<line class="beam" x1="20" y1="233" x2="294" y2="233"/>'
+'<line class="ray ir1" pathLength="1" x1="294" y1="233" x2="387" y2="237"/>'
+'<line class="ray ir2" pathLength="1" x1="294" y1="233" x2="389" y2="243"/>'
+'<line class="ray ir3" pathLength="1" x1="294" y1="233" x2="390" y2="247"/>'
+'<line class="ray ir4" pathLength="1" x1="294" y1="233" x2="392" y2="253"/>'
+'<line class="ray ir5" pathLength="1" x1="294" y1="233" x2="393" y2="257"/>'
+'<line class="ray ir6" pathLength="1" x1="294" y1="233" x2="395" y2="263"/>'
+'<line class="ray r1" pathLength="1" x1="387" y1="237" x2="620" y2="252"/>'
+'<line class="ray r2" pathLength="1" x1="389" y1="243" x2="620" y2="266"/>'
+'<line class="ray r3" pathLength="1" x1="390" y1="247" x2="620" y2="278"/>'
+'<line class="ray r4" pathLength="1" x1="392" y1="253" x2="620" y2="290"/>'
+'<line class="ray r5" pathLength="1" x1="393" y1="257" x2="620" y2="301"/>'
+'<line class="ray r6" pathLength="1" x1="395" y1="263" x2="620" y2="313"/>'
+'<text class="lab" x="20" y="214">AI · 输入</text>'
+'<text class="lab" x="478" y="352">科学课 · 输出</text>'
+'<text class="lab fig" x="20" y="442">FIG.01 — 光的色散：棱镜内色散，出射再折射 / DISPERSION</text>';
}

var pb=document.querySelector('.prism-box');
if(pb&&!pb.querySelector('.hero-cards')){
var hc=document.createElement('div');hc.className='hero-cards';
hc.innerHTML='<a class="hcard" href="#/labs"><span class="hc-img"><img src="assets/water-heating.png" alt="互动网页示例"></span><b>互动网页</b><i>在线使用 · 投屏即用</i></a>'
+'<a class="hcard" href="#/prompts"><span class="hc-txt">用deepseek生成<br>课堂互动网页…</span><b>提示词库</b><i>复制即用 · 覆盖全流程</i></a>'
+'<a class="hcard" href="#/design"><span class="hc-doc">教案.docx</span><b>教学设计</b><i>新课标 · 教学评一致</i></a>';
pb.appendChild(hc);
}

var nav=document.querySelector('header nav');
if(nav&&!nav.querySelector('a[href="#/practice"]')){
var ab=nav.querySelector('a[href="#/about"]');
['practice|AI实践','cases|教学案例','videos|科普视频'].forEach(function(pair){
var pp=pair.split('|'),a=document.createElement('a');
a.href='#/'+pp[0];a.textContent=pp[1];a.setAttribute('data-kwroute',pp[0]);
nav.insertBefore(a,ab);
});
}

var note=document.querySelector('.index-note');
if(note&&!document.querySelector('.row[data-route="#/practice"]')){
var EMS={practice:'AI PRACTICE',cases:'AI CASES',videos:'AI VIDEOS'};
function mkRow(no,name,route,cid,unit,desc){
var art=document.createElement('article');art.className='row reveal';
art.setAttribute('data-name',name);art.setAttribute('data-route',route);
art.innerHTML='<span class="row-no">'+no+'</span><div class="row-main"><h3>'+name+'<em>'+EMS[route.slice(2)]+'</em></h3><p>'+desc+'</p></div><div class="row-meta"><strong id="'+cid+'">0</strong><span>'+unit+'</span></div><a class="row-go" href="'+route+'">查看全部 <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M2 7h9M8 3.5L11.5 7 8 10.5" stroke="currentColor" stroke-width="1.5"/></svg></a>';
note.parentNode.insertBefore(art,note);
art.addEventListener('click',function(){location.hash=route});
if(window.io)io.observe(art);
}
mkRow('04','AI实践经验','#/practice','cntPractice','篇','公众号文章与小红书图文：我的AI备课方法、工具用法与踩坑记录。');
mkRow('05','AI教学案例','#/cases','cntCases','例','公开课课件与录课视频：AI融入真实课堂的完整案例。');
mkRow('06','AI科普视频','#/videos','cntVideos','个','公众号与B站的AI科普视频，课堂导入或课后拓展都能用。');
var cP=document.getElementById('cntPractice');if(cP)cP.textContent=ARTICLES.length;
var cC=document.getElementById('cntCases');if(cC)cC.textContent=CASES.length;
var cV=document.getElementById('cntVideos');if(cV)cV.textContent=VIDEOS.length;
var sR=document.getElementById('stRes');if(sR)sR.textContent=(PROMPTS.length+WEB_ITEMS.length+TOOLS.length+DESIGNS.length+ARTICLES.length+CASES.length+VIDEOS.length).toLocaleString('zh-Hans-CN');
if(/家当|在线预览|教材单元/.test(note.textContent))note.textContent='// 支持按「年级 × 类型 × 学科核心素养」交叉筛选，全站资源可在右侧搜索栏直达。';
}

var mainEl=document.querySelector('main.wrap')||document.querySelector('main');
var cont=document.createElement('div');cont.id='kwPage';cont.hidden=true;
if(mainEl)mainEl.appendChild(cont);else document.body.appendChild(cont);

function kgrp(title,arr,kind){
var h='<p class="kw-sub">'+title+' · '+arr.length+'</p><div class="kgrid">';
if(!arr.length)h+='<p class="s-empty">内容整理中——我会陆续把这一类填进来。</p>';
arr.forEach(function(o){
var tags=[o.g,o.ty,o.src].filter(Boolean).map(function(v){return '<span class="ktag">'+E(v)+'</span>'}).join('');
var cov=(o.img||o.cover)?'<img class="kcover" src="'+E(o.img||o.cover)+'" alt="'+E(o.t)+'">':'';
var btns='';
if(kind==='a'){btns=IT(o.lk)?'<span class="ksoon">链接整理中</span>':'<a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">↗ 阅读原文</a>';}
else if(kind==='c'){btns=(IT(o.lk)?'<span class="ksoon">课件整理中</span>':'<button class="kbtn" data-kwdl="'+E(o.lk)+'" data-kwname="'+E(o.t)+'">⬇ 下载课件</button>')+((o.lk2&&!IT(o.lk2))?'<a class="kbtn ghost" target="_blank" rel="noopener" href="'+E(o.lk2)+'">▶ 观看录课</a>':'');}
else{btns=IT(o.lk)?'<span class="ksoon">视频整理中</span>':'<a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">▶ 去观看</a>';}
h+='<div class="kcard" style="cursor:default">'+cov+'<h3>'+E(o.t)+'</h3><div class="ktags">'+tags+'</div><p class="kdesc">'+E(o.d)+'</p><div style="margin-top:.4rem">'+btns+'</div></div>';
});
return h+'</div>';
}
function renderKW(key){
var T={practice:['05 — AI PRACTICE','AI 实践经验','公众号文章与小红书图文：AI备课的方法、工具与踩坑记录。'],cases:['06 — AI CASES','AI 教学案例','公开课课件与录课视频，看AI如何真实融入一堂课。'],videos:['07 — AI VIDEOS','AI 科普视频','公众号与B站的AI科普视频，课堂导入或课后拓展都能用。']}[key];
var B={practice:kgrp('公众号文章',ARTICLES.filter(function(o){return o.src==='公众号文章'}),'a')+kgrp('小红书图文',ARTICLES.filter(function(o){return o.src==='小红书图文'}),'a'),
cases:kgrp('公开课课件',CASES.filter(function(o){return o.ty==='公开课课件'}),'c')+kgrp('录课视频',CASES.filter(function(o){return o.ty==='录课视频'}),'c'),
videos:kgrp('B站',VIDEOS.filter(function(o){return o.src==='B站'}),'v')+kgrp('公众号',VIDEOS.filter(function(o){return o.src==='公众号'}),'v')}[key];
cont.innerHTML='<section><div class="sec-head" style="margin-top:2.6rem"><span class="sec-no">'+T[0]+'</span><h2>'+T[1]+'</h2><p>'+T[2]+'</p></div>'+B+'</section>';
cont.querySelectorAll('[data-kwdl]').forEach(function(b){
b.addEventListener('click',function(){if(window.kwGate)kwGate(b.getAttribute('data-kwname'),b.getAttribute('data-kwdl'))});
});
patchKw();
}
function kwRoute(){
var h=location.hash||'#/home';
var m=h.match(/^#\/(practice|cases|videos)/);
if(m){
var hr=document.getElementById('homeRoot'),pg=document.getElementById('pageRoot');
if(hr)hr.style.display='none';
if(pg)pg.hidden=true;
document.querySelectorAll('nav a[data-route]').forEach(function(a){a.classList.remove('active')});
document.querySelectorAll('nav a[data-kwroute]').forEach(function(a){a.classList.toggle('active',a.getAttribute('data-kwroute')===m[1])});
renderKW(m[1]);
cont.hidden=false;
window.scrollTo(0,0);
}else{cont.hidden=true;}
}
window.addEventListener('hashchange',kwRoute);
kwRoute();

if(!document.getElementById('sTab')){
function mk(id,cls,txt){var b=document.createElement('button');b.id=id;b.className=cls;b.type='button';b.textContent=txt;document.body.appendChild(b);return b}
var sTab=mk('sTab','s-tab','搜索资源'),sFab=mk('sFab','s-fab','🔍');
var sMask=document.createElement('div');sMask.id='sMask';sMask.className='s-mask';document.body.appendChild(sMask);
var sPanel=document.createElement('aside');sPanel.id='sPanel';sPanel.className='s-panel';
sPanel.innerHTML='<div class="s-head"><input class="s-input" id="sInput" placeholder="搜索：课题 / 单元 / 素养 / 提示词…" autocomplete="off"><button class="s-close" id="sClose" type="button" aria-label="关闭">×</button></div><div class="s-body" id="sBody"><p class="s-empty">输入关键词，一次搜索全部资源：<br>网页 · 工具 · 提示词 · 教学设计 · 文章 · 案例 · 视频</p></div>';
document.body.appendChild(sPanel);
var sInput=sPanel.querySelector('#sInput'),sBody=sPanel.querySelector('#sBody');
var EMPTY=sBody.innerHTML;
function openS(){sMask.classList.add('open');sPanel.classList.add('open');setTimeout(function(){sInput.focus()},250)}
function closeS(){sMask.classList.remove('open');sPanel.classList.remove('open')}
sTab.addEventListener('click',openS);sFab.addEventListener('click',openS);
sPanel.querySelector('#sClose').addEventListener('click',closeS);
sMask.addEventListener('click',closeS);
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeS()});
function jump(kind,id){
closeS();
if(kind==='a'||kind==='c'||kind==='v'){location.hash='#/'+({a:'practice',c:'cases',v:'videos'})[kind];return}
if(kind==='p'){window.pending=id;if(location.hash==='#/prompts'){if(window.route)route()}else{location.hash='#/prompts'}return}
if(kind==='d'){location.hash='#/design';return}
if(window.FL){FL.tab=(kind==='w')?'web':'tools';}
window.pending=id;
if(location.hash==='#/labs'){if(window.route)route()}else{location.hash='#/labs'}
}
sInput.addEventListener('input',function(){
var q=sInput.value.trim().toLowerCase();
if(!q){sBody.innerHTML=EMPTY;return}
function hit(o){return (String(o.t)+' '+(o.d||'')+' '+(o.u||'')+' '+(o.c||'')+' '+(o.g||'')+' '+(o.ty||'')+' '+(o.src||'')).toLowerCase().indexOf(q)>-1}
var out='';
function grp(title,items,kind,meta){
if(!items.length)return;
out+='<p class="s-grp">'+title+'<i>'+items.length+'</i></p>'+items.map(function(o){
return '<button class="s-item" data-k="'+kind+'" data-id="'+o.id+'"><b>'+E(o.t)+'</b><span>'+E(meta(o)||'')+'</span></button>';
}).join('');
}
grp('教学网页',WEB_ITEMS.filter(hit),'w',function(o){return [o.g,o.ty].filter(Boolean).join(' · ')});
grp('课堂工具',TOOLS.filter(hit),'t',function(o){return [o.g,o.ty].filter(Boolean).join(' · ')});
grp('提示词',PROMPTS.filter(function(o){return (String(o.t)+' '+(o.txt||'')).toLowerCase().indexOf(q)>-1}),'p',function(o){return o.plat||''});
grp('教学设计',DESIGNS.filter(hit),'d',function(o){return o.g||''});
grp('AI实践',ARTICLES.filter(hit),'a',function(o){return o.src||''});
grp('教学案例',CASES.filter(hit),'c',function(o){return [o.g,o.ty].filter(Boolean).join(' · ')});
grp('科普视频',VIDEOS.filter(hit),'v',function(o){return o.src||''});
sBody.innerHTML=out||'<p class="s-empty">没有找到与「'+E(sInput.value.trim())+'」相关的资源</p>';
});
sBody.addEventListener('click',function(e){var b=e.target.closest('.s-item');if(b)jump(b.getAttribute('data-k'),b.getAttribute('data-id'))});
sInput.addEventListener('keydown',function(e){if(e.key==='Enter'){var f=sBody.querySelector('.s-item');if(f)jump(f.getAttribute('data-k'),f.getAttribute('data-id'))}});
}
});
})();
/* ---------- 追加补丁 v2：棱镜内光线显色 / 二维码限尺寸 / 删导航公众号 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
var st=document.createElement('style');st.id='kwPatch2';
st.textContent=[
/* 棱镜内部六色光线：白光入射后，在玻璃内部就开始色散 */
'.ir1{stroke:#C0392B;opacity:.92;stroke-dasharray:none!important}',
'.ir2{stroke:#D68227;opacity:.92;stroke-dasharray:none!important}',
'.ir3{stroke:#C9A227;opacity:.92;stroke-dasharray:none!important}',
'.ir4{stroke:#3F7A3A;opacity:.92;stroke-dasharray:none!important}',
'.ir5{stroke:#2C6E9B;opacity:.92;stroke-dasharray:none!important}',
'.ir6{stroke:#5B4E9E;opacity:.92;stroke-dasharray:none!important}',
/* 二维码限死尺寸：首页关于 / 弹窗 / 公众号页 都不超过150px */
'.qr img{width:auto!important;height:auto!important;max-width:150px!important;max-height:150px!important;object-fit:contain!important}'
].join('');
document.head.appendChild(st);
/* 删除导航栏里的"公众号" */
document.querySelectorAll('header nav a').forEach(function(a){
if(a.textContent.replace(/\s/g,'')==='公众号')a.remove();
});
});
})();
/* ---------- 追加补丁 v3：色散图几何修正 / 二维码强制限宽 / 关于页重写 / 编号与文案 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
/* 0) 样式：关于页三版块 */
var st=document.createElement('style');st.id='kwPatch3';
st.textContent=[
'.ab-grid{display:grid;grid-template-columns:1.55fr .85fr;gap:2.4rem;margin-top:1.8rem}',
'.ab-intro{font:400 14.5px/2 var(--sans);color:var(--ink-soft);max-width:640px;margin:.9rem 0 0}',
'.ablock{border-top:2px solid var(--ink);padding:1.1rem 0 1.5rem;margin-top:1.4rem}',
'.ab-no{font:500 10px var(--mono);letter-spacing:.24em;color:var(--accent);display:block;margin-bottom:.45rem}',
'.ablock h4{font:700 17px var(--serif);margin:0 0 .6rem;color:var(--ink)}',
'.ablock p{font:400 14px/2.05 var(--sans);color:var(--ink-soft);margin:0}',
'.ab-side{padding-top:1.8rem}',
'.ab-qrcard{border:1px solid var(--line);background:var(--paper);padding:1.2rem;display:flex;flex-direction:column;gap:.7rem;align-items:flex-start;position:sticky;top:90px}',
'.ab-qrcard b{font:700 14px var(--serif);color:var(--ink)}',
'.ab-qrcard span{font:400 12px/1.9 var(--sans);color:var(--ink2)}',
'@media(max-width:860px){.ab-grid{grid-template-columns:1fr}.ab-side{padding-top:0}}'
].join('');
document.head.appendChild(st);

/* 1) 色散图：白光斜入射 → 棱镜内散开(红上紫下) → 出射面向下展开 */
var pr=document.querySelector('.prism');
if(pr){
pr.setAttribute('aria-label','三棱镜色散示意图：白光斜入射，棱镜内色散，出射展开');
pr.innerHTML=
'<path class="pz" pathLength="1" d="M340 80 L430 380 L250 380 Z"/>'+
'<line class="beam" x1="20" y1="272" x2="294" y2="233"/>'+
'<line class="ray ir1" pathLength="1" x1="294" y1="233" x2="388" y2="244"/>'+
'<line class="ray ir2" pathLength="1" x1="294" y1="233" x2="389" y2="249"/>'+
'<line class="ray ir3" pathLength="1" x1="294" y1="233" x2="390" y2="254"/>'+
'<line class="ray ir4" pathLength="1" x1="294" y1="233" x2="392" y2="259"/>'+
'<line class="ray ir5" pathLength="1" x1="294" y1="233" x2="393" y2="264"/>'+
'<line class="ray ir6" pathLength="1" x1="294" y1="233" x2="395" y2="271"/>'+
'<line class="ray r1" pathLength="1" x1="388" y1="244" x2="620" y2="288"/>'+
'<line class="ray r2" pathLength="1" x1="389" y1="249" x2="620" y2="302"/>'+
'<line class="ray r3" pathLength="1" x1="390" y1="254" x2="620" y2="316"/>'+
'<line class="ray r4" pathLength="1" x1="392" y1="259" x2="620" y2="331"/>'+
'<line class="ray r5" pathLength="1" x1="393" y1="264" x2="620" y2="347"/>'+
'<line class="ray r6" pathLength="1" x1="395" y1="271" x2="620" y2="365"/>'+
'<text class="lab" x="20" y="252">AI · 输入</text>'+
'<text class="lab" x="470" y="392">科学课 · 输出</text>'+
'<text class="lab fig" x="20" y="442">FIG.01 — 光的色散：棱镜内色散，出射展开 / DISPERSION</text>';
}

/* 2) 二维码：不再依赖CSS类，直接内联锁死150px（弹窗/关于/公众号全适用） */
function fixQr(){
document.querySelectorAll('img').forEach(function(im){
var s=im.getAttribute('src')||'';
if(/qr/i.test(s)){
im.style.width='150px';im.style.height='150px';
im.style.maxWidth='150px';im.style.maxHeight='150px';
im.style.objectFit='contain';im.style.display='block';
}
});
}
fixQr();

/* 3) 文案替换（含动态渲染页面的兜底） */
function repAll(root,re,val){
root.querySelectorAll('p,span,strong,em,b,i,h1,h2,h3,h4,h5,button,a,div').forEach(function(el){
if(re.test(el.textContent)&&!el.querySelector('p,div,h1,h2,h3,h4,section,article,ul,li')){
el.textContent=el.textContent.replace(re,val);
}
});
}
function repDyn(){
repAll(document,/三类家当，随取随用。[\s\S]*?下载码。/g,'我的 AI 教学家当！');
var ft=document.querySelector('footer');
if(ft)repAll(ft,/开物AI赋能教育教学全流程/g,'致力于将AI融入教育教学全流程、全要素');
fixQr();
}
repDyn();
new MutationObserver(function(){repDyn()}).observe(document.body,{childList:true,subtree:true});

/* 4) 板块改名 + 编号理顺：实践04 / 案例05 / 视频06 / 关于07(最后) */
var navA=document.querySelector('header nav a[href="#/practice"]');
if(navA)navA.textContent='实践经验';
document.querySelectorAll('.row').forEach(function(r){
var h=r.querySelector('h3');
if(h&&h.textContent.indexOf('AI实践经验')>-1){
h.childNodes[0].textContent='实践经验';
var em=h.querySelector('em');if(em)em.textContent='EXPERIENCE';
}
});

var cont=document.getElementById('kwPage');
if(cont){
function renderAbout(){
cont.innerHTML='<section>'
+'<div class="sec-head" style="margin-top:2.6rem"><span class="sec-no">07 — ABOUT</span><h2>关于开物AI科学</h2><p>好工具，不该锁在抽屉里。</p></div>'
+'<p class="ab-intro">开物AI科学是一个免费开放的中小学科学教学资源库。</p>'
+'<div class="ab-grid"><div class="ab-main">'
+'<div class="ablock"><span class="ab-no">01 — 初心</span><h4>AI，不应该制造教育焦虑</h4><p>AI，并不是为教育制造焦虑和烦恼而生。它应该用来，让每一位老师更加公平地拥有教学能力和资源。无论身处哪所学校、面对什么样的教学条件，我都希望老师们能够平等地学到真正有用的 AI 能力，拿到真正能帮助备课、教学的资源。这，就是我做「开物AI科学」的初心。</p></div>'
+'<div class="ablock"><span class="ab-no">02 — 缘起</span><h4>我为什么做这个网站</h4><p>我是一名小学科学老师。在自己的备课和课堂实践中，我开始越来越多地使用 AI：用 AI 打磨教学设计，用 AI 制作互动教学网页，用 AI 解决一些课堂上“不容易看见、不容易操作、不容易理解”的科学问题，也把自己反复使用、不断修改后真正有效的提示词整理下来。过程中，我发现：AI 真正有价值的地方，并不是让我们追赶一个又一个新工具，而是让一个普通老师，也能够拥有过去需要专业团队才能完成的一些能力。所以，我把这些真实使用过的东西整理出来，放在这里。</p></div>'
+'<div class="ablock"><span class="ab-no">03 — 分享</span><h4>关于分享</h4><p>「开物AI科学」的内容，也会同步分享到公众号、小红书和微信群。我会继续无偿分享 AI 应用指南、教学资源和实践经验。希望 AI 带来的，不只是更快的工作效率，而是让更多普通教师，也能够拥有创造好课堂、做好教学的能力。把 AI 变成科学课的实验器材。这也是「开物AI科学」想做的事情。</p></div>'
+'</div><aside class="ab-side"><div class="ab-qrcard">'
+'<img src="assets/qr.png.jpg" alt="开物AI科学公众号二维码" style="width:150px;height:150px;object-fit:contain;display:block;background:#fff">'
+'<b>公众号 · 开物AI科学</b><span>资源上新与下载码，都在这里发放。微信扫一扫，即刻关注。</span></div></aside></div></section>';
}
function fixNo(){
var m=location.hash.match(/^#\/(practice|cases|videos)/);
if(m){
var map={practice:['04 — EXPERIENCE','实践经验'],cases:['05 — CASES','AI 教学案例'],videos:['06 — VIDEOS','AI 科普视频']}[m[1]];
var no=cont.querySelector('.sec-no'),h2=cont.querySelector('.sec-head h2');
if(no)no.textContent=map[0];
if(h2)h2.textContent=map[1];
}
}
window.addEventListener('hashchange',function(){
if(location.hash==='#/about'){
var hr=document.getElementById('homeRoot'),pg=document.getElementById('pageRoot');
if(hr)hr.style.display='none';if(pg)pg.hidden=true;
document.querySelectorAll('nav a[data-route],nav a[data-kwroute]').forEach(function(a){a.classList.remove('active')});
var ab=document.querySelector('nav a[href="#/about"]');if(ab)ab.classList.add('active');
renderAbout();cont.hidden=false;window.scrollTo(0,0);
}else{fixNo()}
});
fixNo();
if(location.hash==='#/about'){
var hr2=document.getElementById('homeRoot'),pg2=document.getElementById('pageRoot');
if(hr2)hr2.style.display='none';if(pg2)pg2.hidden=true;
renderAbout();cont.hidden=false;
}
}
});
})();
/* ---------- 追加补丁 v4：色散图最终版 / 实践经验8篇 / 关于页删句 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
if(typeof ARTICLES==='undefined')return;

/* 1) 实践经验：8篇真实公众号文章（替换原示例） */
ARTICLES.length=0;
ARTICLES.push(
{id:"a01",t:"DeepSeek+小学科学：我用AI生成的「月相积分系统」，让学生在课堂上玩嗨了！",src:"公众号文章",d:"用 DeepSeek+AI 做出「月相积分系统」，把十个小组的观月任务变成课堂积分游戏。",lk:"https://mp.weixin.qq.com/s/fZInT04dr92NcHKeRIgZIw",img:""},
{id:"a02",t:"蚂蚁「爬」进了PPT？零基础也能做的3D模型，带你玩转AR科学课！",src:"公众号文章",d:"零基础用 AI 生成 3D 模型，把蚂蚁「搬」进 PPT 与 AR 科学课。",lk:"https://mp.weixin.qq.com/s/aiamRE_-CJ35pEXSK_JxTA",img:""},
{id:"a03",t:"用Deepseek生成「基于标准」的教学设计（自带格式）",src:"公众号文章",d:"用 DeepSeek 生成「基于标准」的教学设计：课标引用一字不差，目标动词规范，评价任务一一对应。",lk:"https://mp.weixin.qq.com/s/R_-lllVK1l70_t6s0XB9fw",img:""},
{id:"a04",t:"用AI命制了一份「小学科学期末试卷」！附超详细提示词",src:"公众号文章",d:"用 AI 命制专业化、结构化的小学科学期末试卷，附超详细可复用提示词。",lk:"https://mp.weixin.qq.com/s/ej5TlITaBL0BQszQZoFBuQ",img:""},
{id:"a05",t:"用deepseek生成课堂互动网页，实操+技术指南（提示词可复用）",src:"公众号文章",d:"DeepSeek 生成课堂互动网页的实操与技术指南，从「听话的小车」讲起，提示词可复用。",lk:"https://mp.weixin.qq.com/s/6uVYL4_SXix752J_5s_zmw",img:""},
{id:"a06",t:"用AI生成3D虚拟实验室/开发互动教学资源：提示词解读+万能提示词模版",src:"公众号文章",d:"用 AI 生成 3D 虚拟实验室与互动教学资源，附提示词解读与万能模版。",lk:"https://mp.weixin.qq.com/s/C6oIP34RCnn-64GLzoAyQg",img:""},
{id:"a07",t:"Agent教学指南：中小学教师WorkBuddy教育全景实操指南",src:"公众号文章",d:"中小学教师 WorkBuddy 教育教学全流程实操指南。",lk:"https://mp.weixin.qq.com/s/xaMl51a9K1ITRvgQb25t0A",img:""},
{id:"a08",t:"用workbuddy生成整本教案，生成教学设计skill",src:"公众号文章",d:"用 WorkBuddy 生成全套教案与教学设计 skill 的方法。",lk:"https://mp.weixin.qq.com/s/4Dd61leqqCKZ6c-h3rriCQ",img:""}
);
var cP=document.getElementById('cntPractice');if(cP)cP.textContent=ARTICLES.length;
var sR=document.getElementById('stRes');
if(sR)sR.textContent=(PROMPTS.length+WEB_ITEMS.length+TOOLS.length+DESIGNS.length+ARTICLES.length+CASES.length+VIDEOS.length).toLocaleString('zh-Hans-CN');

/* 2) 色散图最终版：白光水平入射→棱镜内红上紫下散开→出射全斜向下；输入/AI/输出标签 */
var pr=document.querySelector('.prism');
if(pr){
pr.dataset.kwFixed='1';
pr.setAttribute('aria-label','教师输入需求，经AI三棱镜色散出多彩教学成果');
pr.innerHTML=
'<path class="pz" pathLength="1" d="M340 80 L430 380 L250 380 Z"/>'+
'<line class="beam" x1="20" y1="230" x2="294" y2="230"/>'+
'<line class="ray ir1" pathLength="1" x1="295" y1="230" x2="373" y2="190"/>'+
'<line class="ray ir2" pathLength="1" x1="295" y1="230" x2="368" y2="207"/>'+
'<line class="ray ir3" pathLength="1" x1="295" y1="230" x2="373" y2="224"/>'+
'<line class="ray ir4" pathLength="1" x1="295" y1="230" x2="381" y2="250"/>'+
'<line class="ray ir5" pathLength="1" x1="295" y1="230" x2="388" y2="275"/>'+
'<line class="ray ir6" pathLength="1" x1="295" y1="230" x2="406" y2="300"/>'+
'<line class="ray r1" pathLength="1" x1="373" y1="190" x2="620" y2="238"/>'+
'<line class="ray r2" pathLength="1" x1="368" y1="207" x2="620" y2="258"/>'+
'<line class="ray r3" pathLength="1" x1="373" y1="224" x2="620" y2="280"/>'+
'<line class="ray r4" pathLength="1" x1="381" y1="250" x2="620" y2="318"/>'+
'<line class="ray r5" pathLength="1" x1="388" y1="275" x2="620" y2="354"/>'+
'<line class="ray r6" pathLength="1" x1="406" y1="300" x2="620" y2="398"/>'+
'<text class="lab" x="20" y="210">输入</text>'+
'<text class="lab" x="320" y="326" style="fill:#C8451B;font-weight:700;letter-spacing:.35em">AI</text>'+
'<text class="lab" x="488" y="196">输出</text>'+
'<text class="lab fig" x="20" y="442">教师输入需求，AI三棱镜色散出多彩教学成果。</text>';
}

/* 3) 关于页：删掉"开物AI科学是一个免费开放的……"那句 */
function trimAbout(){
var p=document.querySelector('.ab-intro');
if(p&&/免费开放的中小学科学教学资源库/.test(p.textContent))p.remove();
}
trimAbout();
window.addEventListener('hashchange',function(){setTimeout(trimAbout,60)});
});

/* 4) 关于页模板兜底：重新渲染时也不再输出那句话 */
var cont=document.getElementById('kwPage');
if(cont){
var mo=new MutationObserver(function(){
var p=cont.querySelector('.ab-intro');
if(p&&/免费开放的中小学科学教学资源库/.test(p.textContent))p.remove();
});
mo.observe(cont,{childList:true,subtree:true});
}
})();
/* ---------- 追加补丁 v6：七色色散图(光线靠拢) + 下载码修复 ---------- */
/* 下载码兜底名单：把公众号每日发的码填在这里（可同时放多个，建议把明天的也提前加上） */
window.KW_CODES=["KWLPTK"];
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
/* A) 样式：第7色(靛)的出场时间 */
if(!document.getElementById('kwPatch6')){
var st=document.createElement('style');st.id='kwPatch6';
st.textContent=['.ir1,.ir2,.ir3,.ir4,.ir5,.ir6,.ir7{stroke-dasharray:none!important;opacity:.95}',
'.r7{animation-delay:2.25s}'].join('');
document.head.appendChild(st);
}
/* B) 色散图：红橙黄绿蓝靛紫七色；棱镜内扇形收拢(红上紫下)；出射全部斜向下 */
var pr=document.querySelector('.prism');
if(pr){
pr.dataset.kwFixed='1';
pr.setAttribute('aria-label','教师输入需求，AI三棱镜色散出七彩教学成果');
pr.innerHTML=
'<path class="pz" pathLength="1" d="M340 80 L430 380 L250 380 Z"/>'+
'<line class="beam" x1="20" y1="230" x2="295" y2="230"/>'+
'<line class="ray ir1" style="stroke:#C0392B;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="380" y2="212"/>'+
'<line class="ray ir2" style="stroke:#D68227;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="382" y2="219"/>'+
'<line class="ray ir3" style="stroke:#C9A227;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="384" y2="226"/>'+
'<line class="ray ir4" style="stroke:#3F7A3A;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="386" y2="233"/>'+
'<line class="ray ir5" style="stroke:#2C6E9B;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="388" y2="240"/>'+
'<line class="ray ir6" style="stroke:#3D4FA1;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="390" y2="246"/>'+
'<line class="ray ir7" style="stroke:#5B4E9E;stroke-dasharray:none" pathLength="1" x1="295" y1="230" x2="392" y2="253"/>'+
'<line class="ray r1" style="stroke:#C0392B" pathLength="1" x1="380" y1="212" x2="620" y2="262"/>'+
'<line class="ray r2" style="stroke:#D68227" pathLength="1" x1="382" y1="219" x2="620" y2="285"/>'+
'<line class="ray r3" style="stroke:#C9A227" pathLength="1" x1="384" y1="226" x2="620" y2="308"/>'+
'<line class="ray r4" style="stroke:#3F7A3A" pathLength="1" x1="386" y1="233" x2="620" y2="331"/>'+
'<line class="ray r5" style="stroke:#2C6E9B" pathLength="1" x1="388" y1="240" x2="620" y2="354"/>'+
'<line class="ray r6" style="stroke:#3D4FA1" pathLength="1" x1="390" y1="246" x2="620" y2="374"/>'+
'<line class="ray r7" style="stroke:#5B4E9E" pathLength="1" x1="392" y1="253" x2="620" y2="398"/>'+
'<text class="lab" x="20" y="210">输入</text>'+
'<text class="lab" x="326" y="336" style="fill:#C8451B;font-weight:700;letter-spacing:.35em">AI</text>'+
'<text class="lab" x="468" y="216">输出</text>'+
'<text class="lab fig" x="20" y="442">教师输入需求，AI三棱镜色散出七彩教学成果。</text>';
}
/* C) 下载码修复：解除6位输入限制(自动作用于弹窗输入框) */
document.querySelectorAll('input').forEach(function(i){if(i.maxLength>0&&i.maxLength<=6)i.maxLength=12});
new MutationObserver(function(){
document.querySelectorAll('input').forEach(function(i){if(i.maxLength>0&&i.maxLength<=6)i.maxLength=12});
}).observe(document.body,{childList:true,subtree:true});
/* D) 兜底验证：码在 KW_CODES 名单里 → 直接放行下载 */
var pending=null;
document.addEventListener('click',function(e){
var b=e.target.closest('button[data-dl]');
if(b)pending={link:b.getAttribute('data-dl'),name:b.getAttribute('data-name')||''};
},true);
function norm(v){return String(v||'').replace(/\s+/g,'').toUpperCase()}
function inputOf(el){var p=el;for(var i=0;i<6&&p;i++){if(p.querySelector&&p.querySelector('input'))return p.querySelector('input');p=p.parentElement}return null}
function closeModal(from){
var p=from;
for(var i=0;i<6&&p;i++){
if(p.querySelectorAll){
var x=p.querySelector('button');
p.querySelectorAll('button').forEach(function(x){if((x.textContent||'').trim()==='×')x.click()});
if(p.querySelector('input'))break;
}
p=p.parentElement;
}
document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
}
document.addEventListener('click',function(e){
var b=e.target.closest('button');if(!b||!pending)return;
if((b.textContent||'').indexOf('验证')>-1){
var inp=inputOf(b);
if(inp&&window.KW_CODES.indexOf(norm(inp.value))>-1){
window.open(pending.link,'_blank');closeModal(b);
e.stopPropagation();e.preventDefault();
}
}
},true);
document.addEventListener('keydown',function(e){
if(e.key!=='Enter'||!pending||!e.target||e.target.tagName!=='INPUT')return;
if(window.KW_CODES.indexOf(norm(e.target.value))>-1){
window.open(pending.link,'_blank');closeModal(e.target);
e.stopPropagation();e.preventDefault();
}
},true);
});
})();
/* ---------- 追加补丁 v7：实践页强制带图重绘 + 教学案例8条真实数据 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
if(typeof ARTICLES==='undefined'||typeof CASES==='undefined')return;
var E=window.esc||function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})};
var IT=window.isTodo||function(v){return !v||String(v).indexOf('TODO')>-1};

/* 1) 实践经验8篇封面图（重复挂载，幂等） */
var GM={a01:"assets/gzh1.png",a02:"assets/gzh2.png",a03:"assets/gzh3.png",a04:"assets/gzh4.png",a05:"assets/gzh5.png",a06:"assets/gzh6.png",a07:"assets/gzh7.png",a08:"assets/gzh8.png"};
ARTICLES.forEach(function(o){if(GM[o.id])o.img=GM[o.id]});

/* 2) 教学案例8条（按你表格，行序对应 jxal1~8；教学设计不配图） */
CASES.length=0;
CASES.push(
{id:"c01",t:"月有阴晴圆缺",ty:"教学课件",g:"",d:"《月有阴晴圆缺》配套课件，金山文档在线查看。",lk:"https://www.kdocs.cn/l/ciZpQdFxQYaA",img:"assets/jxal1.png"},
{id:"c02",t:"月有阴晴圆缺",ty:"教学设计",g:"",d:"《月有阴晴圆缺》教案，金山文档在线查看。",lk:"https://www.kdocs.cn/l/cjVyvSleYj1o",img:""},
{id:"c03",t:"哺乳动物",ty:"教学课件",g:"",d:"《哺乳动物》教学课件，金山文档在线查看。",lk:"https://www.kdocs.cn/l/cs9AKC8yQUoj",img:"assets/jxal3.png"},
{id:"c04",t:"哺乳动物",ty:"教学设计",g:"",d:"《哺乳动物》教案，金山文档在线查看。",lk:"https://www.kdocs.cn/l/chqrerKwNJjN",img:""},
{id:"c05",t:"两极相遇了",ty:"教学课件",g:"三年级下",d:"《两极相遇了》教学课件，金山文档在线查看。",lk:"https://www.kdocs.cn/l/crRcVqFFfpYw",img:"assets/jxal5.png"},
{id:"c06",t:"两极相遇了",ty:"教学设计",g:"三年级下",d:"《两极相遇了》教学设计，金山文档在线查看。",lk:"https://www.kdocs.cn/l/coI30cLJps1D",img:""},
{id:"c07",t:"两极相遇了 · 课堂实录",ty:"课堂实录",g:"三年级下",d:"AI赋能课堂《两极相遇了》（大象版新教材三年级下册第一单元）课堂实录。",lk:"https://www.bilibili.com/video/BV1wieQ6gEqo/",img:"assets/jxal7.png"},
{id:"c08",t:"不一样的岩石 · 课堂实录",ty:"课堂实录",g:"四年级下",d:"AI赋能课堂《不一样的岩石》（大象版四年级下册第二单元）课堂实录。",lk:"https://www.bilibili.com/video/BV1WSet6GEer/",img:"assets/jxal8.png"}
);

/* 3) 计数更新 */
var cC=document.getElementById('cntCases');if(cC)cC.textContent=CASES.length;
var sR=document.getElementById('stRes');
if(sR)sR.textContent=(PROMPTS.length+WEB_ITEMS.length+TOOLS.length+DESIGNS.length+ARTICLES.length+CASES.length+VIDEOS.length).toLocaleString('zh-Hans-CN');

/* 4) 卡片渲染（图挂了自动隐藏，不会出现裂图） */
function card(o,btn){
var tags=[o.g,o.ty].filter(Boolean).map(function(v){return '<span class="ktag">'+E(v)+'</span>'}).join('');
var cov=o.img?'<img class="kcover" src="'+E(o.img)+'" alt="'+E(o.t)+'" onerror="this.remove()">':'';
return '<div class="kcard" style="cursor:default">'+cov+'<h3>'+E(o.t)+'</h3><div class="ktags">'+tags+'</div><p class="kdesc">'+E(o.d)+'</p><div style="margin-top:.4rem">'+btn+'</div></div>';
}
function grp(title,arr,btnFn){
var h='<p class="kw-sub">'+title+' · '+arr.length+'</p><div class="kgrid">';
if(!arr.length)h+='<p class="s-empty">内容整理中。</p>';
arr.forEach(function(o){h+=card(o,btnFn(o))});
return h+'</div>';
}

/* 5) 覆盖渲染：实践页（带图）/ 案例页（三组新分类） */
var cont=document.getElementById('kwPage');
function override(){
if(!cont||cont.hidden)return;
var h=location.hash||'';
var sec=cont.querySelector('section');if(!sec)return;
var head=sec.querySelector('.sec-head');
if(h.indexOf('#/practice')===0){
var body=grp('公众号文章',ARTICLES.filter(function(o){return o.src==='公众号文章'}),function(o){
return IT(o.lk)?'<span class="ksoon">链接整理中</span>':'<a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">↗ 阅读原文</a>';});
sec.innerHTML='';if(head)sec.appendChild(head);sec.insertAdjacentHTML('beforeend',body);
}else if(h.indexOf('#/cases')===0){
if(head){var p=head.querySelector('p');if(p)p.textContent='教学课件、教学设计与课堂实录，看 AI 如何真实融入一堂课。';}
var body2=grp('教学课件',CASES.filter(function(o){return o.ty==='教学课件'}),function(o){
return IT(o.lk)?'<span class="ksoon">整理中</span>':'<a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">↗ 在线查看</a>';})
+grp('教学设计',CASES.filter(function(o){return o.ty==='教学设计'}),function(o){
return IT(o.lk)?'<span class="ksoon">整理中</span>':'<a class="kbtn ghost" target="_blank" rel="noopener" href="'+E(o.lk)+'">↗ 在线查看</a>';})
+grp('课堂实录',CASES.filter(function(o){return o.ty==='课堂实录'}),function(o){
return IT(o.lk)?'<span class="ksoon">视频整理中</span>':'<a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">▶ 去观看</a>';});
sec.innerHTML='';if(head)sec.appendChild(head);sec.insertAdjacentHTML('beforeend',body2);
}
}
window.addEventListener('hashchange',function(){setTimeout(override,30)});
setTimeout(override,60);setTimeout(override,400);
});
})();
/* ---------- 追加补丁 v8：科普视频59集（AI天文科普52 + AI生活科普7），系列横幅+主题分组 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
if(typeof VIDEOS==='undefined')return;
var E=window.esc||function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})};
var B="https://www.bilibili.com/video/",pad=function(n){return ("0"+n).slice(-2)};

/* 1) 系列一：AI天文科普 52 集 [标题, 主题, BV号] —— 以后加新集，往数组末尾照格式添一行即可 */
var TW=[
["太空梦","月球探索","BV1h6shzXECH"],["太空梦之登月","月球探索","BV1LZs3z6EpR"],["太空梦之探月","月球探索","BV1WBsnz5E5a"],["太空梦之回家","月球探索","BV1Q5szzkEWL"],["太空梦之爱国","月球探索","BV16VsmzjEVD"],
["空间站实验研究","载人航天","BV1ZkykBfEye"],["空间站对接","载人航天","BV1icybBnEk3"],["月球基地实验研究","月球探索","BV1rH1MBgEtH"],
["太空梦之向火星出发","火星探测","BV1HY1qBsESJ"],["太空梦之登陆火星","火星探测","BV12v1CB7EEZ"],["探索火星之心怀祖国","火星探测","BV1rW19BvEzu"],["搭建火星基地","火星探测","BV1wsyPBZE7H"],["探索火星地壳水冰","火星探测","BV1KjyPBPEJd"],["探索火星之钻取水冰","火星探测","BV1it2WB5EBa"],["探索火星之种植土豆","火星探测","BV1vh2WBGE5s"],["探索火星之太阳耀斑爆发","火星探测","BV1A32yB2EJv"],["探索火星之货运飞船","火星探测","BV1fwmJBzES3"],["探索火星之留守系统","火星探测","BV1nZmtBPEc6"],["探索火星之返程","火星探测","BV1o5mGB7E5L"],["探索火星之不忘初心","火星探测","BV1qom2B2ECJ"],["探索火星之返回地球","火星探测","BV1MPmXB7EZy"],["探索火星之成功降落地面","火星探测","BV1chmiBREJP"],
["太空电梯","航天创想","BV1AYqhBQEGU"],["从太空看地球","对地观测","BV1f9q5BAEpU"],
["太空梦之前往木星","木星探测","BV1cfB4BGE6u"],["探索木星之「引力弹弓」","木星探测","BV1viBMB5E9y"],["探索木星之轨道空间站","木星探测","BV1JWB7B5EaC"],["探测木星之探测大红斑","木星探测","BV1akBQBFEsD"],["探索木星之危机袭来","木星探测","BV1WWBYB8Ek9"],["探索木星之登陆木卫二","木星探测","BV1fXBCBJEuY"],["探索木星之搭建木卫二基地","木星探测","BV19vvkBwEeE"],["探索木星之钻探冰层","木星探测","BV1XpveBMExJ"],["探索木星之采集木卫二水样","木星探测","BV1b3iEBsETC"],["探测木星之木卫二冰下世界","木星探测","BV1wCibBVEVc"],["探索木星之返程","木星探测","BV1YyqPBbEen"],["探索木星","木星探测","BV1vNi1BVEdM"],["探索木卫一","木星探测","BV1z3i9B3EdM"],["探索木卫二","木星探测","BV1xwkKBeEU9"],["探索木卫三","木星探测","BV1Z5z6BkEoS"],["探索木卫四","木星探测","BV1emFUzyEHX"],["探索木卫五","木星探测","BV19DZiBTEFz"],["探索木卫六","木星探测","BV15dADzZEHg"],["探索木卫七","木星探测","BV1uhAazAEKZ"],["探索木星之返回地球","木星探测","BV1xNPgzeEq9"],
["探索土星之出发","土星探测","BV1G7PezCE7V"],["探索土星之对接空间站","土星探测","BV1ZUSoBSESV"],["探索土星之穿越土星环","土星探测","BV1RjSoBrEGe"],["探索土星之采集冰晶","土星探测","BV1RjSoBrEgy"],["探索土星之登陆土卫六","土星探测","BV1RLSoBGEC4"],["探索土星之3D打印土卫六基地","土星探测","BV1LLSoBVEjP"],["探索土星之土卫六科研种植","土星探测","BV1QLoFBhEZr"],["探索土星之前往空天航母","土星探测","BV1KrVT6iEfU"]
];
/* 2) 系列二：AI生活科普 7 集 */
var SH=[
["落花生的由来","植物与农业","BV1zc4hzTEHk"],["蔡伦造纸","传统科技","BV1834BzCE9H"],["二十四节气之寒露","气象与节气","BV14kWnzmEqW"],["蚕的一生","动物与生命","BV1iBsczkEFz"],["榫卯：飞檐下的千年智慧","传统科技","BV1FJW1z6Eqc"],["浪漫的霞","大气与天象","BV11fWSzbE6x"],["春夏秋冬","气象与节气","BV146s4zTEZG"]
];
VIDEOS.length=0;
TW.forEach(function(a,i){VIDEOS.push({id:"v"+pad(i+1),t:a[0],src:"AI天文科普",ty:a[1],d:"AI 航天科普系列第 "+(i+1)+" 集，主题："+a[0]+"。AI 生成。",lk:B+a[2]+"/",cover:""})});
SH.forEach(function(a,i){VIDEOS.push({id:"v"+pad(TW.length+i+1),t:a[0],src:"AI生活科普",ty:a[1],d:"AI 生活科普系列第 "+(i+1)+" 集，主题："+a[0]+"。AI 生成。",lk:B+a[2]+"/",cover:""})});
var cV=document.getElementById('cntVideos');if(cV)cV.textContent=VIDEOS.length;
var sR=document.getElementById('stRes');
if(sR)sR.textContent=(PROMPTS.length+WEB_ITEMS.length+TOOLS.length+DESIGNS.length+ARTICLES.length+CASES.length+VIDEOS.length).toLocaleString('zh-Hans-CN');

/* 3) 样式：系列横幅 + 紧凑视频卡 */
if(!document.getElementById('kwPatch8')){
var st=document.createElement('style');st.id='kwPatch8';
st.textContent=['.vbanner{display:flex;gap:1.1rem;align-items:center;border:1px solid var(--line);background:var(--paper);padding:.9rem;margin:1.5rem 0 .5rem}',
'.vbanner img{width:200px;aspect-ratio:16/9;object-fit:cover;border:1px solid var(--line);flex-shrink:0}',
'.vbanner b{font:700 16px var(--serif);display:block;margin-bottom:.35rem;color:var(--ink)}',
'.vbanner span{font:400 12.5px/1.9 var(--sans);color:var(--ink2);display:block}',
'.vgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:.7rem}',
'.vcard{border:1px solid var(--line);background:var(--paper);padding:.7rem .8rem;display:flex;flex-direction:column;gap:.35rem;transition:.2s}',
'.vcard:hover{border-color:var(--accent);transform:translateY(-2px)}',
'.vno{font:500 9.5px var(--mono);letter-spacing:.18em;color:var(--accent)}',
'.vcard b{font:700 13.5px var(--serif);color:var(--ink)}',
'.vcard p{font:400 11.5px/1.75 var(--sans);color:var(--ink2);margin:0;flex:1}',
'.vcard .kbtn{align-self:flex-start;padding:.35em .9em;font-size:11.5px}',
'@media(max-width:640px){.vbanner{flex-direction:column;align-items:flex-start}.vbanner img{width:100%}}'].join('');
document.head.appendChild(st);
}

/* 4) 科普视频页：系列横幅 → 主题分组 → 紧凑卡片 */
var cont=document.getElementById('kwPage');
function vcard(o,idx){
return '<div class="vcard"><span class="vno">EP.'+pad(idx)+'</span><b>'+E(o.t)+'</b><p>'+E(o.d)+'</p><a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">▶ 去观看</a></div>';
}
function vgrp(title,arr,startIdx){
var h='<p class="kw-sub">'+title+' · '+arr.length+'</p><div class="vgrid">';
arr.forEach(function(o,i){h+=vcard(o,startIdx+i+1)});
return h+'</div>';
}
function vbanner(img,name,desc,count){
return '<div class="vbanner"><img src="'+img+'" alt="'+name+'" onerror="this.remove()"><div><b>'+name+'</b><span>'+desc+'（共 '+count+' 集）</span></div></div>';
}
function renderVideos(){
if(!cont||cont.hidden)return;
if((location.hash||'').indexOf('#/videos')!==0)return;
var sec=cont.querySelector('section');if(!sec)return;
var head=sec.querySelector('.sec-head');
if(head){var p=head.querySelector('p');if(p)p.textContent='AI 天文科普与 AI 生活科普系列短视频，点击直达B站观看。课堂导入或课后拓展都能用。';}
var tw=VIDEOS.filter(function(o){return o.src==='AI天文科普'});
var sh=VIDEOS.filter(function(o){return o.src==='AI生活科普'});
var twOrder=["月球探索","载人航天","火星探测","木星探测","土星探测","航天创想","对地观测"];
var shOrder=["植物与农业","动物与生命","气象与节气","大气与天象","传统科技"];
var body=vbanner("assets/kpsp-tw.png","AI 天文科普","从月球、空间站到火星、木星、土星，AI 生成的太空探索系列故事。",tw.length);
var base=0;
twOrder.forEach(function(ty){var g=tw.filter(function(o){return o.ty===ty});if(!g.length)return;body+=vgrp(ty,g,base);base+=g.length});
body+=vbanner("assets/kpsp-sh.png","AI 生活科普","节气、动植物、传统科技与大气天象，AI 生成的生活科普短片。",sh.length);
base=0;
shOrder.forEach(function(ty){var g=sh.filter(function(o){return o.ty===ty});if(!g.length)return;body+=vgrp(ty,g,base);base+=g.length});
sec.innerHTML='';if(head)sec.appendChild(head);sec.insertAdjacentHTML('beforeend',body);
}
window.addEventListener('hashchange',function(){setTimeout(renderVideos,30)});
setTimeout(renderVideos,80);setTimeout(renderVideos,450);
});
})();
/* ---------- 追加补丁 v9：教学设计9条真实链接+提取码，下载时自动复制提取码 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
if(typeof DESIGNS==='undefined')return;

/* 1) 9条真实蓝奏云合集：年级 / 册别 / 链接 / 访问密码 */
var DL=[
{g:"一年级",v:"上",lk:"https://wwbdj.lanzoul.com/b00g4hqova",cd:"279x"},
{g:"二年级",v:"上",lk:"https://wwbdj.lanzoul.com/b00g4hqowb",cd:"50d9"},
{g:"三年级",v:"上",lk:"https://wwbdj.lanzoul.com/b00g4hqoxc",cd:"audk"},
{g:"四年级",v:"上",lk:"https://wwbdj.lanzoul.com/b00g4hqp0f",cd:"ct66"},
{g:"五年级",v:"上",lk:"https://wwbdj.lanzoul.com/b00g4hqp2h",cd:"d9yr"},
{g:"六年级",v:"上",lk:"https://wwbdj.lanzoul.com/b00g4hqp5a",cd:"bm8z"},
{g:"一年级",v:"下",lk:"https://wwbdj.lanzoul.com/b00g4hqp7c",cd:"6fe1"},
{g:"二年级",v:"下",lk:"https://wwbdj.lanzoul.com/b00g4hqpaf",cd:"f2mq"},
{g:"三年级",v:"下",lk:"https://wwbdj.lanzoul.com/b00g4hqpch",cd:"5p2j"}
];

/* 2) 按标题里的年级+册别匹配，替换链接/提取码/简介 */
DESIGNS.forEach(function(o){
var t=String(o.t||'');
for(var i=0;i<DL.length;i++){var r=DL[i];
if(t.indexOf(r.g)>-1&&t.indexOf(r.v)>-1){
o.lk=r.lk;o.cd=r.cd;
o.d="大象版科学"+r.g+r.v+"册 AI 生成教学设计（教案）合集。";
break;}
}
});

/* 3) 已渲染出来的旧按钮（还挂着假链接）就地矫正 */
function findRow(link,name){
for(var i=0;i<DL.length;i++){if(link&&link.indexOf(DL[i].lk)>-1)return DL[i]}
if(name){for(i=0;i<DL.length;i++){var r=DL[i];
if(String(name).indexOf(r.g)>-1&&String(name).indexOf(r.v)>-1)return r}}
return null;
}
function patchDom(){
document.querySelectorAll('button[data-dl],a[data-dl]').forEach(function(b){
var link=b.getAttribute('data-dl')||'';
if(/wwdj\.lanzoul\.com|TODO/.test(link)){
var row=findRow('',b.getAttribute('data-name')||'');
if(row)b.setAttribute('data-dl',row.lk);
}
});
}

/* 4) 提示条 + 复制提取码 */
if(!document.getElementById('kwToastStyle')){
var st=document.createElement('style');st.id='kwToastStyle';
st.textContent='.kw-toast{position:fixed;left:50%;bottom:2.2rem;transform:translateX(-50%);background:#191714;color:#F5F0E6;font:500 12.5px var(--mono,monospace);padding:.7em 1.2em;z-index:99;border-radius:3px;letter-spacing:.06em;box-shadow:0 8px 24px rgba(25,23,20,.25);transition:opacity .3s}';
document.head.appendChild(st);
}
function toast(msg){
var d=document.createElement('div');d.className='kw-toast';d.textContent=msg;document.body.appendChild(d);
setTimeout(function(){d.style.opacity='0'},2200);setTimeout(function(){d.remove()},2600);
}
function copyCd(s){
if(navigator.clipboard&&navigator.clipboard.writeText){
navigator.clipboard.writeText(s).then(function(){toast('提取码 '+s+' 已复制，打开网盘后直接粘贴')},function(){fb(s)});
}else fb(s);
}
function fb(s){var ta=document.createElement('textarea');ta.value=s;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();try{document.execCommand('copy');toast('提取码 '+s+' 已复制，打开网盘后直接粘贴')}catch(e){toast('本条提取码：'+s)}ta.remove()}

/* 5) 点下载按钮 → 立即复制对应提取码（无论之后走网站验证还是兜底名单） */
document.addEventListener('click',function(e){
var b=e.target.closest('button[data-dl],a[data-dl]');if(!b)return;
var row=findRow(b.getAttribute('data-dl')||'',b.getAttribute('data-name')||'');
if(row)copyCd(row.cd);
},true);

patchDom();
window.addEventListener('hashchange',function(){setTimeout(patchDom,80)});
setTimeout(patchDom,300);
try{if(location.hash==='#/design'&&window.route)window.route()}catch(e){}
});
})();
/* ---------- 追加补丁 v10：色散图字号↑ / 统计求真 / 筛选组标题 / 视频页系列置顶+合并网格 ---------- */
(function(){
function ready(fn){document.readyState!=='loading'?fn():document.addEventListener('DOMContentLoaded',fn)}
ready(function(){
var E=window.esc||function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})};
var pad=function(n){n=String(n);while(n.length<2)n='0'+n;return n};

/* A) 样式 */
if(!document.getElementById('kwPatch10')){
var st=document.createElement('style');st.id='kwPatch10';
st.textContent=[
'.prism .lab{font-size:15px!important;font-weight:600;fill:#3B3730}',
'.prism .lab.fig{font-size:14px!important;font-weight:400;letter-spacing:.03em}',
'.kw-flab{display:inline-flex;align-items:center;font:600 12.5px/1 var(--sans);color:#8A5A2B;background:none!important;border:none!important;box-shadow:none!important;padding:0 .8em 0 .15em!important;margin:0!important;cursor:default!important;letter-spacing:.14em}',
'.vindex{display:grid;grid-template-columns:1fr 1fr;gap:.9rem;margin:1.4rem 0 .4rem}',
'.vidx{display:flex;gap:.9rem;align-items:center;border:1px solid var(--line);background:var(--paper);padding:.8rem;cursor:pointer;transition:.2s}',
'.vidx:hover{border-color:var(--accent);transform:translateY(-2px)}',
'.vidx img{width:150px;aspect-ratio:16/9;object-fit:cover;border:1px solid var(--line);flex-shrink:0}',
'.vidx b{font:700 15px var(--serif);display:block;margin-bottom:.3rem;color:var(--ink)}',
'.vidx span{font:400 11.5px/1.8 var(--sans);color:var(--ink2);display:block}',
'.vidx .go{font:500 10px var(--mono);color:var(--accent);letter-spacing:.15em;margin-top:.35em}',
'.vsep{grid-column:1/-1;font:500 11px var(--mono);letter-spacing:.2em;color:var(--accent);border-top:1px solid var(--line);padding:.9rem 0 .2rem}',
'@media(max-width:700px){.vindex{grid-template-columns:1fr}.vidx img{width:110px}}'
].join('');
document.head.appendChild(st);
}

/* B) 色散图：AI 字样再放大一档 */
var pr=document.querySelector('.prism');
if(pr)pr.querySelectorAll('text').forEach(function(t){
if(t.textContent.replace(/\s/g,'')==='AI')t.setAttribute('font-size','17.5');
});

/* C) 筛选行加"年级/类型/学科核心概念"说明标签（纯提示，不可点击） */
function norm(s){return String(s).replace(/\s+/g,'')}
function findLeaf(txt){
var els=document.querySelectorAll('button,span,a,div,li,em,b');
for(var i=0;i<els.length;i++){var el=els[i];
if(el.children.length===0&&el.offsetParent&&norm(el.textContent)===norm(txt))return el}
return null}
var FLAB=[['三年级上','年级'],['仪器','类型'],['10.地球系统','学科核心概念']];
function addFlab(){
if(!findLeaf('10.地球系统'))return;
FLAB.forEach(function(it){
var el=findLeaf(it[0]);if(!el)return;
var prev=el.previousElementSibling;
if(prev&&prev.classList&&prev.classList.contains('kw-flab'))return;
var lab=document.createElement('span');lab.className='kw-flab';lab.textContent=it[1];
lab.setAttribute('aria-hidden','true');
el.parentNode.insertBefore(lab,el);
});
}
setInterval(addFlab,1200);setTimeout(addFlab,400);

/* D) 统计求真：资源总数=七类真实合计；累计下载=真实点击数（本浏览器） */
function TOTAL(){var n=0;
try{n+=PROMPTS.length}catch(e){}try{n+=WEB_ITEMS.length}catch(e){}
try{n+=TOOLS.length}catch(e){}try{n+=DESIGNS.length}catch(e){}
try{n+=ARTICLES.length}catch(e){}try{n+=CASES.length}catch(e){}
try{n+=VIDEOS.length}catch(e){}return n}
function dw(){try{return parseInt(localStorage.getItem('kwDl')||'0',10)||0}catch(e){return 0}}
document.addEventListener('click',function(e){
var b=e.target.closest&&e.target.closest('button[data-dl],a[data-dl]');
if(!b)return;try{localStorage.setItem('kwDl',dw()+1)}catch(err){}
},true);
function labLeaf(label){
var best=null;
document.querySelectorAll('span,div,b,p,i,em,dt,dd,h3,h4').forEach(function(el){
var tx=el.textContent||'';
if(tx.indexOf(label)===-1||tx.length>30)return;
for(var c=el.firstElementChild;c;c=c.nextElementSibling){if((c.textContent||'').indexOf(label)>-1)return}
if(!best||(best.compareDocumentPosition(el)&Node.DOCUMENT_POSITION_FOLLOWING))best=el;
});
return best}
function setStat(label,val){
var leaf=labLeaf(label);if(!leaf)return false;
var cell=leaf.parentElement;
for(var k=0;k<3&&cell;k++){
var tx=cell.textContent||'';
if(tx.indexOf(label)>-1&&/[\d,]/.test(tx)&&tx.length<70)break;
cell=cell.parentElement}
if(!cell)return false;
var ns=cell.querySelectorAll('*');
for(var j=0;j<ns.length;j++){var n=ns[j];
if(n===leaf||n.children.length>0)continue;
var t=(n.textContent||'').trim(),m=t.match(/^([\d,]{1,9})(\s*[^\d,].*)?$/);
if(m){n.textContent=String(val)+(m[2]||'');return true}}
return false}
function assertStats(){setStat('资源总数',TOTAL());setStat('累计下载',dw())}
setTimeout(assertStats,400);setInterval(assertStats,1500);

/* E) 视频页：双系列卡置顶(点击平滑滚动) + 每系列一张连续网格(主题=通栏分隔条) */
var twOrder=["月球探索","载人航天","火星探测","木星探测","土星探测","航天创想","对地观测"];
var shOrder=["植物与农业","动物与生命","气象与节气","大气与天象","传统科技"];
function epcard(o){
var m=String(o.d||'').match(/第\s*(\d+)\s*集/);
return '<div class="vcard"><span class="vno">EP.'+pad(m?m[1]:'--')+'</span><b>'+E(o.t)+'</b><p>'+E(o.d)+'</p><a class="kbtn" target="_blank" rel="noopener" href="'+E(o.lk)+'">▶ 去观看</a></div>';
}
function vidx(img,name,desc,n,target){
return '<div class="vidx" data-target="'+target+'"><img src="'+img+'" alt="'+name+'" onerror="this.remove()"><div><b>'+E(name)+' · '+n+' 集</b><span>'+E(desc)+'</span><span class="go">点击直达该系列 ↓</span></div></div>';
}
function seriesGrid(name,arr,order){
var h='<p class="kw-sub" style="margin-top:1.8rem">'+E(name)+' · '+arr.length+' 集</p><div class="vgrid">';
order.forEach(function(ty){
var g=arr.filter(function(o){return o.ty===ty});if(!g.length)return;
h+='<div class="vsep">'+E(ty)+' · '+g.length+'</div>';
g.forEach(function(o){h+=epcard(o)});
});
return h+'</div>';
}
function renderVideos2(){
var cont=document.getElementById('kwPage');
if(!cont||cont.hidden)return;
if((location.hash||'').indexOf('#/videos')!==0)return;
var sec=cont.querySelector('section');if(!sec)return;
var head=sec.querySelector('.sec-head');
var tw=VIDEOS.filter(function(o){return o.src==='AI天文科普'});
var sh=VIDEOS.filter(function(o){return o.src==='AI生活科普'});
var html='<div class="vindex">'
+vidx('assets/kpsp-tw.png','AI 天文科普','从月球、空间站到火星、木星、土星的太空探索故事。',tw.length,'kw-tw')
+vidx('assets/kpsp-sh.png','AI 生活科普','节气、动植物、传统科技与大气天象。',sh.length,'kw-sh')
+'</div>'
+'<div id="kw-tw">'+seriesGrid('AI 天文科普',tw,twOrder)+'</div>'
+'<div id="kw-sh" style="margin-top:2.2rem">'+seriesGrid('AI 生活科普',sh,shOrder)+'</div>';
sec.innerHTML='';if(head){
var p=head.querySelector('p');if(p)p.textContent='两大 AI 科普系列 · 59 集短视频；点顶部系列卡可直达对应展区。';
sec.appendChild(head)}
sec.insertAdjacentHTML('beforeend',html);
sec.querySelectorAll('.vidx').forEach(function(c){
c.addEventListener('click',function(){
var t=document.getElementById(c.getAttribute('data-target'));
if(t)t.scrollIntoView({behavior:'smooth',block:'start'});
});
});
}
window.addEventListener('hashchange',function(){setTimeout(renderVideos2,45)});
setTimeout(renderVideos2,500);setTimeout(renderVideos2,750);
});
})();

