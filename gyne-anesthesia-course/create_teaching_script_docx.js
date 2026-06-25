const fs = require("fs");
const path = require("path");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  LevelFormat,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableOfContents,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} = require("docx");

const out = path.join(__dirname, "gyne-anesthesia-teaching-script-bilingual.docx");
const blue = "284F63";
const teal = "2D8178";
const coral = "A65B50";
const gray = "5D6B78";
const line = "C9D5DC";

function tr(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: 22, ...opts });
}

function p(children, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 100, line: 320 },
    children: Array.isArray(children) ? children : [tr(children)],
    ...opts,
  });
}

function title(text) {
  return new Paragraph({
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { before: 260, after: 120 },
    children: [tr(text, { size: 42, bold: true, color: blue })],
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [tr(text, { size: 30, bold: true, color: blue })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 100 },
    children: [tr(text, { size: 25, bold: true, color: teal })],
  });
}

function label(text) {
  return p([tr(text, { bold: true, color: coral })], { spacing: { before: 140, after: 40 } });
}

function bullet(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 20, after: 60, line: 300 },
    children: [tr(text)],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function cell(text, width, fill = "FFFFFF", bold = false) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: { fill, type: ShadingType.CLEAR },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: line },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: line },
      left: { style: BorderStyle.SINGLE, size: 1, color: line },
      right: { style: BorderStyle.SINGLE, size: 1, color: line },
    },
    children: [p([tr(text, { bold })], { spacing: { before: 40, after: 40 } })],
  });
}

function table(headers, rows, widths) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: widths,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((x, i) => cell(x, widths[i], "DDEBF0", true)),
      }),
      ...rows.map((row) => new TableRow({
        children: row.map((x, i) => cell(x, widths[i])),
      })),
    ],
  });
}

function mixedLine(zh, en) {
  return p([
    tr(zh),
    tr(" "),
    tr(en, { bold: true, color: blue }),
  ]);
}

function sectionBlock({ heading, goal, talk, english, terms, questions }) {
  const children = [h1(heading), label("讲课目标 Teaching goal"), p(goal)];
  children.push(label("你可以这样讲 Suggested Chinese script"));
  talk.forEach((x) => children.push(p(x)));
  children.push(label("可以直接对学生说 English line"));
  english.forEach((x) => children.push(p([tr("“" + x + "”", { italics: true, color: blue })])));
  children.push(label("重点词汇 Key terms"));
  terms.forEach((x) => children.push(bullet(x)));
  children.push(label("可提问问题 Questions to ask"));
  questions.forEach((x) => children.push(bullet(x, "questions")));
  return children;
}

const glossaryRows = [
  ["腹腔镜全子宫切除", "total laparoscopic hysterectomy, TLH", "切除子宫的腹腔镜手术；麻醉重点是气腹、头低位、通气和循环变化。"],
  ["经阴道闭孔无张力尿道中段悬吊术", "transobturator tape, TVT-O", "治疗压力性尿失禁的中段尿道悬吊术；时间较短，但仍需关注苏醒舒适度和术后排尿。"],
  ["全身麻醉", "general anesthesia, GA", "通过药物让患者进入可控的无意识、镇痛和肌松状态。"],
  ["气管插管", "endotracheal intubation, ETT", "将导管置入气管，用于控制通气、保护气道和监测 EtCO2。"],
  ["气腹", "pneumoperitoneum", "腹腔镜手术中向腹腔充入 CO2，为手术创造空间。"],
  ["头低位", "Trendelenburg position", "头低脚高体位，可改善盆腔暴露，也会影响呼吸循环。"],
  ["呼气末二氧化碳", "end-tidal CO2, EtCO2", "确认气管导管位置和评估通气/CO2 排出的关键指标。"],
  ["神经肌肉阻滞", "neuromuscular blockade", "让肌肉松弛，帮助插管和手术暴露。"],
  ["术后恶心呕吐", "postoperative nausea and vomiting, PONV", "妇科手术和全麻后常需预防与处理的问题。"],
];

const sections = [
  {
    heading: "1. 开场介绍 Opening",
    goal: "让学生先知道今天不是单纯“看手术”，而是观察一条完整的围术期麻醉照护链。",
    talk: [
      "今天我们会看三台妇科手术：两台腹腔镜全子宫切除，一台 TVT-O。你可以把它们当作三种不同强度的手术刺激：腹腔镜手术更关注气腹、头低位和通气循环；TVT-O 更短，但仍要关注气道、苏醒和术后舒适度。",
      "你今天最重要的任务不是记住每一种药的剂量，而是理解麻醉医生每一步的目的：为什么要评估气道，为什么要预氧合，为什么要插管，为什么 EtCO2 很重要，为什么苏醒拔管不是最后一步。",
    ],
    english: [
      "Today, please do not try to memorize every drug dose. Focus on the logic: what the surgery does to the patient, and how anesthesia keeps the patient safe.",
      "You can think of anesthesia as continuous protection: before surgery, during surgical stress, and during recovery.",
    ],
    terms: [
      "TLH = total laparoscopic hysterectomy，腹腔镜全子宫切除。",
      "TVT-O = transobturator tape，闭孔路径中段尿道悬吊术。",
      "GA with ETT = general anesthesia with endotracheal tube，全麻气管插管。",
    ],
    questions: [
      "What are the three cases today?",
      "What should an anesthesia observer focus on besides the surgical field?",
    ],
  },
  {
    heading: "2. 术前怎么带她看病人 Pre-op Briefing",
    goal: "把术前核查讲成安全系统，而不是机械流程。",
    talk: [
      "麻醉开始前，我们先确认病人、手术、禁食情况、过敏史、既往麻醉史、气道情况、静脉通路和基础生命体征。每一项都不是形式主义，而是在降低可预防风险。",
      "比如气道评估不是为了“预测一定困难”，而是提前准备。如果患者张口受限、颈部活动差、肥胖或有反流风险，我们的诱导和气道计划就会不同。",
    ],
    english: [
      "Before anesthesia, we confirm the patient, the procedure, fasting status, allergies, previous anesthesia history, airway, IV access, and baseline vital signs.",
      "Airway assessment is not about predicting perfectly. It is about preparing safely.",
    ],
    terms: [
      "fasting status，禁食情况。",
      "allergy，过敏史。",
      "airway assessment，气道评估。",
      "IV access，静脉通路。",
      "baseline vital signs，基础生命体征。",
    ],
    questions: [
      "Why do we ask about fasting before anesthesia?",
      "What airway signs would make you prepare more carefully?",
    ],
  },
  {
    heading: "3. 先讲手术，再讲麻醉 Surgery First",
    goal: "让学生理解麻醉方案围绕手术刺激和体位变化设计。",
    talk: [
      "腹腔镜全子宫切除需要气腹和头低位。气腹会升高腹内压，头低位会让膈肌上移，所以肺顺应性下降、气道压力升高、EtCO2 变化，这些都需要麻醉医生持续调整通气和麻醉深度。",
      "TVT-O 通常时间较短，手术区域更局部，但患者仍处于全麻状态，仍然需要完整监测、气道管理、疼痛和恶心呕吐管理，以及恢复期观察。",
    ],
    english: [
      "For laparoscopic hysterectomy, the surgical position and pneumoperitoneum directly affect ventilation and circulation.",
      "Even a shorter procedure still needs full anesthesia monitoring and a safe recovery plan.",
    ],
    terms: [
      "pneumoperitoneum，气腹。",
      "Trendelenburg position，头低位。",
      "lithotomy position，截石位。",
      "lung compliance，肺顺应性。",
      "surgical stimulation，手术刺激。",
    ],
    questions: [
      "What changes after pneumoperitoneum is established?",
      "Why may airway pressure rise in Trendelenburg position?",
    ],
  },
  {
    heading: "4. 为什么选择全麻气管插管 Why GA with ETT",
    goal: "解释插管的价值：控制通气、保护气道、监测和稳定围术期。",
    talk: [
      "我们常用全身麻醉气管插管，是因为这类手术常需要肌松、气腹和特殊体位。插管后，麻醉医生可以更稳定地控制通气和氧合，也可以通过 EtCO2 持续判断通气状态和导管位置。",
      "气管导管型号不是死公式。通常会结合性别、体型、体重、气道评估和手术需要来选择。临床上最重要的是合适、安全、通气可靠，而不是只记住一个数字。",
    ],
    english: [
      "We use an endotracheal tube because it gives us reliable control of ventilation, oxygenation, and carbon dioxide elimination.",
      "Tube size is chosen by clinical judgment, not by one fixed formula.",
    ],
    terms: [
      "endotracheal tube, ETT，气管导管。",
      "oxygenation，氧合。",
      "ventilation，通气。",
      "airway pressure，气道压力。",
      "tube depth，导管深度。",
    ],
    questions: [
      "What evidence tells us the tube is in the trachea?",
      "Why is EtCO2 more useful than just watching chest movement?",
    ],
  },
  {
    heading: "5. 诱导流程讲稿 Induction Script",
    goal: "把诱导讲成有顺序、有安全检查点的过程。",
    talk: [
      "诱导前先连接监测，确认静脉通路，然后预氧合。预氧合不是多余步骤，它是在给病人建立氧储备，减少插管期间低氧风险。",
      "常用药物可以按角色讲：咪达唑仑帮助镇静和遗忘；舒芬太尼提供镇痛并抑制插管刺激；依托咪酯让病人进入催眠状态；罗库溴铵提供肌松，帮助插管和手术暴露。",
      "给药后我们会面罩通气，确认可以通气，再进行喉镜暴露和气管插管。插管后最重要的确认包括 EtCO2 波形、双肺呼吸音、胸廓起伏、导管深度和气道压力。",
    ],
    english: [
      "Induction is a sequence: monitors, IV access, preoxygenation, medications, ventilation, laryngoscopy, intubation, and confirmation.",
      "The first reliable confirmation of tracheal intubation is a consistent EtCO2 waveform.",
    ],
    terms: [
      "preoxygenation，预氧合。",
      "hypnotic，催眠药。",
      "opioid，阿片类镇痛药。",
      "neuromuscular blocker，神经肌肉阻滞药。",
      "laryngoscopy，喉镜暴露。",
      "bilateral breath sounds，双肺呼吸音。",
    ],
    questions: [
      "Why do we preoxygenate before induction?",
      "What are the main checks after intubation?",
    ],
  },
  {
    heading: "6. 维持麻醉讲稿 Maintenance",
    goal: "让学生理解维持麻醉是在持续滴定平衡，不是把药物打开后不管。",
    talk: [
      "维持期我们会根据手术刺激、生命体征、呼吸参数和恢复目标调整麻醉。七氟醚、丙泊酚、瑞芬太尼分别帮助维持催眠、镇痛和应激控制。药物选择和比例会根据病人和手术变化而调整。",
      "这里可以告诉学生：麻醉医生像在平衡五件事，睡眠、镇痛、肌松、通气和循环。任何一项变化，都可能提示我们需要重新评估。",
    ],
    english: [
      "Maintenance anesthesia is active titration. We adjust anesthetics according to surgical stimulation, vital signs, ventilation, and recovery goals.",
      "The anesthesiologist is balancing sleep, analgesia, relaxation, ventilation, and hemodynamics.",
    ],
    terms: [
      "sevoflurane，七氟醚。",
      "propofol，丙泊酚。",
      "remifentanil，瑞芬太尼。",
      "titration，滴定调整。",
      "hemodynamics，血流动力学。",
    ],
    questions: [
      "If blood pressure rises during surgical stimulation, what could be the possible causes?",
      "Why do we not judge anesthetic depth by one number alone?",
    ],
  },
  {
    heading: "7. 术中观察问题 Intra-op Monitoring",
    goal: "把监护仪讲成临床推理工具。",
    talk: [
      "术中让学生从监护仪开始看。心率和血压反映循环和刺激反应，SpO2 反映氧合，EtCO2 反映通气和 CO2 排出，气道压力反映通气阻力和肺顺应性变化，TOF 帮助判断肌松恢复。",
      "我会强调一句：趋势比单个数字更重要。一个数字变化时，要问它和手术步骤、体位、气腹、药物和通气设置有什么关系。",
    ],
    english: [
      "The monitor is not just numbers. It is a real-time story of circulation, oxygenation, ventilation, relaxation, and surgical stress.",
      "Trend matters more than a single value.",
    ],
    terms: [
      "ECG，心电图。",
      "NIBP，无创血压。",
      "SpO2，脉搏氧饱和度。",
      "EtCO2，呼气末二氧化碳。",
      "TOF, train-of-four，四个成串刺激肌松监测。",
    ],
    questions: [
      "Which monitor changed first after pneumoperitoneum?",
      "What does a sudden loss of EtCO2 waveform make you worry about?",
    ],
  },
  {
    heading: "8. 苏醒拔管与 PACU 交接 Emergence and Handoff",
    goal: "强调苏醒和交接也是麻醉安全的一部分。",
    talk: [
      "手术结束不代表麻醉结束。我们要减少或停止麻醉药，评估自主呼吸、肌松恢复、意识水平、气道保护能力、疼痛、恶心呕吐和体温。只有条件合适，才考虑拔管。",
      "PACU 交接要让接班人员知道：做了什么手术，麻醉过程是否平稳，气道有没有困难，用了哪些关键药物，液体和出血情况如何，术后最需要观察什么。",
    ],
    english: [
      "Emergence is not simply waking up. It is a safety transition from controlled anesthesia to spontaneous recovery.",
      "A good PACU handoff tells the next team what happened, what was given, and what to watch next.",
    ],
    terms: [
      "emergence，苏醒期。",
      "extubation，拔管。",
      "spontaneous ventilation，自主呼吸。",
      "airway protection，气道保护能力。",
      "PACU, post-anesthesia care unit，麻醉后恢复室。",
    ],
    questions: [
      "What conditions should be checked before extubation?",
      "What information must be included in a PACU handoff?",
    ],
  },
  {
    heading: "9. 术后复盘 Debrief",
    goal: "用五个问题帮助学生把观察转化为麻醉思维。",
    talk: [
      "每台手术后，用两分钟复盘五个问题。第一，这是什么手术？第二，为什么适合全麻插管？第三，气道管理最关键的证据是什么？第四，术中哪一个指标变化最值得注意？第五，下一台手术我想主动观察什么？",
      "这样的复盘不追求完整答案，而是训练她把零散画面连接成临床逻辑。",
    ],
    english: [
      "After each case, answer five questions: What was the surgery? Why GA with ETT? How did we confirm the airway? What changed intraoperatively? What should I watch next?",
      "The goal is not a perfect answer. The goal is to connect observation with reasoning.",
    ],
    terms: [
      "debrief，复盘。",
      "clinical reasoning，临床推理。",
      "observation，观察。",
      "reflection，反思。",
    ],
    questions: [
      "What was one thing you understood better after this case?",
      "What is one question you still have about anesthesia?",
    ],
  },
];

const children = [
  title("妇科手术麻醉见习带教讲稿"),
  p([tr("Bilingual Teaching Script for Gynecologic Surgery Anesthesia Observation", { bold: true, color: teal, size: 26 })], { alignment: AlignmentType.CENTER }),
  p([tr("适用对象：美国籍华人本科一年级学生，未来学习方向为麻醉。中文可理解一部分，医学专业词汇以英文标识。", { color: gray })], { alignment: AlignmentType.CENTER }),
  p([tr("使用方式：术前 10-15 分钟讲全局，术中按阶段提醒，术后用五个问题复盘。", { color: gray })], { alignment: AlignmentType.CENTER }),
  h1("使用原则 How to Use This Script"),
  bullet("中文讲逻辑，英文标医学专业词汇；必要时直接读英文句子给学生。"),
  bullet("不讲固定剂量，不替代本院规范、麻醉记录或上级医师判断。"),
  bullet("每台手术抓住三件事：手术刺激、麻醉保护、监测反馈。"),
  bullet("如果时间很紧，只讲 Opening、Why GA with ETT、Induction、Monitoring、Debrief 五部分。"),
  h1("目录 Table of Contents"),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  pageBreak(),
  h1("核心词汇速查 Quick Glossary"),
  table(["中文", "English term", "现场解释"], glossaryRows, [2600, 3000, 3760]),
  pageBreak(),
];

sections.forEach((s, index) => {
  if (index > 0 && [2, 4, 6, 8].includes(index)) children.push(pageBreak());
  children.push(...sectionBlock(s));
});

children.push(
  pageBreak(),
  h1("可直接使用的英文短句 English Phrases"),
  h2("Before induction"),
  mixedLine("麻醉开始前，我们先确认安全信息。", "Before anesthesia, we confirm key safety information."),
  mixedLine("预氧合是在给病人建立氧储备。", "Preoxygenation builds an oxygen reserve before induction."),
  h2("After intubation"),
  mixedLine("我们用 EtCO2 波形确认气管插管。", "We confirm tracheal intubation with a consistent EtCO2 waveform."),
  mixedLine("接下来我们会检查双肺呼吸音、导管深度和气道压力。", "Next, we check bilateral breath sounds, tube depth, and airway pressure."),
  h2("During laparoscopic surgery"),
  mixedLine("气腹和头低位会改变呼吸力学。", "Pneumoperitoneum and Trendelenburg position change respiratory mechanics."),
  mixedLine("如果 EtCO2 或气道压力变化，我们要结合手术步骤一起判断。", "If EtCO2 or airway pressure changes, we interpret it together with the surgical step."),
  h2("Emergence"),
  mixedLine("拔管前要确认病人能安全恢复自主呼吸和气道保护。", "Before extubation, we confirm safe spontaneous ventilation and airway protection."),
  mixedLine("交接的重点是告诉 PACU 接下来要观察什么。", "The key handoff point is what the PACU team should watch next."),
  h1("参考资料与配套课件"),
  p("配套 HTML 课件：/Users/anita/Documents/room1/gyne-anesthesia-course/index.html"),
  p("跨电脑独立版：/Users/anita/Documents/room1/gyne-anesthesia-course/index-standalone.html"),
  p("参考来源：ACOG Hysterectomy FAQ；AUGS Mid-Urethral Sling FAQ；SLS Prevention & Management, Chapter 15。"),
  p([tr("教学声明：本讲稿仅用于见习教学，不替代患者个体化评估、医院规范、麻醉记录或医师临床判断。", { bold: true, color: coral })])
);

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Title",
        name: "Title",
        basedOn: "Normal",
        run: { size: 42, bold: true, color: blue, font: "Arial" },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 } },
      },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 30, bold: true, color: blue, font: "Arial" },
        paragraph: { spacing: { before: 320, after: 140 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 25, bold: true, color: teal, font: "Arial" },
        paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 520, hanging: 260 } } },
        }],
      },
      {
        reference: "questions",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 520, hanging: 260 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 1100, right: 1050, bottom: 1050, left: 1050 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            tr("Gynecologic Anesthesia Observation Teaching Script  |  Page ", { size: 18, color: gray }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: gray }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(out, buffer);
  console.log(out);
});
