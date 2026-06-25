const fs = require("fs");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  LevelFormat,
  Packer,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} = require("docx");

const out = "/Users/anita/Documents/room1/妇科手术术前麻醉带教汇总_中英对照.docx";

const border = { style: BorderStyle.SINGLE, size: 1, color: "C9D1D9" };
const borders = { top: border, bottom: border, left: border, right: border };

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before ?? 80, after: opts.after ?? 80, line: 300 },
    alignment: opts.align,
    children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, color: opts.color, size: opts.size ?? 22 })],
  });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}

function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}

function bullet(text, ref = "bullet-list") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 60, line: 300 },
    children: [new TextRun({ text, size: 22 })],
  });
}

function cell(children, width, fill, bold = false) {
  const body = Array.isArray(children) ? children : [p(String(children), { after: 40, bold })];
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: body,
  });
}

function table(headers, rows, widths) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: widths,
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((x, i) => cell([new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: x, bold: true, size: 21 })],
        })], widths[i], "DDEBF7")),
      }),
      ...rows.map((r) => new TableRow({
        children: r.map((x, i) => cell(Array.isArray(x) ? x : String(x), widths[i])),
      })),
    ],
  });
}

const cases = [
  {
    key: "case-a",
    id: "患者 A / Patient A",
    surgery: "腹腔镜下全子宫切除术 / Laparoscopic total hysterectomy",
    basic: "56岁女性，绝经后。因宫颈病变入院，资料提示 TCT HSIL、HPV16 阳性，阴道镜活检提示 CIN2 级，拟行全子宫切除相关手术。 / A 56-year-old postmenopausal woman admitted for cervical disease. Available data show TCT HSIL, HPV16 positivity, and colposcopy biopsy suggesting CIN2. She is scheduled for hysterectomy-related surgery.",
    history: [
      "既往有甲状腺癌手术史，术后未见需放化疗记录；2020 年曾因宫颈病变行宫颈锥形电切术。 / Past history includes thyroid cancer surgery, with no documented need for postoperative chemoradiotherapy in the provided record; she also had cervical conization/electrosurgical excision for cervical disease in 2020.",
      "否认冠心病、高血压、糖尿病、肝炎、结核等病史；否认药物及食物过敏史。无烟酒及毒麻药嗜好。 / She denies coronary heart disease, hypertension, diabetes, hepatitis, tuberculosis, medication or food allergies, and tobacco, alcohol, or illicit drug use.",
      "病程中无发热、胸闷心悸、腹痛腹胀等明显症状。 / During this illness course, she reports no fever, chest tightness, palpitations, abdominal pain, or abdominal distension."
    ],
    tests: [
      "血常规：WBC 5.82 x10^9/L，Hb 124 g/L，PLT 202 x10^9/L，整体无明显贫血或感染提示。 / Complete blood count: WBC 5.82 x10^9/L, Hb 124 g/L, PLT 202 x10^9/L; no obvious anemia or infection signal.",
      "凝血：PT 12.4 s，INR 1.08，APTT 31.0 s，FIB 2.34 g/L，D-dimer 0.15，凝血功能基本正常。 / Coagulation profile: PT 12.4 s, INR 1.08, APTT 31.0 s, fibrinogen 2.34 g/L, D-dimer 0.15; overall coagulation is acceptable.",
      "肝肾功能及代谢：白蛋白 41.4 g/L，肌酐 53 umol/L，eGFR 109 ml/min/1.73m2，葡萄糖 4.59 mmol/L，未见明显麻醉禁忌信号。 / Liver, renal, and metabolic tests: albumin 41.4 g/L, creatinine 53 umol/L, eGFR 109 ml/min/1.73m2, glucose 4.59 mmol/L; no obvious contraindication signal from these data.",
      "心电图：正常心电图。 / ECG: normal electrocardiogram."
    ],
    anesthesia: [
      "重点不是年龄本身，而是腹腔镜气腹和头低位对呼吸循环的影响。 / The key issue is not age alone, but how pneumoperitoneum and Trendelenburg positioning affect ventilation and circulation.",
      "既往甲状腺手术史需术前询问颈部活动度、声音嘶哑、吞咽和气道情况。 / Prior thyroid surgery means we should ask about neck mobility, hoarseness, swallowing difficulty, and potential airway concerns.",
      "宫颈病变本身不一定增加麻醉风险，但手术范围、体位、出血和术后疼痛管理需要关注。 / The cervical lesion itself may not directly increase anesthetic risk, but surgical extent, positioning, bleeding, and postoperative pain control matter."
    ],
    oneLine: "This case is a good example of how anesthesia focuses on physiology, not only the surgical diagnosis."
  },
  {
    key: "case-b",
    id: "患者 B / Patient B",
    surgery: "腹腔镜下全子宫切除术 / Laparoscopic total hysterectomy",
    basic: "46岁女性，发现子宫肌瘤伴增大 10 余年，近 4 月尿频。B 超提示多发子宫肌瘤，最大约 9.6 x 9.5 x 7.1 cm，拟手术治疗。 / A 46-year-old woman with uterine fibroids enlarging for more than 10 years and urinary frequency for about 4 months. Ultrasound suggests multiple uterine fibroids, the largest about 9.6 x 9.5 x 7.1 cm. Surgery is planned.",
    history: [
      "平素体健，否认冠心病、高血压、糖尿病等慢性病史。 / She is generally healthy and denies chronic diseases such as coronary heart disease, hypertension, and diabetes.",
      "既往约 10 年前行尖锐湿疣手术，约 8 年前行痔疮手术。 / Prior surgeries include condyloma acuminatum surgery about 10 years ago and hemorrhoid surgery about 8 years ago.",
      "有青霉素过敏史。无烟酒及毒麻药嗜好。月经规律，LMP 2026-06-06。 / She has a history of penicillin allergy. She denies tobacco, alcohol, or illicit drug use. Menstruation is regular; LMP was 2026-06-06."
    ],
    tests: [
      "血常规：WBC 5.82 x10^9/L，Hb 132 g/L，PLT 166 x10^9/L，未见明显贫血或感染提示。 / Complete blood count: WBC 5.82 x10^9/L, Hb 132 g/L, PLT 166 x10^9/L; no obvious anemia or infection signal.",
      "凝血：PT 10.1 s，INR 0.86，APTT 25.9 s，FIB 2.15 g/L，D-dimer 0.18，基本正常。 / Coagulation profile: PT 10.1 s, INR 0.86, APTT 25.9 s, fibrinogen 2.15 g/L, D-dimer 0.18; overall acceptable.",
      "肝肾功能及代谢：总蛋白 66.2 g/L，白蛋白 41.4 g/L，肌酐 53 umol/L，eGFR 109 ml/min/1.73m2，葡萄糖 4.59 mmol/L，整体可接受。 / Liver, renal, and metabolic tests: total protein 66.2 g/L, albumin 41.4 g/L, creatinine 53 umol/L, eGFR 109 ml/min/1.73m2, glucose 4.59 mmol/L; overall acceptable.",
      "心电图：正常心电图。 / ECG: normal electrocardiogram."
    ],
    anesthesia: [
      "巨大或多发肌瘤可能增加手术时间、出血量和输血准备需求。 / Large or multiple fibroids may increase operative time, blood loss, and the need for blood preparation.",
      "青霉素过敏必须在麻醉记录、围术期抗菌药选择和交接中明确。 / Penicillin allergy must be clearly documented for the anesthesia record, perioperative antibiotic selection, and team handoff.",
      "腹腔镜同样要解释气腹、头低位、通气压力升高和术后恶心呕吐。 / As with other laparoscopic cases, explain pneumoperitoneum, Trendelenburg position, increased airway pressure, and postoperative nausea or vomiting."
    ],
    oneLine: "This case links a pelvic mass to both surgical exposure and anesthetic planning."
  },
  {
    key: "case-c",
    id: "患者 C / Patient C",
    surgery: "经阴道闭孔无张力尿道中段悬吊术 / TVT-O",
    basic: "57岁女性，绝经后。咳嗽后漏尿 7 年，症状逐渐加重，门诊拟压力性尿失禁收住入院，计划行 TVT-O。 / A 57-year-old postmenopausal woman with urine leakage after coughing for 7 years, gradually worsening. She was admitted with suspected stress urinary incontinence and is scheduled for TVT-O.",
    history: [
      "既往高血压病史约 5 年，口服降压治疗；糖尿病病史约 5 年，口服降糖治疗。 / She has a history of hypertension for about 5 years treated with oral antihypertensive medication, and diabetes for about 5 years treated with oral glucose-lowering medication.",
      "约 30 年前因阑尾炎行阑尾切除术。否认冠心病、肝炎、结核等病史，否认药物及食物过敏史。 / She underwent appendectomy for appendicitis about 30 years ago. She denies coronary heart disease, hepatitis, tuberculosis, and medication or food allergies.",
      "体格检查可见 BP 131/95 mmHg，P 63 次/分，身高 156 cm，体重 62.2 kg。 / Physical examination shows BP 131/95 mmHg, pulse 63/min, height 156 cm, and weight 62.2 kg."
    ],
    tests: [
      "血常规：WBC 6.40 x10^9/L，Hb 118 g/L，PLT 313 x10^9/L，血红蛋白接近下限但可接受；嗜碱细胞比例和绝对值升高，需结合临床判断。 / Complete blood count: WBC 6.40 x10^9/L, Hb 118 g/L, PLT 313 x10^9/L. Hemoglobin is near the lower limit but acceptable; elevated basophil percentage and absolute count should be interpreted clinically.",
      "肝肾功能及代谢：GGT 140.8 U/L 升高，ALT/AST 基本正常；肌酐 50 umol/L，eGFR 103 ml/min/1.73m2；电解质基本正常。 / Liver, renal, and metabolic tests: GGT 140.8 U/L is elevated, while ALT/AST are basically normal; creatinine 50 umol/L, eGFR 103 ml/min/1.73m2; electrolytes are generally normal.",
      "心电图：正常心电图。 / ECG: normal electrocardiogram.",
      "资料中未见完整凝血截图，术前需核对凝血功能是否已完成并正常。 / The provided images do not show a complete coagulation panel for this patient; verify before surgery that coagulation testing has been completed and is normal."
    ],
    anesthesia: [
      "TVT-O 通常手术时间较短，但患者合并高血压、糖尿病，需关注血压、血糖和围术期感染风险。 / TVT-O is usually a shorter operation, but hypertension and diabetes make blood pressure, glucose control, and perioperative infection risk important.",
      "截石位可能影响下肢神经、静脉回流和压力点保护。 / Lithotomy position may affect lower-limb nerves, venous return, and pressure-point protection.",
      "需向学生解释压力性尿失禁是腹压升高时尿道支持不足，不是单纯“膀胱太弱”。 / Explain to the student that stress urinary incontinence is caused by insufficient urethral support during increased abdominal pressure, not simply by a weak bladder."
    ],
    oneLine: "This case is ideal for explaining how chronic diseases affect even a relatively short operation."
  }
];

const children = [];

children.push(new Paragraph({
  heading: HeadingLevel.TITLE,
  alignment: AlignmentType.CENTER,
  children: [new TextRun("三例妇科手术术前麻醉带教汇总 / Preoperative Anesthesia Teaching Summary")],
}));
children.push(p("对象：美国康奈尔大学生物学大一学生，未来方向为麻醉 / Audience: Cornell freshman biology student interested in anesthesiology", { align: AlignmentType.CENTER, color: "44546A" }));
children.push(p("说明：本文基于所提供图片整理，已去标识化，仅用于术前带教和临床思维训练；不能替代正式麻醉评估、知情同意或医嘱。图片显示不清之处已按“需核对”处理。 / Note: This document is based on the provided images, is de-identified, and is intended only for preoperative teaching and clinical reasoning practice. It does not replace formal anesthesia assessment, informed consent, or medical orders. Unclear image details are marked as items to verify.", { color: "C00000" }));

children.push(h1("一、三位患者快速总览 / Case Snapshot"));
children.push(table(
  ["病例 / Case", "拟行手术 / Planned surgery", "核心诊断 / Core diagnosis", "麻醉关注点 / Anesthesia focus"],
  [
    ["患者 A / Patient A", cases[0].surgery, "宫颈高级别病变/CIN2，HPV16 阳性 / High-grade cervical lesion/CIN2 with HPV16 positivity", "腹腔镜气腹和头低位；既往甲状腺手术后气道询问；凝血和血红蛋白基本可接受 / Pneumoperitoneum and Trendelenburg position; airway questions after prior thyroid surgery; coagulation and hemoglobin are acceptable"],
    ["患者 B / Patient B", cases[1].surgery, "多发子宫肌瘤，最大约 9.6 cm，伴尿频 / Multiple uterine fibroids, largest about 9.6 cm, with urinary frequency", "潜在出血和手术时间；青霉素过敏；腹腔镜生理影响 / Potential blood loss and operative time; penicillin allergy; physiologic effects of laparoscopy"],
    ["患者 C / Patient C", cases[2].surgery, "压力性尿失禁，合并高血压和糖尿病 / Stress urinary incontinence with hypertension and diabetes", "血压血糖控制；截石位保护；GGT 升高需结合用药和肝胆情况核对 / Blood pressure and glucose control; lithotomy positioning protection; elevated GGT should be interpreted with medications and hepatobiliary history"],
  ],
  [1300, 2600, 2600, 2860]
));

children.push(h1("二、逐例术前情况汇总 / Individual Preoperative Summaries"));
for (const c of cases) {
  children.push(h2(`${c.id} - ${c.surgery}`));
  children.push(table(
    ["项目 / Item", "内容 / Findings"],
    [
      ["基本情况 / Basic profile", c.basic],
      ["病史要点 / History", c.history.map((x) => bullet(x, `${c.key}-hist`))],
      ["主要检查 / Key tests", c.tests.map((x) => bullet(x, `${c.key}-test`))],
      ["麻醉关注点 / Anesthesia focus", c.anesthesia.map((x) => bullet(x, `${c.key}-anes`))],
      ["给学生的一句话 / Teaching sentence", c.oneLine],
    ],
    [2300, 7060]
  ));
}

children.push(h1("三、手术与麻醉教学主线 / Teaching Framework"));
children.push(h2("腹腔镜下全子宫切除术 / Laparoscopic Total Hysterectomy"));
children.push(bullet("腹腔镜手术需要建立 CO2 气腹，腹内压升高会把膈肌向上推，使肺更难展开。 / CO2 pneumoperitoneum raises intra-abdominal pressure, pushes the diaphragm upward, and can make ventilation more difficult."));
children.push(bullet("头低位有利于暴露盆腔，但可能增加静脉回流、气道压力和面部/眼部压力。 / Trendelenburg position improves pelvic exposure, but it can increase venous return, airway pressure, and facial or ocular pressure."));
children.push(bullet("麻醉医生要同时管理无意识、镇痛、肌松、通气、循环和体温。 / The anesthesiologist manages unconsciousness, analgesia, muscle relaxation, ventilation, circulation, and temperature at the same time."));

children.push(h2("TVT-O / Transobturator Tension-free Vaginal Tape"));
children.push(bullet("TVT-O 是通过闭孔路径放置吊带，支撑尿道中段，用于治疗压力性尿失禁。 / TVT-O places a sling through the obturator route to support the mid-urethra and treat stress urinary incontinence."));
children.push(bullet("它通常不是大出血手术，但要关注截石位、膀胱/尿道损伤、尿潴留、感染和慢病控制。 / It is usually not a major blood-loss surgery, but positioning, bladder or urethral injury, urinary retention, infection, and chronic disease control matter."));

children.push(h1("四、检查与术语中英对照 / Medical Terms and Test Interpretation"));
children.push(table(
  ["中文术语 / Chinese term", "英文术语 / English term", "面向学生的解释 / Plain explanation", "麻醉为什么关心 / Why anesthesia cares"],
  [
    ["血红蛋白 Hb / Hemoglobin", "Hemoglobin", "红细胞里携带氧气的蛋白。 / The oxygen-carrying protein inside red blood cells.", "低 Hb 会降低携氧能力，影响失血耐受。 / Low Hb reduces oxygen-carrying capacity and tolerance of blood loss."],
    ["白细胞 WBC / White blood cells", "White blood cell count", "反映感染、炎症或应激的线索。 / A clue to infection, inflammation, or stress.", "感染会影响择期手术时机和围术期风险。 / Infection may affect timing of elective surgery and perioperative risk."],
    ["血小板 PLT / Platelets", "Platelet count", "参与止血的细胞碎片。 / Cell fragments that help form clots and stop bleeding.", "过低会增加出血风险，过高需结合血栓风险。 / Low platelets increase bleeding risk; high platelets should be interpreted with thrombotic risk."],
    ["PT/INR / 凝血酶原时间和国际标准化比值", "Prothrombin time / INR", "评估外源性凝血通路，常与华法林等抗凝有关。 / Tests the extrinsic coagulation pathway and is often related to anticoagulants such as warfarin.", "异常提示出血风险，影响椎管内麻醉和手术安全。 / Abnormal results suggest bleeding risk and affect neuraxial anesthesia and procedural safety."],
    ["APTT / 活化部分凝血活酶时间", "Activated partial thromboplastin time", "评估内源性凝血通路。 / Tests the intrinsic coagulation pathway.", "明显延长可能提示凝血因子问题或抗凝药影响。 / Marked prolongation may indicate coagulation factor problems or anticoagulant effect."],
    ["FIB / 纤维蛋白原", "Fibrinogen", "纤维蛋白原，是形成血凝块的重要原料。 / Fibrinogen is a key raw material for clot formation.", "过低时止血能力下降。 / Low fibrinogen weakens clot formation and hemostasis."],
    ["D-dimer / D-二聚体", "D-dimer", "血栓形成和溶解后的片段。 / A fragment produced after clot formation and breakdown.", "升高需结合症状判断血栓风险，单独轻度异常不等于血栓。 / Elevation must be interpreted with symptoms; a mild isolated elevation does not equal thrombosis."],
    ["ALT/AST / 转氨酶", "Alanine / Aspartate aminotransferase", "肝细胞受损时可升高的酶。 / Enzymes that may rise when liver cells are injured.", "肝功能影响药物代谢和凝血因子合成。 / Liver function affects drug metabolism and synthesis of coagulation factors."],
    ["GGT / γ-谷氨酰转肽酶", "Gamma-glutamyl transferase", "常与胆道、酒精、药物诱导或肝胆问题相关。 / Often associated with biliary disease, alcohol exposure, drug induction, or hepatobiliary conditions.", "升高时需核对病史、用药和其他肝功能指标。 / When elevated, review history, medications, and other liver tests."],
    ["肌酐 Cr/eGFR / 肾功能指标", "Creatinine / estimated GFR", "反映肾脏清除功能。 / Reflects kidney filtration and clearance function.", "肾功能影响药物清除、液体管理和电解质稳定。 / Renal function affects drug clearance, fluid management, and electrolyte stability."],
    ["心电图 ECG / Electrocardiogram", "Electrocardiogram", "记录心脏电活动。 / Records the electrical activity of the heart.", "筛查心律失常、缺血线索，帮助评估围术期心脏风险。 / Screens for arrhythmia or ischemic clues and helps assess perioperative cardiac risk."],
    ["TCT / 宫颈液基细胞学检查", "ThinPrep cytology test", "宫颈细胞学筛查。 / A cervical cytology screening test.", "主要决定妇科诊疗方向，对麻醉风险影响间接。 / It mainly guides gynecologic diagnosis and treatment; its anesthetic impact is indirect."],
    ["HPV16 / 人乳头瘤病毒16型", "Human papillomavirus type 16", "高危型 HPV，与宫颈癌前病变相关。 / A high-risk HPV type associated with cervical precancerous lesions.", "帮助理解手术原因，不直接改变麻醉方案。 / It explains why surgery is needed but usually does not directly change the anesthetic plan."],
    ["CIN2 / 宫颈上皮内瘤变2级", "Cervical intraepithelial neoplasia grade 2", "宫颈上皮中高级别癌前病变。 / A moderate to high-grade precancerous lesion of the cervical epithelium.", "提示手术背景，麻醉重点仍是全身状态和手术方式。 / It provides surgical context; anesthesia still focuses on overall physiology and surgical approach."],
    ["压力性尿失禁 / Stress urinary incontinence", "Stress urinary incontinence", "咳嗽、打喷嚏、运动时腹压升高导致漏尿。 / Urine leakage when coughing, sneezing, or exercising increases abdominal pressure.", "解释 TVT-O 的目的，并关注体位、感染和慢病管理。 / It explains the purpose of TVT-O and highlights positioning, infection risk, and chronic disease management."],
    ["气腹 / Pneumoperitoneum", "Pneumoperitoneum", "向腹腔充入 CO2 形成操作空间。 / Insufflating CO2 into the abdomen to create working space.", "会影响通气、循环和酸碱平衡。 / It can affect ventilation, circulation, and acid-base balance."],
    ["头低位 / Trendelenburg position", "Trendelenburg position", "头低脚高体位，常用于盆腔腹腔镜。 / A head-down, feet-up position often used for pelvic laparoscopy.", "可能增加气道压力、反流风险、面部水肿和眼压。 / It may increase airway pressure, reflux risk, facial edema, and intraocular pressure."],
  ],
  [1700, 2300, 3000, 2360]
));

children.push(h1("五、给大一学生的问题 / Questions for a Freshman Biology Student"));
children.push(table(
  ["问题 / Question", "参考答案 / Suggested answer"],
  [
    ["1. 为什么血红蛋白对麻醉医生很重要？ / Why does hemoglobin matter to anesthesiologists?", "因为麻醉医生需要保证组织获得足够氧气。Hb 低时，即使血氧饱和度看起来正常，总携氧量也可能下降。 / Hemoglobin determines oxygen-carrying capacity. With low Hb, oxygen saturation may look normal but total oxygen delivery can be reduced."],
    ["2. 腹腔镜为什么会让呼吸管理更复杂？ / Why can laparoscopy make ventilation harder?", "CO2 气腹和头低位会抬高膈肌、降低肺顺应性、增加气道压力，还可能造成 CO2 吸收增加。 / Pneumoperitoneum and Trendelenburg position elevate the diaphragm, reduce lung compliance, raise airway pressure, and increase CO2 absorption."],
    ["3. 为什么凝血检查影响麻醉方式？ / Why do coagulation tests affect the anesthetic plan?", "如果凝血异常，椎管内麻醉或有创操作后的出血风险增加。 / Abnormal coagulation increases bleeding risk, especially for neuraxial anesthesia and invasive procedures."],
    ["4. 为什么青霉素过敏在术前要反复确认？ / Why must penicillin allergy be confirmed preoperatively?", "围术期常需预防性抗菌药，过敏史会改变药物选择，并影响抢救准备。 / Prophylactic antibiotics are common; allergy history changes drug choice and preparation for allergic reactions."],
    ["5. 高血压和糖尿病为什么会影响一个短小手术？ / Why do hypertension and diabetes matter even for a short operation?", "它们会影响心血管稳定、感染风险、伤口愈合和围术期血糖管理。 / They affect cardiovascular stability, infection risk, wound healing, and glucose control."],
    ["6. HPV16、TCT、CIN2 与麻醉有什么关系？ / How are HPV16, TCT, and CIN2 related to anesthesia?", "它们解释为什么要手术，但麻醉医生更关心患者全身状态、手术方式和并发症风险。 / They explain why surgery is needed, while anesthesia focuses on whole-body physiology, surgical approach, and perioperative risk."],
    ["7. 为什么 TVT-O 患者要关注截石位？ / Why does lithotomy position matter in TVT-O?", "截石位可能压迫神经和软组织，也会影响静脉回流。 / Lithotomy can compress nerves and soft tissues and affect venous return."],
    ["8. 麻醉医生术前最想回答的三个问题是什么？ / What are the three key preoperative questions?", "能不能安全麻醉？需要改变计划吗？需要提前准备什么？ / Is anesthesia safe enough? Should the plan change? What must be prepared in advance?"],
  ],
  [4300, 5060]
));

children.push(h1("六、带教时可直接使用的开场话术 / Short Teaching Script"));
children.push(p("今天我们不只是看三个妇科病例，而是练习麻醉医生如何把“诊断”翻译成“生理风险”。妇科医生关注病灶在哪里、怎么切；麻醉医生关注病人在睡着以后，氧气、通气、循环、出血、体位和药物代谢是否都还能稳定。 / Today we are not just reviewing three gynecologic cases. We are practicing how anesthesiologists translate a diagnosis into physiologic risk. Surgeons focus on where the lesion is and how to remove it; anesthesiologists focus on oxygenation, ventilation, circulation, bleeding, positioning, and drug handling while the patient is unconscious."));

children.push(h1("七、术前核对清单 / Preoperative Teaching Checklist"));
[
  "核对身份、手术名称、手术部位和知情同意。 / Confirm identity, procedure, surgical site, and consent.",
  "核对禁食禁饮、过敏史、既往麻醉史和困难气道线索。 / Check fasting status, allergies, prior anesthesia, and difficult-airway clues.",
  "核对血常规、凝血、肝肾功能、电解质、血糖、心电图。 / Review CBC, coagulation, liver and renal function, electrolytes, glucose, and ECG.",
  "对腹腔镜病例重点解释气腹和头低位。 / For laparoscopic cases, emphasize pneumoperitoneum and Trendelenburg position.",
  "对 TVT-O 病例重点解释压力性尿失禁、截石位和慢病控制。 / For TVT-O, emphasize stress urinary incontinence, lithotomy position, and chronic disease control.",
  "把“看不清或缺失的检查”列为术前需核对事项。 / Treat unclear or missing data as items to verify before surgery."
].forEach((x) => children.push(bullet(x)));

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 }, paragraph: { spacing: { line: 300 } } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal",
        run: { size: 34, bold: true, color: "1F4E79", font: "Arial" },
        paragraph: { spacing: { before: 240, after: 180 }, alignment: AlignmentType.CENTER } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "1F4E79", font: "Arial" },
        paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: "2F5597", font: "Arial" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, color: "404040", font: "Arial" },
        paragraph: { spacing: { before: 140, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 220 } } } }],
      },
      ...cases.flatMap((c) => ["hist", "test", "anes"].map((k) => ({
        reference: `${c.key}-${k}`,
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 220 } } } }],
      }))),
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1200, right: 1000, bottom: 1000, left: 1000 } } },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "去标识化教学资料 / De-identified teaching material  |  Page ", size: 18 }), new TextRun({ children: [PageNumber.CURRENT], size: 18 })],
      })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(out, buffer);
  console.log(out);
});
