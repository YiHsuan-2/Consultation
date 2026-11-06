import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  BorderStyle, 
  AlignmentType, 
  HeadingLevel,
  VerticalAlign,
  ImageRun
} from 'docx';
import { saveAs } from 'file-saver';
import { AppState } from '../types';

const DOCX_STYLES = {
  default: {
    document: {
      run: {
        font: { ascii: "Times New Roman", eastAsia: "標楷體", hAnsi: "Times New Roman", cs: "Times New Roman" },
        size: 24, // 12pt
      },
    },
    heading1: {
      run: {
        font: { ascii: "Times New Roman", eastAsia: "Times New Roman", hAnsi: "Times New Roman" },
        size: 36, // 18pt
        bold: true,
      },
    },
    heading2: {
      run: {
        font: { ascii: "Times New Roman", eastAsia: "Times New Roman", hAnsi: "Times New Roman" },
        size: 32, // 16pt (需求1: 一、二、三、四 Times New Roman, 16pt)
        bold: true,
        color: "2E74B5",
      },
    },
    heading3: {
      run: {
        font: { ascii: "標楷體", eastAsia: "標楷體", hAnsi: "標楷體" },
        size: 28, // 14pt (需求2: 子標題 標楷體, 14pt)
        bold: true,
        color: "2E74B5",
      },
    },
  },
};

const createHeaderCell = (text: string, fill: string = "F1F5F9") => new TableCell({
  children: [new Paragraph({ 
    children: [new TextRun({ 
      text, 
      bold: true,
      font: { ascii: "Times New Roman", eastAsia: "標楷體", hAnsi: "Times New Roman" },
      size: 24 // 12pt
    })],
    alignment: AlignmentType.CENTER 
  })],
  shading: { fill },
  verticalAlign: VerticalAlign.CENTER,
});

const createValueCell = (text: string) => new TableCell({
  children: [new Paragraph({ 
    children: [new TextRun({ 
      text: text || "N/A",
      font: { ascii: "Times New Roman", eastAsia: "標楷體", hAnsi: "Times New Roman" },
      size: 24 // 12pt
    })]
  })],
  verticalAlign: VerticalAlign.CENTER,
});

const formatExercises = (state: AppState, includeFactor = false) => {
  const list = state.clientHx.exerciseList || [];
  let result = "無";
  if (list.length > 0) {
    const formatted = list
      .map(ex => {
        const parts: string[] = [];
        if (ex.frequency) parts.push(ex.frequency);
        if (ex.type) parts.push(ex.type);
        if (ex.name) parts.push(`(${ex.name})`);
        return parts.join(' ').trim();
      })
      .filter(Boolean);
    if (formatted.length > 0) {
      result = formatted.join('、');
    }
  } else {
    const single = state.clientHx.exercise;
    const parts: string[] = [];
    if (single.frequency) parts.push(single.frequency);
    if (single.type) parts.push(single.type);
    if (single.name) parts.push(`(${single.name})`);
    const str = parts.join(' ').trim();
    if (str) {
      result = str;
    }
  }

  if (includeFactor) {
    return `${result} [因子: ${state.clientHx.exercise.activityFactor || 'N/A'}]`;
  }
  return result;
};

const calculateAge = (birthday: string) => {
  if (!birthday) return 0;
  const birthDate = new Date(birthday);
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const m = today.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
};

const calcRecommendedKcalStr = (state: AppState): string => {
  const bmi = parseFloat(state.anthropometry.bmi);
  const weight = parseFloat(state.anthropometry.weight);
  const abw = parseFloat(state.anthropometry.abw);
  const factor = state.clientHx.exercise?.activityFactor;

  if (!bmi || !weight || !factor) return '';

  let baseWeight = weight;
  if (bmi < 18.5 || bmi >= 24) {
    baseWeight = abw;
  }

  if (factor === '無' || factor === '輕度') {
    const minVal = Math.round(baseWeight * 20);
    const maxVal = Math.round(baseWeight * 25);
    return `${minVal} ~ ${maxVal}`;
  } else if (factor === '中度') {
    return `${Math.round(baseWeight * 30)}`;
  } else if (factor === '重度') {
    return `${Math.round(baseWeight * 35)}`;
  }

  return '';
};

const parseCalorie = (val: string | number | undefined | null): number => {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const str = val.toString().trim();
  if (!str) return 0;
  const parts = str.split(/[~-]/).map(p => parseFloat(p.trim())).filter(n => !isNaN(n));
  if (parts.length === 0) return 0;
  return parts[parts.length - 1]; // Return the last element (upper bound)
};

const getRecommendedHBKcal = (state: AppState) => {
  const gender = state.clientHx.gender;
  const weight = parseFloat(state.anthropometry.weight);
  const height = parseFloat(state.anthropometry.height);
  const age = calculateAge(state.clientHx.birthday);

  if (isNaN(weight) || isNaN(height) || age <= 0) {
    return { err: '數據不足', bee: 0, total: 0 };
  }

  let bee = 0;
  if (gender === '男') {
    bee = 66 + 13.7 * weight + 5 * height - 6.8 * age;
  } else {
    bee = 655 + 9.6 * weight + 1.8 * height - 4.7 * age;
  }
  bee = Math.round(bee);
  const act = (state.guidelineSelections as any)?.hbActivity !== undefined ? parseFloat((state.guidelineSelections as any).hbActivity) : 1.3;
  const total = Math.round(bee * act);
  return { err: '', bee, total };
};

const getMacroDistribution = (state: AppState) => {
  const recommendedKcalStr = calcRecommendedKcalStr(state);
  const recKcalVal = parseCalorie(recommendedKcalStr);
  const kcal = parseCalorie(state.diet.targetKcal) || recKcalVal;
  const config = state.intervention?.macroConfig || { carbsPercent: 55, proteinPercent: 15, fatPercent: 30 };
  const cp = parseFloat(config.carbsPercent as any) || 0;
  const pp = parseFloat(config.proteinPercent as any) || 0;
  const fp = parseFloat(config.fatPercent as any) || 0;
  
  if (!kcal) {
    return {
      kcal: 0,
      carbsPer: cp,
      proteinPer: pp,
      fatPer: fp,
      carbsG: '0.0',
      proteinG: '0.0',
      fatG: '0.0'
    };
  }

  return {
    kcal,
    carbsPer: cp,
    proteinPer: pp,
    fatPer: fp,
    carbsG: ((kcal * (cp / 100)) / 4).toFixed(1),
    proteinG: ((kcal * (pp / 100)) / 4).toFixed(1),
    fatG: ((kcal * (fp / 100)) / 9).toFixed(1)
  };
};

const PORT_VALS: Record<string, { p: number, c: number, f: number, k: number }> = {
  '低脂乳品': { p: 8, c: 12, f: 4, k: 120 },
  '全脂乳品': { p: 8, c: 12, f: 8, k: 150 },
  '全榖根莖': { p: 2, c: 15, f: 0, k: 70 },
  '低脂豆魚蛋肉': { p: 7, c: 0, f: 3, k: 55 },
  '中脂豆魚蛋肉': { p: 7, c: 0, f: 5, k: 75 },
  '蔬菜': { p: 1, c: 5, f: 0, k: 25 },
  '水果': { p: 0, c: 15, f: 0, k: 60 },
  '堅果': { p: 0, c: 0, f: 5, k: 45 },
  '低氮澱粉': { p: 1, c: 15, f: 0, k: 64 }
};

export const generateWordDoc = async (state: AppState) => {
  // ... (existing code remains same)
  const doc = new Document({
    styles: DOCX_STYLES,
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          text: "營養諮詢紀錄 (NCP Record)",
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // Basic Info
        new Paragraph({
          children: [
            new TextRun({ text: `姓名: ${state.clientHx.name || "未填寫"}`, bold: true }),
            new TextRun({ text: `\t諮詢日期: ${state.consultDate}`, bold: true }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `諮詢目標: ${state.goal || "未填寫"}`, bold: true }),
          ],
          spacing: { after: 400 },
        }),

        // 1. Assessment
        new Paragraph({
          text: "一、營養評估 (Nutrition Assessment)",
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 400, after: 200 },
        }),

        new Paragraph({ text: "1. 個案史 (Client Hx)", heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }),
        ...(() => {
          // 1. Birthday, Age, Gender: e.g. 1965/03/31, 61y/o, M
          const bdayRaw = (state.clientHx.birthday || '').replace(/-/g, '/');
          const ageVal = calculateAge(state.clientHx.birthday);
          const ageStr = ageVal > 0 ? `${ageVal}y/o` : '';
          const g = state.clientHx.gender;
          const genderStr = (g === '男' || g === 'M' || g === 'male') ? 'M' : ((g === '女' || g === 'F' || g === 'female') ? 'F' : g);
          const bdayLine = [bdayRaw, ageStr, genderStr].filter(Boolean).join(', ') || '--';

          // 2. 工作狀況：在職中
          const jobParts: string[] = [];
          if (state.clientHx.job) jobParts.push(state.clientHx.job);
          if (state.clientHx.jobDescription) jobParts.push(`(${state.clientHx.jobDescription})`);
          const jobStr = jobParts.join(' ') || '未填寫';

          // 3. SHx：與妻子、女兒同住
          const shxStr = state.clientHx.socialHx?.trim() || '無';

          // 4. FHx: HTN(M) -> Maternal HTN, HTN(F) -> Paternal HTN, etc.
          const formatFhxItem = (str: string) => {
            let s = str.trim();
            s = s.replace(/([A-Za-z0-9\u4e00-\u9fa5]+)\s*[（(][Mm母][）)]/g, 'Maternal $1');
            s = s.replace(/([A-Za-z0-9\u4e00-\u9fa5]+)\s*[（(][Ff父][）)]/g, 'Paternal $1');
            s = s.replace(/^[Mm母]\s*[:：]\s*(.*)$/, 'Maternal $1');
            s = s.replace(/^[Ff父]\s*[:：]\s*(.*)$/, 'Paternal $1');
            return s.trim();
          };

          const rawFhx = (state.clientHx.familyHx || '').trim();
          const fhxItems = rawFhx
            ? rawFhx
                .split(/\r?\n|[,，;；、]/)
                .map(s => formatFhxItem(s))
                .filter(Boolean)
            : [];

          // 5. 生活習慣：無菸，約每週一次社交聚餐，聚餐時會飲酒
          const habits = state.clientHx.habits;
          let habitStr = '無';
          if (habits) {
            if (habits.none) {
              habitStr = '無特殊菸酒習慣';
            } else {
              const hParts: string[] = [];
              if (habits.smoke) {
                hParts.push(`抽菸${habits.smokeFrequency ? ` (${habits.smokeFrequency})` : ''}`);
              } else {
                hParts.push('無菸');
              }
              if (habits.drink) {
                const freq = (habits.drinkFrequency || '').trim();
                if (freq) {
                  if (freq.includes('酒') || freq.includes('飲')) {
                    hParts.push(freq);
                  } else {
                    hParts.push(`喝酒 (${freq})`);
                  }
                } else {
                  hParts.push('有飲酒習慣');
                }
              } else {
                if (!habits.smoke && !habits.drink) {
                  hParts.length = 0;
                  hParts.push('無菸酒習慣');
                }
              }
              habitStr = hParts.join('，') || '無';
            }
          }

          // 6. 運動習慣: 若為多個運動 【運動】：【頻率】, 【具體運動】
          const formatExerciseItem = (ex: { type?: string; frequency?: string; name?: string }) => {
            const t = (ex.type || '').trim().replace(/[:：]$/, '');
            const f = (ex.frequency || '').trim();
            const n = (ex.name || '').trim();

            if (t && f && n) return `${t}：${f}, ${n}`;
            if (t && f && !n) return `${t}：${f}`;
            if (t && !f && n) return `${t}：${n}`;
            if (!t && f && n) return `${f}, ${n}`;
            if (n) return n;
            if (t) return t;
            if (f) return f;
            return '';
          };

          const rawExList = state.clientHx.exerciseList || [];
          const exItems = rawExList
            .filter(ex => ex.type?.trim() || ex.frequency?.trim() || ex.name?.trim())
            .map(ex => formatExerciseItem(ex))
            .filter(Boolean);

          if (exItems.length === 0) {
            const single = formatExerciseItem(state.clientHx.exercise || {});
            if (single) exItems.push(single);
          }

          const paragraphs: Paragraph[] = [
            new Paragraph({
              children: [new TextRun({ text: `• ${bdayLine}` })],
              spacing: { after: 60 },
            }),
            new Paragraph({
              children: [new TextRun({ text: `• 工作狀況：${jobStr}` })],
              spacing: { after: 60 },
            }),
            new Paragraph({
              children: [new TextRun({ text: `• SHx：${shxStr}` })],
              spacing: { after: 60 },
            }),
          ];

          // FHx
          if (fhxItems.length > 0) {
            paragraphs.push(
              new Paragraph({
                children: [new TextRun({ text: '• FHx：' })],
                spacing: { after: 40 },
              })
            );
            fhxItems.forEach((item, idx) => {
              paragraphs.push(
                new Paragraph({
                  indent: { left: 400 },
                  children: [new TextRun({ text: item })],
                  spacing: { after: idx === fhxItems.length - 1 ? 60 : 40 },
                })
              );
            });
          } else {
            paragraphs.push(
              new Paragraph({
                children: [new TextRun({ text: '• FHx：無' })],
                spacing: { after: 60 },
              })
            );
          }

          // 生活習慣
          paragraphs.push(
            new Paragraph({
              children: [new TextRun({ text: `• 生活習慣：${habitStr}` })],
              spacing: { after: 60 },
            })
          );

          // 運動習慣
          if (exItems.length > 1) {
            paragraphs.push(
              new Paragraph({
                children: [new TextRun({ text: '• 運動習慣：' })],
                spacing: { after: 40 },
              })
            );
            exItems.forEach((item, idx) => {
              paragraphs.push(
                new Paragraph({
                  indent: { left: 400 },
                  children: [new TextRun({ text: item })],
                  spacing: { after: idx === exItems.length - 1 ? 150 : 40 },
                })
              );
            });
          } else if (exItems.length === 1) {
            paragraphs.push(
              new Paragraph({
                children: [new TextRun({ text: `• 運動習慣：${exItems[0]}` })],
                spacing: { after: 150 },
              })
            );
          } else {
            paragraphs.push(
              new Paragraph({
                children: [new TextRun({ text: '• 運動習慣：無' })],
                spacing: { after: 150 },
              })
            );
          }

          return paragraphs;
        })(),

        new Paragraph({ text: "2. 體位測量 (Anthropometry)", heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }),
        ...(() => {
          const rawH = (state.anthropometry.height || '').trim();
          const h = rawH ? (rawH.endsWith('cm') ? rawH : `${rawH}cm`) : '--';
          const rawW = (state.anthropometry.weight || '').trim();
          const w = rawW ? (rawW.endsWith('kg') ? rawW : `${rawW}kg`) : '--';
          const wDate = state.anthropometry.weightDate ? ` (${state.anthropometry.weightDate})` : '';

          const bmiVal = state.anthropometry.bmi || '--';
          const bmiNum = parseFloat(state.anthropometry.bmi || '0');
          let bmiStatus = '';
          if (!isNaN(bmiNum) && bmiNum > 0) {
            if (bmiNum < 18.5) bmiStatus = ' 體重過輕';
            else if (bmiNum < 24) bmiStatus = ' 正常';
            else if (bmiNum < 27) bmiStatus = ' 過重';
            else bmiStatus = ' 肥胖';
          }

          const ibw = state.anthropometry.ibw || '--';
          const abw = state.anthropometry.abw || '--';
          // 結合體位變化: BW change：【體重】(日期)→【體重】(日期)，日期由最早到最近期
          const weightDateMap = new Map<string, string>();
          if (state.monitoring?.weightHistory) {
            state.monitoring.weightHistory.forEach(r => {
              if (r.date && r.weight && String(r.weight).trim() !== '' && String(r.weight).trim() !== '--') {
                weightDateMap.set(r.date.trim().replace(/-/g, '/'), String(r.weight).trim());
              }
            });
          }
          if (state.monitoring?.history) {
            state.monitoring.history.forEach(r => {
              if (r.date && r.weight && String(r.weight).trim() !== '' && String(r.weight).trim() !== '--') {
                const d = r.date.trim().replace(/-/g, '/');
                if (!weightDateMap.has(d)) {
                  weightDateMap.set(d, String(r.weight).trim());
                }
              }
            });
          }
          if (state.anthropometry.weight && String(state.anthropometry.weight).trim() !== '') {
            const curDate = (state.anthropometry.weightDate || state.consultDate || '').trim().replace(/-/g, '/');
            if (curDate) {
              weightDateMap.set(curDate, String(state.anthropometry.weight).trim());
            }
          }

          const sortedWeightDates = Array.from(weightDateMap.keys()).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());
          const formatW = (val: string) => {
            const v = val.trim();
            return v.endsWith('kg') ? v : `${v}kg`;
          };

          let bwChange = '';
          if (sortedWeightDates.length >= 2) {
            bwChange = sortedWeightDates
              .map(d => `${formatW(weightDateMap.get(d)!)} (${d})`)
              .join(' → ');
          } else if (sortedWeightDates.length === 1) {
            const d = sortedWeightDates[0];
            const wStr = formatW(weightDateMap.get(d)!);
            bwChange = `${wStr} (${d}) ${state.anthropometry.weightChange || '近期無變化'}`.trim();
          } else {
            bwChange = state.anthropometry.weightChange || '近期無變化';
          }

          const rawWaist = (state.anthropometry.waist || '').trim();
          const waistStr = rawWaist ? (rawWaist.endsWith('cm') ? rawWaist : `${rawWaist}cm`) : '--';
          const rawFat = (state.anthropometry.bodyFat || '').trim();
          const fatStr = rawFat ? (rawFat.endsWith('%') ? rawFat : `${rawFat}%`) : '--';

          return [
            new Paragraph({
              children: [
                new TextRun({ text: `• ${h} ,${w}${wDate}` }),
              ],
              spacing: { after: 60 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: `• BMI ${bmiVal}${bmiStatus}` }),
              ],
              spacing: { after: 60 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: `• IBW / ABW： ${ibw} / ${abw}` }),
              ],
              spacing: { after: 60 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: `• BW change：${bwChange}` }),
              ],
              spacing: { after: 60 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: `• 腰圍：${waistStr} ,體脂${fatStr}` }),
              ],
              spacing: { after: 150 },
            }),
          ];
        })(),

        new Paragraph({ text: "3. 生化數值 (Biochemistry)", heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }),
        ...(() => {
          const BIOCHEM_KEY_MAP: { [key: string]: string[] } = {
            'BP': ['BP', 'bp', '血壓'],
            'AC': ['AC', 'ac', '飯前血糖', '空腹血糖'],
            'PC': ['PC', 'pc', '飯後血糖'],
            'HbA1c': ['HbA1c', 'hba1c', '糖化血色素'],
            'BUN': ['BUN', 'bun', '尿素氮'],
            'Cr': ['Cr', 'cr', '肌酸酐'],
            'eGFR': ['eGFR', 'egfr', '腎絲球過濾率'],
            'UPCR': ['UPCR', 'upcr'],
            'UricAcid': ['UricAcid', 'uricAcid', 'uric_acid', '尿酸'],
            'TG': ['TG', 'tg', '三酸甘油酯'],
            'TC': ['TC', 'tc', '總膽固醇'],
            'HDL': ['HDL', 'hdl', '高密度膽固醇'],
            'LDL': ['LDL', 'ldl', '低密度膽固醇'],
            'AST': ['AST', 'ast', 'GOT', 'got'],
            'ALT': ['ALT', 'alt', 'GPT', 'gpt'],
            'Alb': ['Alb', 'alb', '白蛋白'],
            'Na': ['Na', 'na', '鈉'],
            'K': ['K', 'k', '鉀'],
          };

          const ORDERED_INDICATORS = [
            'BP', 'AC', 'PC', 'HbA1c', 'BUN', 'Cr', 'eGFR', 'UPCR', 'UricAcid', 'TG', 'TC', 'HDL', 'LDL', 'AST', 'ALT', 'Alb', 'Na', 'K'
          ];

          const getValFromRecord = (rec: any, displayKey: string): string => {
            const aliases = BIOCHEM_KEY_MAP[displayKey] || [displayKey];
            for (const alias of aliases) {
              if (rec[alias] !== undefined && rec[alias] !== null && String(rec[alias]).trim() !== '') {
                return String(rec[alias]).trim();
              }
            }
            return '';
          };

          // Collect all dates -> { [indicator]: value }
          const dateRecordMap = new Map<string, Record<string, string>>();

          // 1. From biochemHistory
          if (state.monitoring?.biochemHistory) {
            state.monitoring.biochemHistory.forEach(rec => {
              if (rec.date) {
                const d = rec.date.trim().replace(/-/g, '/');
                const cur = dateRecordMap.get(d) || {};
                Object.keys(BIOCHEM_KEY_MAP).forEach(k => {
                  const v = getValFromRecord(rec, k);
                  if (v && !cur[k]) cur[k] = v;
                });
                Object.entries(rec).forEach(([rk, rv]) => {
                  if (!['id', 'date', 'weight', 'other'].includes(rk) && rv !== undefined && rv !== null && String(rv).trim() !== '') {
                    if (!cur[rk]) cur[rk] = String(rv).trim();
                  }
                });
                dateRecordMap.set(d, cur);
              }
            });
          }

          // 2. From legacy monitoring.history
          if (state.monitoring?.history) {
            state.monitoring.history.forEach(rec => {
              if (rec.date) {
                const d = rec.date.trim().replace(/-/g, '/');
                const cur = dateRecordMap.get(d) || {};
                Object.keys(BIOCHEM_KEY_MAP).forEach(k => {
                  const v = getValFromRecord(rec, k);
                  if (v && !cur[k]) cur[k] = v;
                });
                dateRecordMap.set(d, cur);
              }
            });
          }

          // 3. From current state.biochemistry
          const curBioDate = (state.biochemistryDate || state.consultDate || '').trim().replace(/-/g, '/');
          if (curBioDate) {
            const cur = dateRecordMap.get(curBioDate) || {};
            if (state.biochemistry) {
              Object.entries(state.biochemistry).forEach(([k, v]) => {
                const valStr = String(v || '').trim();
                if (valStr) {
                  const matched = Object.keys(BIOCHEM_KEY_MAP).find(dk => 
                    BIOCHEM_KEY_MAP[dk].some(a => a.toLowerCase() === k.toLowerCase())
                  ) || k;
                  cur[matched] = valStr;
                }
              });
            }
            if (Object.keys(cur).length > 0) {
              dateRecordMap.set(curBioDate, cur);
            }
          }

          // Sort dates chronologically (earliest to latest)
          const sortedDates = Array.from(dateRecordMap.keys()).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());

          if (sortedDates.length === 0) {
            return [
              new Paragraph({
                children: [new TextRun({ text: '• 無生化數值紀錄' })],
                spacing: { after: 120 }
              })
            ];
          }

          // Determine which indicators have at least one non-empty value
          const presentIndicators = ORDERED_INDICATORS.filter(ind => {
            return sortedDates.some(d => {
              const v = dateRecordMap.get(d)?.[ind];
              return v && v !== '--';
            });
          });

          // Add custom indicators if any
          for (const [_, rec] of dateRecordMap) {
            for (const k of Object.keys(rec)) {
              if (!presentIndicators.includes(k) && !ORDERED_INDICATORS.includes(k)) {
                if (sortedDates.some(d => dateRecordMap.get(d)?.[k] && dateRecordMap.get(d)?.[k] !== '--')) {
                  presentIndicators.push(k);
                }
              }
            }
          }

          const paragraphs: Paragraph[] = [];

          // Date line: 2025/08/11 → 2025/11/11 → 2026/02/23
          paragraphs.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: sortedDates.join(' → '),
                  bold: true,
                  color: '2E74B5'
                })
              ],
              spacing: { after: 80 }
            })
          );

          // Indicator lines: • BP：-- → -- → 131/96
          presentIndicators.forEach((ind, idx) => {
            const line = sortedDates.map(d => (dateRecordMap.get(d)?.[ind] || '').trim() || '--').join(' → ');
            paragraphs.push(
              new Paragraph({
                children: [
                  new TextRun({ text: `• ${ind}：${line}` })
                ],
                spacing: { after: (idx === presentIndicators.length - 1 && !state.biochemistryNotes) ? 150 : 50 }
              })
            );
          });

          return paragraphs;
        })(),
        state.biochemistryNotes ? new Paragraph({
          children: [
            new TextRun({ text: "備註: ", bold: true }),
            new TextRun({ text: state.biochemistryNotes }),
          ],
          spacing: { before: 80, after: 150 }
        }) : new Paragraph({ text: "" }),

        new Paragraph({ text: "4. 飲食史 (Diet Hx)", heading: HeadingLevel.HEADING_3, spacing: { before: 200 } }),
        new Paragraph({ text: `飲食型態: ${state.diet.type} / 傾向: ${state.diet.preference}` }),
        new Paragraph({ text: `餐次: ${state.diet.meals.join(", ") || (state.diet.mealsOther ? "" : "未填寫")}${state.diet.mealsOther ? (state.diet.meals.length > 0 ? "、" : "") + state.diet.mealsOther : ""}` }),
        new Paragraph({ text: `飲水量: ${state.diet.currentWater || "0"} ml/d${state.diet.currentWaterNotes ? ` (${state.diet.currentWaterNotes})` : ""}` }),
        new Paragraph({ text: `過敏: ${state.diet.allergies.join(", ") || (state.diet.allergiesOther ? "" : "無")}${state.diet.allergiesOther ? (state.diet.allergies.length > 0 ? "、" : "") + state.diet.allergiesOther : ""}` }),
        new Paragraph({ text: `保健品: ${state.diet.supplements || "無"}` }),
        
        // 目前攝取估計 (熱量、碳水化合物、蛋白質、脂肪)
        new Paragraph({ 
          text: "目前飲食攝取估計 (Current Dietary Intake Estimate)", 
          heading: HeadingLevel.HEADING_4, 
          spacing: { before: 150, after: 100 } 
        }),
        (() => {
          const totals = state.diet.logs.reduce((acc, item) => {
            const qty = item.qty || 0;
            const newCategories = { ...acc.categories };
            if (item.category) {
              newCategories[item.category] = (newCategories[item.category] || 0) + qty;
            }
            const carbs = (item.carbs || 0) * qty;
            const protein = (item.protein || 0) * qty;
            const fat = (item.fat || 0) * qty;
            const kcal = (typeof item.kcal === 'number' ? item.kcal : ((item.carbs || 0) * 4 + (item.protein || 0) * 4 + (item.fat || 0) * 9)) * qty;
            return {
              carbs: acc.carbs + carbs,
              protein: acc.protein + protein,
              fat: acc.fat + fat,
              kcal: acc.kcal + kcal,
              categories: newCategories
            };
          }, { carbs: 0, protein: 0, fat: 0, kcal: 0, categories: {} as Record<string, number> });

          const totalKcal = Math.round(totals.kcal);
          const totalCarbs = parseFloat(totals.carbs.toFixed(1));
          const totalProtein = parseFloat(totals.protein.toFixed(1));
          const totalFat = parseFloat(totals.fat.toFixed(1));

          const carbsPct = totalKcal > 0 ? Math.round(((totalCarbs * 4) / totalKcal) * 100) : 0;
          const proteinPct = totalKcal > 0 ? Math.round(((totalProtein * 4) / totalKcal) * 100) : 0;
          const fatPct = totalKcal > 0 ? Math.round(((totalFat * 9) / totalKcal) * 100) : 0;

          const categoryItems = Object.entries(totals.categories)
            .filter(([_, count]) => count > 0)
            .map(([cat, count]) => `${cat}: ${count.toFixed(1)} 份`);

          return new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  createHeaderCell("項目", "E0F2FE"),
                  createHeaderCell("目前估計攝取量", "E0F2FE"),
                  createHeaderCell("佔總熱量比例 / 熱量換算", "E0F2FE"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("總熱量"),
                  createValueCell(`${totalKcal} kcal`),
                  createValueCell(totalKcal > 0 ? "依飲食詳細登錄估算" : "尚未登錄食物記錄"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("CHO"),
                  createValueCell(`${totalCarbs} g`),
                  createValueCell(`${carbsPct} % (${Math.round(totalCarbs * 4)} kcal)`),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("Protein"),
                  createValueCell(`${totalProtein} g`),
                  createValueCell(`${proteinPct} % (${Math.round(totalProtein * 4)} kcal)`),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("Fat / Lipids"),
                  createValueCell(`${totalFat} g`),
                  createValueCell(`${fatPct} % (${Math.round(totalFat * 9)} kcal)`),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("六大類食物份數估計"),
                  new TableCell({
                    columnSpan: 2,
                    children: [new Paragraph({ text: categoryItems.length > 0 ? categoryItems.join("、 ") : "無詳細食物份數紀錄" })],
                    verticalAlign: VerticalAlign.CENTER,
                  }),
                ]
              }),
            ]
          });
        })(),
        state.diet.notes ? new Paragraph({
          children: [
            new TextRun({ text: "飲食傾向：", bold: true }),
            new TextRun({ text: state.diet.notes }),
          ],
          spacing: { before: 100 }
        }) : new Paragraph({ text: "" }),
        state.diet.intakeNotes ? new Paragraph({
          children: [
            new TextRun({ text: "一日飲食內容:", bold: true }),
            new TextRun({ text: state.diet.intakeNotes }),
          ],
          spacing: { before: 100 }
        }) : new Paragraph({ text: "" }),

        // Detailed food log table
        new Paragraph({ 
          text: "飲食史詳細登錄明細 (Dietary Food Logs Detail)", 
          heading: HeadingLevel.HEADING_3, 
          spacing: { before: 200, after: 100 } 
        }),
        state.diet.logs.length === 0 ? new Paragraph({ text: "目前無詳細登錄之食物記錄" }) : (() => {
          let sumCarbs = 0;
          let sumProtein = 0;
          let sumFat = 0;
          let sumKcal = 0;

          state.diet.logs.forEach(item => {
            const qty = item.qty || 0;
            sumCarbs += (item.carbs || 0) * qty;
            sumProtein += (item.protein || 0) * qty;
            sumFat += (item.fat || 0) * qty;
            const itemKcal = (typeof item.kcal === 'number' ? item.kcal : ((item.carbs || 0) * 4 + (item.protein || 0) * 4 + (item.fat || 0) * 9)) * qty;
            sumKcal += itemKcal;
          });

          const roundedTotalKcal = Math.round(sumKcal);

          const formatPct = (macroKcal: number) => {
            if (roundedTotalKcal === 0) return "--";
            const num = (macroKcal / roundedTotalKcal) * 100;
            const str = num.toFixed(1);
            return str.endsWith('.0') ? `${Math.round(num)}%` : `${str}%`;
          };

          const choPctStr = formatPct(sumCarbs * 4);
          const proPctStr = formatPct(sumProtein * 4);
          const fatPctStr = formatPct(sumFat * 9);

          const sortedLogs = [...state.diet.logs].sort((a, b) => {
            const order = ["早餐", "早點", "午餐", "午點", "晚餐", "晚點"];
            const idxA = order.indexOf(a.meal);
            const idxB = order.indexOf(b.meal);
            return (idxA === -1 ? 99 : idxA) - (idxB === -1 ? 99 : idxB);
          });

          return new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  createHeaderCell("餐次"), 
                  createHeaderCell("食物名稱"), 
                  createHeaderCell("份數"),
                  createHeaderCell("醣類 (g)"),
                  createHeaderCell("蛋白質 (g)"),
                  createHeaderCell("脂肪 (g)"),
                  createHeaderCell("熱量 (kcal)"),
                ]
              }),
              ...sortedLogs.map(item => {
                const qty = item.qty || 0;
                const carbs = (item.carbs || 0) * qty;
                const protein = (item.protein || 0) * qty;
                const fat = (item.fat || 0) * qty;
                const kcal = Math.round((typeof item.kcal === 'number' ? item.kcal : (item.carbs * 4 + item.protein * 4 + item.fat * 9)) * qty);

                return new TableRow({
                  children: [
                    createValueCell(item.meal),
                    createValueCell(item.name),
                    createValueCell(`${qty} 份`),
                    createValueCell(carbs.toFixed(1)),
                    createValueCell(protein.toFixed(1)),
                    createValueCell(fat.toFixed(1)),
                    createValueCell(`${kcal}`),
                  ]
                });
              }),
              // 加總 Row
              new TableRow({
                children: [
                  createHeaderCell("--", "F1F5F9"),
                  createHeaderCell("--", "F1F5F9"),
                  createHeaderCell("--", "F1F5F9"),
                  createHeaderCell(sumCarbs.toFixed(1), "F1F5F9"),
                  createHeaderCell(sumProtein.toFixed(1), "F1F5F9"),
                  createHeaderCell(sumFat.toFixed(1), "F1F5F9"),
                  createHeaderCell(`${roundedTotalKcal}`, "F1F5F9"),
                ]
              }),
              // 百分比 Row
              new TableRow({
                children: [
                  createHeaderCell("--", "F1F5F9"),
                  createHeaderCell("--", "F1F5F9"),
                  createHeaderCell("--", "F1F5F9"),
                  createHeaderCell(choPctStr, "F1F5F9"),
                  createHeaderCell(proPctStr, "F1F5F9"),
                  createHeaderCell(fatPctStr, "F1F5F9"),
                  createHeaderCell("--", "F1F5F9"),
                ]
              }),
            ]
          });
        })(),

        new Paragraph({ text: "5. 臨床狀況 (Clinical)", heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                createHeaderCell("既往病史 (Medical Hx)"),
                createValueCell((() => {
                  const hx = state.clinical.medicalHx
                    .map(h => (h === '腎臟病' && state.clinical.kidneyStage ? `腎臟病 (${state.clinical.kidneyStage})` : h));
                  if (state.clinical.medicalHxOther) {
                    hx.push(state.clinical.medicalHxOther);
                  }
                  return hx.length > 0 ? hx.join("、") : "無特定紀錄";
                })()),
                createHeaderCell("腸胃狀況 (GI Status)"),
                createValueCell((() => {
                  const gi = [...state.clinical.giStatus];
                  if (state.clinical.giStatusOther) {
                    gi.push(state.clinical.giStatusOther);
                  }
                  return gi.length > 0 ? gi.join("、") : "正常 / 無特殊不適";
                })()),
              ]
            }),
            new TableRow({
              children: [
                createHeaderCell("排便狀況 (Stool Status)"),
                createValueCell(state.clinical.stoolStatus || "正常 / 未填寫"),
                createHeaderCell("目前服用藥物 (Medications)"),
                createValueCell(state.clinical.medications || "無使用特殊藥物 / 未填寫"),
              ]
            }),
          ]
        }),

        // 2. Diagnosis
        new Paragraph({
          text: "二、營養診斷 (Nutrition Diagnosis)",
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 400, after: 200 },
        }),
        ...state.diagnoses.map((pes, idx) => 
          new Paragraph({
            children: [
              new TextRun({ text: `PES ${idx + 1}: `, bold: true }),
              new TextRun({ text: `${pes.problem === '其他' ? (pes.problemOther || '其他') : pes.problem} (P) 相關於 ${pes.etiology === '其他' ? (pes.etiologyOther || '其他') : pes.etiology} (E) 經由 ${pes.symptom === '其他' ? (pes.symptomOther || '其他') : pes.symptom} (S) 證實。` }),
            ],
            spacing: { after: 100 },
          })
        ),
        state.diagnoses.length === 0 ? new Paragraph({ text: "尚無診斷紀錄" }) : new Paragraph({ text: "" }),

        // 3. Intervention
        new Paragraph({
          text: "三、營養介入 (Nutrition Intervention)",
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 400, after: 200 },
        }),

        // 新增/加強：建議熱量需求與三大營養素比例
        new Paragraph({ 
          text: "1. 建議熱量需求與三大營養素比例", 
          heading: HeadingLevel.HEADING_3, 
          spacing: { before: 200, after: 100 } 
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: (() => {
            const recommendedKcalStr = calcRecommendedKcalStr(state);
            const hb = getRecommendedHBKcal(state);
            const macros = getMacroDistribution(state);
            const targetKcalStr = state.diet.targetKcal ? `${state.diet.targetKcal} kcal/d` : (recommendedKcalStr ? `${recommendedKcalStr} kcal/d` : "未填寫");
            return [
              new TableRow({
                children: [
                  createHeaderCell("項目"),
                  createHeaderCell("設計與建議數值"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("建議熱量需求 (僅供參考)"),
                  createValueCell(recommendedKcalStr ? `${recommendedKcalStr} kcal/d` : "未定"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("建議熱量需求 (Harris Benedict)"),
                  createValueCell(hb.err ? "資料不足" : `BEE: ${hb.bee} kcal / 總計: ${hb.total} kcal/d`),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("設定介入熱量目標"),
                  createValueCell(targetKcalStr),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("醣類 (Carbohydrates)"),
                  createValueCell(`${macros.carbsG} g (${macros.carbsPer}%)`),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("蛋白質 (Protein)"),
                  createValueCell(`${macros.proteinG} g (${macros.proteinPer}%)`),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("脂肪 (Fat / Lipids)"),
                  createValueCell(`${macros.fatG} g (${macros.fatPer}%)`),
                ]
              }),
            ];
          })()
        }),

        // 新增/加強：六大類食物份數計算與建議份量
        new Paragraph({ 
          text: "2. 六大類食物份數計算與建議份量", 
          heading: HeadingLevel.HEADING_3, 
          spacing: { before: 200, after: 100 } 
        }),
        (() => {
          let totalP = 0, totalC = 0, totalF = 0, totalK = 0;
          const rows = [
            '低脂乳品',
            '全脂乳品',
            '全榖根莖',
            '低脂豆魚蛋肉',
            '中脂豆魚蛋肉',
            '蔬菜',
            '水果',
            '堅果',
            '低氮澱粉'
          ].map(key => {
            const val = state.intervention.portions?.[key] || 0;
            const rowInfo = PORT_VALS[key];
            const p = val * rowInfo.p;
            const c = val * rowInfo.c;
            const f = val * rowInfo.f;
            const k = val * rowInfo.k;
            totalP += p;
            totalC += c;
            totalF += f;
            totalK += k;
            
            return new TableRow({
              children: [
                createHeaderCell(key),
                createValueCell(`${val} 份`),
                createValueCell(`${p.toFixed(1)} g`),
                createValueCell(`${c.toFixed(1)} g`),
                createValueCell(`${f.toFixed(1)} g`),
                createValueCell(`${Math.round(k)} kcal`),
              ]
            });
          });

          // Add Total row
          rows.push(new TableRow({
            children: [
              createHeaderCell("份數計算總計", "E2E8F0"),
              createHeaderCell("", "E2E8F0"),
              createHeaderCell(`${totalP.toFixed(1)} g`, "E2E8F0"),
              createHeaderCell(`${totalC.toFixed(1)} g`, "E2E8F0"),
              createHeaderCell(`${totalF.toFixed(1)} g`, "E2E8F0"),
              createHeaderCell(`${Math.round(totalK)} kcal`, "E2E8F0"),
            ]
          }));

          return new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  createHeaderCell("食物類別"),
                  createHeaderCell("設計份數"),
                  createHeaderCell("蛋白質 (g)"),
                  createHeaderCell("醣類 (g)"),
                  createHeaderCell("脂肪 (g)"),
                  createHeaderCell("熱量 (kcal)"),
                ]
              }),
              ...rows
            ]
          });
        })(),

        new Paragraph({ text: "3. 飲食計畫 (Meal Plan)", heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }),
        new Paragraph({ text: `飲食計畫類型: ${state.intervention.dietType || "未填寫"}` }),
        new Paragraph({ text: `衛教重點: ${state.intervention.educationTopics.join(", ") || "無"}` }),
        ...(state.intervention.educationNotes ? [
          new Paragraph({ text: `自訂衛教備註: ${state.intervention.educationNotes}` })
        ] : []),
        new Paragraph({ text: `轉介建議: ${state.intervention.referral || "無"}`, spacing: { after: 100 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                createHeaderCell("餐次"), 
                createHeaderCell("內容"),
              ],
            }),
            ...["早餐", "早點", "午餐", "午點", "晚餐", "晚點"].map(meal => {
              const mealKeyMap: any = {
                "早餐": "breakfast",
                "早點": "morningSnack",
                "午餐": "lunch",
                "午點": "afternoonSnack",
                "晚餐": "dinner",
                "晚點": "eveningSnack"
              };
              const content = Object.values(state.intervention.mealPlan).map(cat => (cat as any)[mealKeyMap[meal]]).filter(Boolean).join(", ");
              return new TableRow({
                children: [
                  createHeaderCell(meal),
                  createValueCell(content || "依計畫執行"),
                ]
              });
            })
          ]
        }),

        // 4. Monitoring
        new Paragraph({
          text: "四、營養監測 (Nutrition Monitoring)",
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 400, after: 200 },
        }),
        ...(state.monitoring.selectedIndicators && state.monitoring.selectedIndicators.length > 0 ? [
          new Paragraph({ text: `追蹤監測項目: ${state.monitoring.selectedIndicators.join("、")}` })
        ] : []),
        new Paragraph({ text: `下次追蹤日期: ${state.monitoring.nextDate || "未定"}` }),
        new Paragraph({ text: `監測計畫: ${state.monitoring.plan || "無"}` }),
        
        new Paragraph({ text: "追蹤紀錄 (History)", heading: HeadingLevel.HEADING_3, spacing: { before: 200 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                createHeaderCell("日期"), createHeaderCell("體重"), createHeaderCell("HbA1c"), createHeaderCell("eGFR"), createHeaderCell("TG/LDL")
              ]
            }),
            ...state.monitoring.history.map(record => new TableRow({
              children: [
                createValueCell(record.date),
                createValueCell(String(record.weight || '--')),
                createValueCell(String(record.hba1c || '--')),
                createValueCell(String(record.egfr || '--')),
                createValueCell(`${record.tg || '--'}/${record.ldl || '--'}`)
              ]
            }))
          ]
        }),

        new Paragraph({
          text: "\n\n營養師簽章: ____________________",
          alignment: AlignmentType.RIGHT,
          spacing: { before: 800 },
        }),
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `營養諮詢紀錄_${state.clientHx.name || "未命名"}_${state.consultDate}.docx`);
};

export const generateReminderWordDoc = async (state: AppState) => {
  const base64ToUint8Array = (base64: string) => {
    const binaryString = atob(base64.split(',')[1]);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
  };

  const imageRuns = state.educationImages.map(img => {
    try {
      return new Paragraph({
        children: [
          new ImageRun({
            data: base64ToUint8Array(img),
            transformation: {
              width: 400,
              height: 225,
            },
          } as any),
        ],
        spacing: { before: 200, after: 200 },
        alignment: AlignmentType.CENTER,
      });
    } catch (e) {
      console.error("Error adding image to Word:", e);
      return new Paragraph({ text: "[圖片無法顯示]" });
    }
  });

  const doc = new Document({
    styles: DOCX_STYLES,
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          text: "營養諮詢小提醒",
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // 諮詢細節
        new Paragraph({ text: "一、諮詢細節", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 200 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                createHeaderCell("姓名"), createValueCell(state.clientHx.name),
                createHeaderCell("生日"), createValueCell(state.clientHx.birthday),
              ],
            }),
            new TableRow({
              children: [
                createHeaderCell("身高 (cm)"), createValueCell(state.anthropometry.height),
                createHeaderCell("體重 (kg)"), createValueCell(state.anthropometry.weight + (state.anthropometry.weightDate ? ` (${state.anthropometry.weightDate})` : '')),
              ],
            }),
            new TableRow({
              children: [
                createHeaderCell("BMI"), createValueCell(state.anthropometry.bmi),
                createHeaderCell("腰圍 (cm)"), createValueCell(state.anthropometry.waist),
              ],
            }),
            new TableRow({
              children: [
                createHeaderCell("諮詢類型"), createValueCell(state.counselingType),
                createHeaderCell("諮詢日期"), createValueCell(state.consultDate),
              ],
            }),
            new TableRow({
              children: [
                createHeaderCell("運動習慣"), 
                new TableCell({
                  columnSpan: 3,
                  children: [new Paragraph({ text: formatExercises(state, false) })],
                  verticalAlign: VerticalAlign.CENTER,
                })
              ],
            }),
            new TableRow({
              children: [
                createHeaderCell("既往病史"), 
                createValueCell(
                  state.clinical.medicalHx
                    .map(h => (h === '腎臟病' && state.clinical.kidneyStage ? `腎臟病 (${state.clinical.kidneyStage})` : h))
                    .concat(state.clinical.medicalHxOther ? [state.clinical.medicalHxOther] : [])
                    .join("、") || "無特定紀錄"
                ),
                createHeaderCell("腸胃狀況"), 
                createValueCell(
                  state.clinical.giStatus
                    .concat(state.clinical.giStatusOther ? [state.clinical.giStatusOther] : [])
                    .join("、") || "正常"
                ),
              ],
            }),
            new TableRow({
              children: [
                createHeaderCell("排便狀況"), createValueCell(state.clinical.stoolStatus || "正常 / 未填寫"),
                createHeaderCell("目前用藥"), createValueCell(state.clinical.medications || "無使用特殊藥物 / 未填寫"),
              ],
            }),
          ],
        }),

        // 生化數據
        new Paragraph({ text: "二、生化數據", heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 200 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: (() => {
            const entries = Object.entries(state.biochemistry).filter(([_, v]) => v);
            const rows: TableRow[] = [];
            for (let i = 0; i < entries.length; i += 3) {
              const chunk = entries.slice(i, i + 3);
              rows.push(new TableRow({
                children: chunk.flatMap(([key, val]) => [
                  createHeaderCell(key, "EFF6FF"),
                  createValueCell(val)
                ])
              }));
            }
            return rows.length > 0 ? rows : [new TableRow({ children: [createValueCell("無紀錄")] })];
          })(),
        }),
        state.biochemistryNotes ? new Paragraph({
          children: [
            new TextRun({ text: "數據分析備註: ", bold: true }),
            new TextRun({ text: state.biochemistryNotes }),
          ],
          spacing: { before: 200 }
        }) : new Paragraph({ text: "" }),

        // 營養控制目標與目前攝取比較
        new Paragraph({ text: "三、營養控制目標 (Nutrition Control Targets)", heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 200 } }),
        new Paragraph({
          text: "1. 建議熱量需求與三大營養素比例",
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 100, after: 100 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: (() => {
            const recommendedKcalStr = calcRecommendedKcalStr(state);
            const hb = getRecommendedHBKcal(state);
            const macros = getMacroDistribution(state);
            const targetKcalStr = state.diet.targetKcal ? `${state.diet.targetKcal} kcal/d` : (recommendedKcalStr ? `${recommendedKcalStr} kcal/d` : "未填寫");
            
            // 目前飲食攝取計算
            const currentDiet = state.diet.logs.reduce((acc, item) => {
              const qty = item.qty || 0;
              const carbs = (item.carbs || 0) * qty;
              const protein = (item.protein || 0) * qty;
              const fat = (item.fat || 0) * qty;
              const kcal = (typeof item.kcal === 'number' ? item.kcal : ((item.carbs || 0) * 4 + (item.protein || 0) * 4 + (item.fat || 0) * 9)) * qty;
              return {
                carbs: acc.carbs + carbs,
                protein: acc.protein + protein,
                fat: acc.fat + fat,
                kcal: acc.kcal + kcal,
              };
            }, { carbs: 0, protein: 0, fat: 0, kcal: 0 });
            
            const currKcal = Math.round(currentDiet.kcal);
            const currCarbs = parseFloat(currentDiet.carbs.toFixed(1));
            const currProtein = parseFloat(currentDiet.protein.toFixed(1));
            const currFat = parseFloat(currentDiet.fat.toFixed(1));
            const currCarbsPct = currKcal > 0 ? Math.round(((currCarbs * 4) / currKcal) * 100) : 0;
            const currProteinPct = currKcal > 0 ? Math.round(((currProtein * 4) / currKcal) * 100) : 0;
            const currFatPct = currKcal > 0 ? Math.round(((currFat * 9) / currKcal) * 100) : 0;

            return [
              new TableRow({
                children: [
                  createHeaderCell("項目"),
                  createHeaderCell("建議與設定目標"),
                  createHeaderCell("目前估計攝取量 (基線)"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("建議熱量需求 (僅供參考)"),
                  createValueCell(recommendedKcalStr ? `${recommendedKcalStr} kcal/d` : "未定"),
                  createValueCell(currKcal > 0 ? `${currKcal} kcal/d` : "無食物記錄"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("建議熱量需求 (Harris Benedict)"),
                  createValueCell(hb.err ? "資料不足" : `BEE: ${hb.bee} kcal / 總計: ${hb.total} kcal/d`),
                  createValueCell(currKcal > 0 ? `${currKcal} kcal/d` : "--"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("設定介入熱量目標"),
                  createValueCell(targetKcalStr),
                  createValueCell(currKcal > 0 ? `${currKcal} kcal/d` : "--"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("醣類 (Carbohydrates)"),
                  createValueCell(`${macros.carbsG} g (${macros.carbsPer}%)`),
                  createValueCell(currKcal > 0 ? `${currCarbs} g (${currCarbsPct}%)` : "--"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("蛋白質 (Protein)"),
                  createValueCell(`${macros.proteinG} g (${macros.proteinPer}%)`),
                  createValueCell(currKcal > 0 ? `${currProtein} g (${currProteinPct}%)` : "--"),
                ]
              }),
              new TableRow({
                children: [
                  createHeaderCell("脂肪 (Fat / Lipids)"),
                  createValueCell(`${macros.fatG} g (${macros.fatPer}%)`),
                  createValueCell(currKcal > 0 ? `${currFat} g (${currFatPct}%)` : "--"),
                ]
              }),
              ...(state.guidelineSelections.target_kcal || state.guidelineSelections.target_protein || state.guidelineSelections.target_carbs || state.guidelineSelections.target_fat ? [
                new TableRow({
                  children: [
                    createHeaderCell("自訂亮點目標 (勾選)"),
                    new TableCell({
                      columnSpan: 2,
                      children: [new Paragraph({
                        text: [
                          state.guidelineSelections.target_kcal ? `熱量: ${state.guidelineSelections.target_kcal}` : '',
                          state.guidelineSelections.target_protein ? `蛋白: ${state.guidelineSelections.target_protein}` : '',
                          state.guidelineSelections.target_carbs ? `醣類: ${state.guidelineSelections.target_carbs}` : '',
                          state.guidelineSelections.target_fat ? `脂肪: ${state.guidelineSelections.target_fat}` : '',
                        ].filter(Boolean).join("、 ")
                      })],
                      verticalAlign: VerticalAlign.CENTER,
                    }),
                  ]
                })
              ] : [])
            ];
          })()
        }),

        // 六大類食物份數計算與建議份量
        new Paragraph({ 
          text: "2. 六大類食物份數計算與建議份量", 
          heading: HeadingLevel.HEADING_3, 
          spacing: { before: 200, after: 100 } 
        }),
        (() => {
          let totalP = 0, totalC = 0, totalF = 0, totalK = 0;
          const rows = [
            '低脂乳品',
            '全脂乳品',
            '全榖根莖',
            '低脂豆魚蛋肉',
            '中脂豆魚蛋肉',
            '蔬菜',
            '水果',
            '堅果',
            '低氮澱粉'
          ].map(key => {
            const val = state.intervention.portions?.[key] || 0;
            const rowInfo = PORT_VALS[key];
            const p = val * rowInfo.p;
            const c = val * rowInfo.c;
            const f = val * rowInfo.f;
            const k = val * rowInfo.k;
            totalP += p;
            totalC += c;
            totalF += f;
            totalK += k;
            
            return new TableRow({
              children: [
                createHeaderCell(key, "F0F9FF"),
                createValueCell(`${val} 份`),
                createValueCell(`${p.toFixed(1)} g`),
                createValueCell(`${c.toFixed(1)} g`),
                createValueCell(`${f.toFixed(1)} g`),
                createValueCell(`${Math.round(k)} kcal`),
              ]
            });
          });

          // Add Total row
          rows.push(new TableRow({
            children: [
              createHeaderCell("份數計算總計", "EFF6FF"),
              createHeaderCell("", "EFF6FF"),
              createHeaderCell(`${totalP.toFixed(1)} g`, "EFF6FF"),
              createHeaderCell(`${totalC.toFixed(1)} g`, "EFF6FF"),
              createHeaderCell(`${totalF.toFixed(1)} g`, "EFF6FF"),
              createHeaderCell(`${Math.round(totalK)} kcal`, "EFF6FF"),
            ]
          }));

          return new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  createHeaderCell("食物類別", "DBEAFE"),
                  createHeaderCell("設計份數", "DBEAFE"),
                  createHeaderCell("蛋白質 (g)", "DBEAFE"),
                  createHeaderCell("醣類 (g)", "DBEAFE"),
                  createHeaderCell("脂肪 (g)", "DBEAFE"),
                  createHeaderCell("熱量 (kcal)", "DBEAFE"),
                ]
              }),
              ...rows
            ]
          });
        })(),

        // 減重營養方針與個人化追蹤卡 (若選擇減重營養方針)
        ...(() => {
          if (state.counselingType === '減重營養方針') {
            const g = state.guidelineSelections || {};
            const elements = [
              new Paragraph({ text: "四、減重營養方針與個人化追蹤卡", heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 200 } }),
              new Paragraph({
                text: "【減重核心方針】",
                heading: HeadingLevel.HEADING_3,
                spacing: { before: 100, after: 100 }
              }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  new TableRow({
                    children: [
                      createHeaderCell("類別", "F1F5F9"),
                      createHeaderCell("重點行動原則", "F1F5F9"),
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("① 醣類"),
                      createValueCell([
                        g['weight_guide_fixed_carbs'] ? "• 每餐固定主食份量，避免忽多忽少" : "",
                        g['weight_guide_less_sugar'] ? "• 減少精製糖及精製澱粉" : "",
                        g['weight_guide_whole_grain'] ? "• 主食優先選擇：糙米、五穀飯、全穀、地瓜" : "",
                        g['weight_guide_less_bakery'] ? "• 減少：甜麵包、蛋糕、餅乾、精緻麵食" : "",
                        g['weight_guide_no_thickening'] ? "• 避免額外勾芡、糖醋、濃稠醬汁" : "",
                      ].filter(Boolean).join("\n") || "• 每餐固定主食份量\n• 減少精製糖及精製澱粉\n• 主食優先選擇：糙米、五穀飯、全穀、地瓜\n• 避免額外勾芡、糖醋、濃稠醬汁")
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("② 纖維"),
                      createValueCell([
                        g['weight_guide_veg_half_bowl'] ? "• 每餐至少半碗蔬菜" : "",
                        g['weight_guide_veg_colors'] ? "• 優先選擇多種類、不同顏色蔬菜" : "",
                        g['weight_guide_eat_order'] ? "• 建議進食順序：先吃菜 → 蛋白質 → 最後吃主食" : "",
                        g['weight_guide_fruit_mod'] ? "• 水果適量，不以果汁取代水果" : "",
                      ].filter(Boolean).join("\n") || "• 每餐至少半碗蔬菜\n• 建議進食順序：先吃菜 → 蛋白質 → 最後吃主食\n• 水果適量，不以果汁取代水果")
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("③ 蛋白質"),
                      createValueCell([
                        g['weight_guide_protein_match'] ? "• 每餐搭配一份蛋白質" : "",
                        g['weight_guide_protein_priority'] ? "• 優先選擇：魚、雞肉、蛋、豆腐、豆製品" : "",
                        g['weight_guide_no_only_starch'] ? "• 避免只吃澱粉類（麵包+飲料、稀飯+醬菜、麵+飲料）" : "",
                        g['weight_guide_keep_muscle'] ? "• 減重期間注意蛋白質攝取，以維持肌肉量" : "",
                      ].filter(Boolean).join("\n") || "• 每餐搭配一份蛋白質\n• 優先選擇：魚、雞肉、蛋、豆腐、豆製品\n• 避免只吃澱粉類；維持肌肉量")
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("④ 含糖飲料"),
                      createValueCell([
                        g['weight_guide_no_sugar_drinks'] ? "• 完全避免含糖飲料（手搖飲、汽水、果汁、甜咖啡、運動飲料）" : "",
                        g['weight_guide_tea_water'] ? "• 改成：水／無糖茶／無糖咖啡" : "",
                      ].filter(Boolean).join("\n") || "• 以完全避免含糖飲料為目標\n• 改成：水／無糖茶／無糖咖啡")
                    ]
                  }),
                ]
              }),
              new Paragraph({
                text: "【追蹤指標＋營養計畫（這週只做 1–2 個改變）】",
                heading: HeadingLevel.HEADING_3,
                spacing: { before: 200, after: 100 }
              }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  new TableRow({
                    children: [
                      createHeaderCell("體重追蹤"),
                      createValueCell(`本次：${state.anthropometry.weight || '--'} kg   ➔   下次目標：${g['weight_goal_target_kg'] || g['weight_track_target'] || '______'} kg`),
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("飲食狀況"),
                      createValueCell(`每日飲食紀錄：[${g['weight_track_diet_log'] || '未填'}]   |   本週最需要改善：${g['weight_track_diet_improve'] || '______'}`),
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("飽足感與飲食行為"),
                      createValueCell(`飯後是否容易太飽：[${g['weight_track_fullness'] || '未填'}]   |   容易嘴饞/額外點心：[${g['weight_track_cravings'] || '未填'}]   |   含糖飲料：${g['weight_track_sugar_drinks_freq'] || '______'} 次/週`),
                    ]
                  }),
                  ...(g['weight_track_ppg_enabled'] ? [
                    new TableRow({
                      children: [
                        createHeaderCell("飯後血糖變化"),
                        createValueCell(`飯後2小時：${g['weight_track_ppg_val'] || '___'} mg/dL   |   明顯升高：[${g['weight_track_ppg_high'] || '未填'}]\n(觀察重點：挑選 1–2 餐觀察「吃了什麼 → 吃多少 → 飯後血糖／身體反應」)`),
                      ]
                    })
                  ] : []),
                  new TableRow({
                    children: [
                      createHeaderCell("本次營養計畫 (目標)"),
                      createValueCell([
                        g['weight_plan_veg'] ? "☑ 每餐至少半碗蔬菜" : "",
                        g['weight_plan_fixed_carbs'] ? "☑ 每餐固定主食份量" : "",
                        g['weight_plan_add_protein'] ? "☑ 每餐增加一份蛋白質" : "",
                        g['weight_plan_sugar_free'] ? "☑ 含糖飲料改為無糖" : "",
                        g['weight_plan_walk'] ? "☑ 飯後走路 10–15 分鐘" : "",
                        g['weight_plan_log_meals'] ? "☑ 記錄 1–2 餐飲食" : "",
                        g['weight_plan_other'] ? `☑ ${g['weight_plan_other']}` : "",
                      ].filter(Boolean).join("\n") || "□ 每餐至少半碗蔬菜\n□ 每餐固定主食份量\n□ 每餐增加一份蛋白質\n□ 含糖飲料改為無糖\n□ 飯後走路 10–15 分鐘"),
                    ]
                  }),
                  new TableRow({
                    children: [
                      createHeaderCell("下次追蹤安排"),
                      createValueCell(`📅 日期：${g['weight_plan_next_date'] || state.monitoring.nextDate || '____________'}\n下次主要討論：\n① ${g['weight_plan_discuss_1'] || '____________________'}\n② ${g['weight_plan_discuss_2'] || '____________________'}`),
                    ]
                  }),
                ]
              }),
            ];
            return elements;
          }
          return [];
        })(),

        // 備註與注意事項
        new Paragraph({ text: state.counselingType === '減重營養方針' ? "五、備註與注意事項" : "四、備註與注意事項", heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 200 } }),
        new Paragraph({ text: state.reminderNotes || "無" }),

        // 衛教資訊
        new Paragraph({ text: state.counselingType === '減重營養方針' ? "六、衛教資訊與附件" : "五、衛教資訊與附件", heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 200 } }),
        new Paragraph({ text: state.intervention.educationTopics.join(", ") || "無" }),
        ...(state.intervention.educationNotes ? [
          new Paragraph({ text: `自訂衛教備註: ${state.intervention.educationNotes}` })
        ] : []),
        ...imageRuns,

        new Paragraph({
          text: "\n\n營養師: " + state.dietitian,
          alignment: AlignmentType.RIGHT,
          spacing: { before: 800 },
        }),
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `諮詢小提醒_${state.clientHx.name || "未命名"}_${state.consultDate}.docx`);
};
