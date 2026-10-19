import { FoodItem } from '../types';
import { FOOD_DATABASE } from '../constants';

export interface DietSuggestion {
  id: string;
  meal: string;
  rawText: string;
  food: FoodItem;
  qty: number;
}

// Meal mapping helpers
const MEAL_NORM_MAP: Record<string, string> = {
  '早餐': '早餐',
  '早': '早餐',
  '早點': '早點',
  '上午點心': '早點',
  '午餐': '午餐',
  '中餐': '午餐',
  '午': '午餐',
  '中午': '午餐',
  '午點': '午點',
  '午點心': '午點',
  '下午茶': '午點',
  '點心': '午點',
  '晚餐': '晚餐',
  '晚': '晚餐',
  '晚上': '晚餐',
  '晚點': '晚點',
  '晚點心': '晚點',
  '宵夜': '晚點',
  '夜點': '晚點',
};

// Common Taiwanese food synonyms and aliases (mapping to standard food names in allFoods)
const FOOD_ALIASES: Record<string, string> = {
  '荷包蛋': '荷包蛋',
  '煎蛋': '煎蛋',
  '蛋': '荷包蛋',
  '雞蛋': '荷包蛋',
  '水煮蛋': '水煮蛋 / 茶葉蛋',
  '茶葉蛋': '水煮蛋 / 茶葉蛋',
  '滷蛋': '滷蛋 1顆',
  '青菜': '燙青菜 / 炒青菜 1份',
  '燙青菜': '燙青菜 / 炒青菜 1份',
  '炒青菜': '燙青菜 / 炒青菜 1份',
  '蔬菜': '蔬菜100g(半碗)',
  '白飯': '白飯 40g(1/4碗)',
  '飯': '白飯 40g(1/4碗)',
  '米飯': '白飯 40g(1/4碗)',
  '地瓜': '烤地瓜 55g',
  '烤地瓜': '烤地瓜 55g',
  '吐司': '烤吐司 1片',
  '烤吐司': '烤吐司 1片',
  '土司': '烤吐司 1片',
  '豆漿': '無糖豆漿 240cc',
  '無糖豆漿': '無糖豆漿 240cc',
  '牛奶': '全脂牛奶 240 cc',
  '鮮奶': '全脂牛奶 240 cc',
  '低脂牛奶': '低脂牛奶 240 cc',
  '鍋燒意麵': '鍋燒意麵',
  '意麵': '鍋燒意麵',
  '蛋餅': '原味蛋餅',
  '原味蛋餅': '原味蛋餅',
  '滷肉飯': '滷肉飯 (小碗)',
  '排骨便當': '排骨便當',
  '排骨飯': '排骨便當',
  '雞腿便當': '雞腿便當',
  '雞腿飯': '雞腿便當',
  '牛肉麵': '牛肉麵',
  '水餃': '水餃 (10顆)',
  '餃子': '水餃 (10顆)',
  '陽春麵': '陽春麵 (湯)',
  '蘋果': '蘋果 (1顆)',
  '香蕉': '香蕉 (大半根/小1根)',
  '芭樂': '芭樂 (1/3顆)',
  '拿鐵': '拿鐵咖啡 (無糖 1杯)',
  '拿鐵咖啡': '拿鐵咖啡 (無糖 1杯)',
  '雞胸肉': '雞胸肉 30g',
  '板豆腐': '板豆腐 80g (1/2塊）',
  '豆腐': '板豆腐 80g (1/2塊）',
};

// Quantity parser
const parsePortion = (phrase: string): { cleanName: string; qty: number } => {
  let cleaned = phrase.trim();
  let qty = 1;

  if (/(?:^|\s)半(?:碗|份|顆|個|片|條|杯|盤)?(?:\s|$)/.test(cleaned) || /半碗|半份|半顆/.test(cleaned)) {
    qty = 0.5;
    cleaned = cleaned.replace(/(?:^|\s)半(?:碗|份|顆|個|片|條|杯|盤)?(?:\s|$)/g, ' ').trim();
  }

  const numUnitMatch = cleaned.match(/(\d+(?:\.\d+)?)\s*(?:份|碗|顆|個|條|片|杯|盤|包|瓶|罐|匙|餐)/);
  if (numUnitMatch) {
    const val = parseFloat(numUnitMatch[1]);
    if (!isNaN(val) && val > 0) {
      qty = val;
      cleaned = cleaned.replace(numUnitMatch[0], ' ').trim();
    }
  }

  const multMatch = cleaned.match(/[*xX×]\s*(\d+(?:\.\d+)?)/);
  if (multMatch) {
    const val = parseFloat(multMatch[1]);
    if (!isNaN(val) && val > 0) {
      qty = val;
      cleaned = cleaned.replace(multMatch[0], ' ').trim();
    }
  }

  cleaned = cleaned.replace(/[（(][^）)]*[）)]/g, ' ').replace(/[\[【][^\]】]*[\]】]/g, ' ').trim();
  cleaned = cleaned.replace(/^[\d+.\-•*、，,\s]+/, '').replace(/[\s+.\-、，,]+$/, '').trim();

  return { cleanName: cleaned, qty };
};

// Match ONLY against existing foods in database (never create non-existent foods)
const findBestFoodInDatabase = (cleanName: string, allFoods: FoodItem[]): FoodItem | null => {
  if (!cleanName || allFoods.length === 0) return null;
  const cn = cleanName.toLowerCase().trim();

  // 1. Exact name match
  const exact = allFoods.find(f => f.name.toLowerCase().trim() === cn);
  if (exact) return exact;

  // 2. Base name match (e.g. "白飯" matching "白飯 40g(1/4碗)")
  const baseNameMatch = allFoods.find(f => {
    const base = f.name.split(/[\s(（]/)[0].toLowerCase().trim();
    return base === cn;
  });
  if (baseNameMatch) return baseNameMatch;

  // 3. Alias dictionary check (must exist in allFoods)
  if (FOOD_ALIASES[cn]) {
    const aliasTarget = FOOD_ALIASES[cn];
    const aliasFood = allFoods.find(f => f.name === aliasTarget || f.name.includes(aliasTarget));
    if (aliasFood) return aliasFood;
  }

  // 4. Prefix match
  const prefixMatch = allFoods.find(f => {
    const fn = f.name.toLowerCase();
    const base = fn.split(/[\s(（]/)[0];
    return fn.startsWith(cn) || cn.startsWith(base);
  });
  if (prefixMatch) return prefixMatch;

  // 5. Substring inclusion match
  const substrMatch = allFoods.find(f => {
    const fn = f.name.toLowerCase();
    const base = fn.split(/[\s(（]/)[0];
    return fn.includes(cn) || (cn.length >= 2 && cn.includes(base));
  });
  if (substrMatch) return substrMatch;

  // 6. Sub-keyword matching for specific food keywords (e.g. "荷包蛋" might match "滷蛋" or "煎蛋")
  const coreKeywords = ['蛋', '青菜', '菜', '飯', '麵', '肉', '雞', '魚', '牛', '豬', '豆腐', '奶', '豆漿', '地瓜', '便當', '茶', '果'];
  const matchedCore = coreKeywords.find(k => cn.includes(k));
  if (matchedCore) {
    const candidate = allFoods.find(f => f.name.includes(matchedCore));
    if (candidate) return candidate;
  }

  // 7. Check 2-character sub-tokens
  for (let i = 0; i < cn.length - 1; i++) {
    const sub = cn.substring(i, i + 2);
    if (['的', '個', '份', '種', '類'].includes(sub[0]) || ['的', '個', '份', '種', '類'].includes(sub[1])) continue;
    const candidate = allFoods.find(f => f.name.includes(sub));
    if (candidate) return candidate;
  }

  // DO NOT invent foods. If no food matches in the database, return null!
  return null;
};

// Exclusively compare against FOOD_DATABASE from src/constants.ts
export const parseDietIntakeNotes = (
  text: string,
  defaultMeal: string = '早餐'
): DietSuggestion[] => {
  const allFoods = FOOD_DATABASE;
  if (!text || !text.trim() || !allFoods || allFoods.length === 0) return [];

  const lines = text.split(/\r?\n/);
  let currentMeal = defaultMeal;
  const results: DietSuggestion[] = [];
  const seenKeys = new Set<string>();

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;

    // Check if line starts with or is a meal header
    const mealMatch = line.match(/^\s*(?:【|\[)?(早餐|早點|中餐|午餐|下午茶|午點心|午點|點心|晚餐|晚點|宵夜|夜點)(?:】|\])?(?:\s*[:：\-—]\s*|\s+)?(.*)$/);

    let contentToProcess = line;

    if (mealMatch) {
      const detectedMealKeyword = mealMatch[1];
      const normalizedMeal = MEAL_NORM_MAP[detectedMealKeyword] || currentMeal;
      currentMeal = normalizedMeal;
      contentToProcess = mealMatch[2]?.trim() || '';
    }

    if (!contentToProcess) continue;

    // Split items by delimiters: +, ＋, 、, ，, ,, &, 與, 和, /, ;, tabs
    const rawSegments = contentToProcess.split(/[,，+＋、/|&與和;\t]+/);

    for (const seg of rawSegments) {
      if (!seg || !seg.trim()) continue;

      const { cleanName, qty } = parsePortion(seg);
      if (!cleanName || cleanName.length < 1) continue;

      // Skip common non-food noise words
      if (/^(的|了|吃|喝|喝了|吃了|加|配|有|無|沒|跟|大約|約|左右)$/.test(cleanName)) {
        continue;
      }

      // Find matching food ONLY in existing database
      const matchedFood = findBestFoodInDatabase(cleanName, allFoods);

      // If not in database, DO NOT grab or create it
      if (!matchedFood) {
        continue;
      }

      const key = `${currentMeal}__${matchedFood.name}`;
      if (!seenKeys.has(key)) {
        seenKeys.add(key);
        results.push({
          id: `${key}_${Math.random().toString(36).substr(2, 6)}`,
          meal: currentMeal,
          rawText: seg.trim(),
          food: matchedFood,
          qty: Math.max(0.1, qty),
        });
      }
    }
  }

  return results;
};
