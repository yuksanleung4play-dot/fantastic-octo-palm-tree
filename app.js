const sideHustles = [
  {
    id: "online-service",
    name: "線上服務",
    tagline: "以遠端交付為主的服務型副業，例如 AI 自動化、社群代操、線上助理。",
    timeText: "每週 6-15 小時；前 2 週集中建立服務包與銷售腳本。",
    moneyText: "NT$3,000-25,000；工具訂閱、作品集網站、廣告或名單開發。",
    revenueModel: "專案費、月費顧問、維運訂閱、成果分潤。",
    returnCycle: "快；約 1-4 週可取得第一筆成交。",
    scalabilityText: "中高；可產品化服務流程、外包執行或轉為模板與課程。",
    bestFor: "想快速驗證、具備溝通能力、願意主動開發客戶的人。",
    watchouts: ["交付品質會直接影響續約", "需要穩定開發名單", "服務範圍要避免失控"],
    nextMoves: [
      "定義一個 7 天可交付的入門服務包。",
      "整理 3 個假想案例或過往成果作為作品集。",
      "每天主動接觸 10 位目標客戶並追蹤回覆率。"
    ],
    scores: {
      cost: 86,
      time: 78,
      return: 90,
      scale: 72,
      cashflow: 88,
      brand: 68,
      flexibility: 78
    },
    minBudget: 3000,
    minHours: 6,
    returnSpeed: "fast",
    strengths: ["expertise", "operations"]
  },
  {
    id: "content-creation",
    name: "內容創作",
    tagline: "透過短影音、電子報、Podcast 或部落格累積信任與流量。",
    timeText: "每週 8-20 小時；需固定產出、剪輯、互動與數據回顧。",
    moneyText: "NT$2,000-40,000；設備、剪輯工具、設計素材或投放測試。",
    revenueModel: "廣告分潤、品牌合作、會員訂閱、導流至產品或服務。",
    returnCycle: "慢；通常 3-12 個月建立穩定變現基礎。",
    scalabilityText: "高；內容資產可長期導流並延伸社群、課程、商品。",
    bestFor: "願意長期累積個人品牌、具備觀點或教學能力的人。",
    watchouts: ["前期收入不穩", "演算法與平台規則會變動", "需要持續定位與內容實驗"],
    nextMoves: [
      "選定單一受眾與 3 個固定內容欄目。",
      "一次規劃 14 篇內容題目並建立發布節奏。",
      "設計電子報或社群入口，避免只依賴平台流量。"
    ],
    scores: {
      cost: 82,
      time: 62,
      return: 42,
      scale: 94,
      cashflow: 48,
      brand: 96,
      flexibility: 76
    },
    minBudget: 2000,
    minHours: 8,
    returnSpeed: "slow",
    strengths: ["audience", "expertise"]
  },
  {
    id: "ecommerce",
    name: "電商銷售",
    tagline: "銷售實體商品，可從選品、團購、代發貨或小批量庫存開始。",
    timeText: "每週 10-25 小時；選品、供應商、客服、物流與廣告優化。",
    moneyText: "NT$30,000-200,000；樣品、庫存、平台費、物流與廣告。",
    revenueModel: "商品毛利、組合包、回購品、訂閱補充包、團購佣金。",
    returnCycle: "中；約 1-3 個月看見有效銷售數據。",
    scalabilityText: "中高；可透過品牌化、供應鏈、廣告與通路放大。",
    bestFor: "擅長營運、數字追蹤、供應商談判且能承擔庫存風險的人。",
    watchouts: ["現金流與庫存壓力較高", "廣告成本會快速侵蝕毛利", "客服與退換貨流程要先設計"],
    nextMoves: [
      "鎖定一個高頻痛點與 10 個競品做價格帶分析。",
      "先用預購或小批量測試，不急著大量囤貨。",
      "建立毛利、退貨率與廣告投報的每日儀表板。"
    ],
    scores: {
      cost: 38,
      time: 50,
      return: 66,
      scale: 84,
      cashflow: 72,
      brand: 68,
      flexibility: 52
    },
    minBudget: 30000,
    minHours: 10,
    returnSpeed: "balanced",
    strengths: ["operations"]
  },
  {
    id: "professional-freelance",
    name: "專業接案",
    tagline: "把既有專業技能商品化，例如設計、工程、文案、顧問、財務或法務支援。",
    timeText: "每週 5-18 小時；前期重點是定位、作品集與成交流程。",
    moneyText: "NT$1,000-20,000；作品集、軟體工具、合約模板與開發名單。",
    revenueModel: "專案報價、時薪、顧問包、長約 retainer、績效獎金。",
    returnCycle: "最快；約 1-3 週有機會成交。",
    scalabilityText: "中；可提高單價、建立 SOP、轉顧問或產品化知識。",
    bestFor: "已有可出售技能、想用低資金快速創造現金流的人。",
    watchouts: ["收入與個人時間高度綁定", "需管理範疇與付款條款", "高單價需要清楚成果證明"],
    nextMoves: [
      "將技能包裝成 3 個固定方案：入門、標準、高階。",
      "寫出明確交付物、時程、修訂次數與付款條款。",
      "從前同事、社群與平台各取得 5 個潛在客戶對話。"
    ],
    scores: {
      cost: 94,
      time: 82,
      return: 96,
      scale: 62,
      cashflow: 94,
      brand: 74,
      flexibility: 82
    },
    minBudget: 1000,
    minHours: 5,
    returnSpeed: "fast",
    strengths: ["expertise"]
  },
  {
    id: "labor-service",
    name: "勞務服務",
    tagline: "以在地交付為主，例如清潔、寵物照護、居家整理、攝影、活動支援。",
    timeText: "每週 6-25 小時；排班、交通、現場服務與口碑經營。",
    moneyText: "NT$5,000-80,000；器材、保險、交通、制服與地區廣告。",
    revenueModel: "單次服務費、套票、固定週期訂閱、加購品或轉介紹獎勵。",
    returnCycle: "快；約 1-4 週可透過熟人與地區社群接單。",
    scalabilityText: "中；可標準化服務、訓練兼職人員、區域擴張。",
    bestFor: "願意現場服務、重視穩定需求、擁有本地人脈或交通彈性的人。",
    watchouts: ["體力與排班限制明顯", "品質控管與安全責任重要", "地理範圍影響成長速度"],
    nextMoves: [
      "先鎖定半徑 5 公里內的高需求客群。",
      "設計首購體驗價與轉介紹獎勵。",
      "建立服務前後照片、評價與標準檢查表。"
    ],
    scores: {
      cost: 68,
      time: 46,
      return: 82,
      scale: 54,
      cashflow: 82,
      brand: 54,
      flexibility: 42
    },
    minBudget: 5000,
    minHours: 6,
    returnSpeed: "fast",
    strengths: ["localNetwork", "operations"]
  },
  {
    id: "digital-products",
    name: "數位產品",
    tagline: "販售可重複交付的產品，例如模板、電子書、線上課程、資料包或工具。",
    timeText: "每週 8-18 小時；前期投入在產品設計、頁面、內容與銷售漏斗。",
    moneyText: "NT$1,000-50,000；工具、設計、素材、平台費與小額投放。",
    revenueModel: "一次性銷售、組合包、訂閱、授權、升級版或導流顧問。",
    returnCycle: "中；約 1-3 個月可完成 MVP 並驗證銷售。",
    scalabilityText: "很高；邊際交付成本低，可全球銷售與自動化漏斗。",
    bestFor: "有明確知識資產、想降低交付時間並追求長期擴展的人。",
    watchouts: ["需要先驗證痛點與願付價格", "流量不足會讓好產品賣不動", "客服與更新承諾要可控"],
    nextMoves: [
      "先訪談 10 位目標客戶確認痛點與願付價格。",
      "做一個 7 天內可完成的最小版本並開放預購。",
      "建立登陸頁、案例教學與 5 封銷售信件。"
    ],
    scores: {
      cost: 88,
      time: 70,
      return: 58,
      scale: 98,
      cashflow: 62,
      brand: 82,
      flexibility: 92
    },
    minBudget: 1000,
    minHours: 8,
    returnSpeed: "balanced",
    strengths: ["expertise", "audience"]
  }
];

const elements = {
  comparisonBody: document.querySelector("#comparisonBody"),
  categoryCards: document.querySelector("#categoryCards"),
  categoryFilter: document.querySelector("#categoryFilter"),
  sortBy: document.querySelector("#sortBy"),
  selectedInsight: document.querySelector("#selectedInsight"),
  goalForm: document.querySelector("#goalForm"),
  monthlyGoal: document.querySelector("#monthlyGoal"),
  weeklyHours: document.querySelector("#weeklyHours"),
  startupBudget: document.querySelector("#startupBudget"),
  returnSpeed: document.querySelector("#returnSpeed"),
  priority: document.querySelector("#priority"),
  topPick: document.querySelector("#topPick"),
  topSummary: document.querySelector("#topSummary"),
  scoreBar: document.querySelector("#scoreBar"),
  scoreText: document.querySelector("#scoreText"),
  fitReasons: document.querySelector("#fitReasons"),
  nextMoves: document.querySelector("#nextMoves"),
  rankingList: document.querySelector("#rankingList")
};

let selectedCategoryId = sideHustles[0].id;

function getSelectedStrengths() {
  return Array.from(document.querySelectorAll('input[name="strengths"]:checked')).map(
    (input) => input.value
  );
}

function getGoals() {
  return {
    monthlyGoal: Number(elements.monthlyGoal.value),
    weeklyHours: Number(elements.weeklyHours.value),
    startupBudget: Number(elements.startupBudget.value),
    returnSpeed: elements.returnSpeed.value,
    priority: elements.priority.value,
    strengths: getSelectedStrengths()
  };
}

function normalizeGoalAmbition(monthlyGoal) {
  if (monthlyGoal <= 30000) return 10;
  if (monthlyGoal <= 80000) return 4;
  if (monthlyGoal <= 150000) return -2;
  return -8;
}

function calculateFitScore(item, goals) {
  const base =
    item.scores.cost * 0.16 +
    item.scores.time * 0.14 +
    item.scores.return * 0.18 +
    item.scores.scale * 0.18 +
    item.scores[goals.priority] * 0.24 +
    item.scores.flexibility * 0.1;

  const budgetGap = goals.startupBudget >= item.minBudget ? 10 : -18;
  const timeGap = goals.weeklyHours >= item.minHours ? 10 : -16;
  const speedFit =
    goals.returnSpeed === item.returnSpeed
      ? 10
      : goals.returnSpeed === "balanced" && item.returnSpeed !== "slow"
        ? 4
        : goals.returnSpeed === "slow" && item.scores.scale >= 85
          ? 8
          : -6;
  const strengthFit = item.strengths.filter((strength) => goals.strengths.includes(strength)).length * 7;
  const ambitionFit =
    goals.monthlyGoal >= 120000 && item.scores.scale >= 84
      ? 8
      : goals.monthlyGoal <= 50000 && item.scores.return >= 80
        ? 6
        : normalizeGoalAmbition(goals.monthlyGoal);

  return Math.max(0, Math.min(100, Math.round(base + budgetGap + timeGap + speedFit + strengthFit + ambitionFit - 18)));
}

function getRankedItems() {
  const goals = getGoals();
  return sideHustles
    .map((item) => ({ ...item, fitScore: calculateFitScore(item, goals) }))
    .sort((a, b) => b.fitScore - a.fitScore);
}

function getSortedItems() {
  const sortBy = elements.sortBy.value;
  const ranked = getRankedItems();
  const scoreMap = {
    fit: "fitScore",
    cost: "cost",
    return: "return",
    scale: "scale",
    time: "time"
  };

  return ranked.sort((a, b) => {
    const key = scoreMap[sortBy];
    const left = key === "fitScore" ? a.fitScore : a.scores[key];
    const right = key === "fitScore" ? b.fitScore : b.scores[key];
    return right - left;
  });
}

function filterItems(items) {
  const filter = elements.categoryFilter.value;
  const checks = {
    all: () => true,
    lowBudget: (item) => item.minBudget <= 5000,
    fastReturn: (item) => item.scores.return >= 80,
    highScale: (item) => item.scores.scale >= 84,
    lowTime: (item) => item.minHours <= 6
  };

  return items.filter(checks[filter]);
}

function renderComparison() {
  const items = filterItems(getSortedItems());
  elements.comparisonBody.innerHTML = items
    .map(
      (item) => `
        <tr tabindex="0" data-id="${item.id}">
          <td><strong>${item.name}</strong>${item.tagline}</td>
          <td>${item.timeText}</td>
          <td>${item.moneyText}</td>
          <td>${item.revenueModel}</td>
          <td>${item.returnCycle}</td>
          <td>${item.scalabilityText}</td>
          <td><span class="pill">${item.fitScore}</span></td>
        </tr>
      `
    )
    .join("");
}

function renderCards() {
  const items = filterItems(getSortedItems());
  elements.categoryCards.innerHTML = items
    .map(
      (item) => `
        <article class="category-card ${item.id === selectedCategoryId ? "is-selected" : ""}" data-id="${item.id}">
          <div>
            <p class="eyebrow">${item.fitScore} / 100</p>
            <h3>${item.name}</h3>
          </div>
          <p>${item.bestFor}</p>
          <ul>
            ${item.watchouts.map((watchout) => `<li>${watchout}</li>`).join("")}
          </ul>
        </article>
      `
    )
    .join("");
}

function renderSelectedInsight() {
  const rankedItems = getRankedItems();
  const selectedItem =
    rankedItems.find((item) => item.id === selectedCategoryId) || rankedItems[0];

  elements.selectedInsight.innerHTML = `
    <div>
      <p class="eyebrow">Selected Detail</p>
      <h3>${selectedItem.name}詳細分析</h3>
      <p>${selectedItem.tagline}</p>
    </div>
    <div class="selected-insight__grid">
      <div class="selected-insight__item">
        <strong>時間／資金投入</strong>
        <p>${selectedItem.timeText}</p>
        <p>${selectedItem.moneyText}</p>
      </div>
      <div class="selected-insight__item">
        <strong>營收與回報</strong>
        <p>${selectedItem.revenueModel}</p>
        <p>${selectedItem.returnCycle}</p>
      </div>
      <div class="selected-insight__item">
        <strong>擴展與行動</strong>
        <p>${selectedItem.scalabilityText}</p>
        <p>${selectedItem.nextMoves[0]}</p>
      </div>
    </div>
  `;
}

function buildReasons(item, goals) {
  const reasons = [];

  if (goals.startupBudget >= item.minBudget) {
    reasons.push(`啟動資金門檻符合目前預算，最低約 NT$${item.minBudget.toLocaleString("zh-TW")} 起。`);
  } else {
    reasons.push("目前預算低於建議門檻，若選此方向需先用預售、合作或縮小版本降低風險。");
  }

  if (goals.weeklyHours >= item.minHours) {
    reasons.push(`每週可投入時間足以支撐初期啟動，建議至少保留 ${item.minHours} 小時。`);
  } else {
    reasons.push("時間投入偏緊，應優先選擇更小的產品範圍或提高價格以避免過度交付。");
  }

  if (item.scores[goals.priority] >= 80) {
    reasons.push("此類型高度符合你目前最重視的成果指標。");
  }

  if (item.strengths.some((strength) => goals.strengths.includes(strength))) {
    reasons.push("你的既有優勢能直接降低冷啟動難度。");
  }

  if (goals.monthlyGoal >= 120000 && item.scores.scale >= 84) {
    reasons.push("收入目標較高，此方向具備較好的規模化與長期放大空間。");
  }

  return reasons;
}

function renderRecommendation() {
  const goals = getGoals();
  const ranked = getRankedItems();
  const top = ranked[0];

  selectedCategoryId = selectedCategoryId || top.id;
  elements.topPick.textContent = top.name;
  elements.topSummary.textContent = `${top.tagline} ${top.bestFor}`;
  elements.scoreBar.style.width = `${top.fitScore}%`;
  elements.scoreText.textContent = `推薦分數 ${top.fitScore} / 100`;
  elements.fitReasons.innerHTML = buildReasons(top, goals)
    .map((reason) => `<li>${reason}</li>`)
    .join("");
  elements.nextMoves.innerHTML = top.nextMoves.map((move) => `<li>${move}</li>`).join("");

  elements.rankingList.innerHTML = ranked
    .map(
      (item, index) => `
        <div class="ranking__item">
          <span class="ranking__position">${index + 1}</span>
          <div>
            <strong>${item.name}</strong>
            <div class="ranking__bar" aria-hidden="true"><span style="width:${item.fitScore}%"></span></div>
          </div>
          <span class="pill">${item.fitScore}</span>
        </div>
      `
    )
    .join("");
}

function render() {
  renderComparison();
  renderCards();
  renderSelectedInsight();
  renderRecommendation();
}

function selectCategory(categoryId) {
  selectedCategoryId = categoryId;
  renderCards();
  renderSelectedInsight();
  elements.selectedInsight.scrollIntoView({ behavior: "smooth", block: "center" });
}

document.addEventListener("click", (event) => {
  const row = event.target.closest("[data-id]");
  if (row) selectCategory(row.dataset.id);
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    const row = event.target.closest("tr[data-id]");
    if (row) selectCategory(row.dataset.id);
  }
});

elements.goalForm.addEventListener("input", render);
elements.categoryFilter.addEventListener("change", render);
elements.sortBy.addEventListener("change", render);

render();
