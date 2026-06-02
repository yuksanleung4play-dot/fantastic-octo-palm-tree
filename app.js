(function () {
  "use strict";

  var state = { district: "seongsu", filter: "全部", activeShop: null };

  var el = {
    tabs: document.getElementById("tabs"),
    name: document.getElementById("district-name"),
    tagline: document.getElementById("district-tagline"),
    intro: document.getElementById("district-intro"),
    station: document.getElementById("district-station"),
    time: document.getElementById("district-time"),
    svg: document.getElementById("map-svg"),
    routeNote: document.getElementById("route-note"),
    routeList: document.getElementById("route-list"),
    chips: document.getElementById("filter-chips"),
    cards: document.getElementById("cards"),
  };

  var SVGNS = "http://www.w3.org/2000/svg";
  var TIER_LABEL = { 1: "₩", 2: "₩₩", 3: "₩₩₩" };

  function svgEl(tag, attrs) {
    var node = document.createElementNS(SVGNS, tag);
    for (var k in attrs) node.setAttribute(k, attrs[k]);
    return node;
  }

  function dist(a, b) {
    return Math.sqrt(Math.pow(a.x - b.x, 2) + Math.pow(a.y - b.y, 2));
  }

  // 依示意座標估算步行分鐘數（純參考，非實際距離）
  function walkMinutes(a, b) {
    return Math.max(3, Math.round(dist(a, b) / 55) + 2);
  }

  function getShopsSorted(d) {
    return d.shops.slice().sort(function (x, y) { return x.order - y.order; });
  }

  /* ---------- 區域簡介 ---------- */
  function renderIntro(d) {
    el.name.textContent = d.name + "　" + d.nameKo;
    el.tagline.textContent = d.tagline;
    el.intro.textContent = d.intro;
    el.station.textContent = d.nearestStation;
    el.time.textContent = d.bestTime;
  }

  /* ---------- SVG 地圖 ---------- */
  function renderMap(d) {
    el.svg.innerHTML = "";
    var shops = getShopsSorted(d);

    // 街道
    d.mapStreets.forEach(function (s) {
      el.svg.appendChild(svgEl("line", {
        x1: s.x1, y1: s.y1, x2: s.x2, y2: s.y2, class: "street-line",
      }));
      var t = svgEl("text", { x: (s.x1 + s.x2) / 2, y: (s.y1 + s.y2) / 2 - 12, class: "street-label" });
      t.textContent = s.label;
      el.svg.appendChild(t);
    });

    // 路線虛線
    var pts = shops.map(function (s) { return s.x + "," + s.y; }).join(" ");
    el.svg.appendChild(svgEl("polyline", { points: pts, class: "route-path" }));

    // 標記點
    shops.forEach(function (s) {
      var g = svgEl("g", { class: "marker", "data-id": s.id });
      g.appendChild(svgEl("circle", { cx: s.x, cy: s.y, r: 18 }));
      var num = svgEl("text", { x: s.x, y: s.y });
      num.textContent = s.order;
      g.appendChild(num);
      var anchor = s.x > 780 ? "end" : "start";
      var nx = s.x > 780 ? s.x - 24 : s.x + 24;
      var nm = svgEl("text", { x: nx, y: s.y + 5, class: "marker-name", "text-anchor": anchor });
      nm.textContent = s.name;
      g.appendChild(nm);

      g.addEventListener("click", function () { focusShop(s.id); });
      g.addEventListener("mouseenter", function () { highlight(s.id, false); });
      g.addEventListener("mouseleave", function () { if (state.activeShop !== s.id) clearHighlight(); });
      el.svg.appendChild(g);
    });
  }

  /* ---------- 步行路線 ---------- */
  function renderRoute(d) {
    el.routeNote.textContent = d.routeNote;
    el.routeList.innerHTML = "";
    var shops = getShopsSorted(d);
    shops.forEach(function (s, i) {
      var li = document.createElement("li");
      var walk = i === 0 ? "由最近車站起步" : "步行約 " + walkMinutes(shops[i - 1], s) + " 分鐘";

      var nameSpan = document.createElement("span");
      nameSpan.className = "r-name";
      nameSpan.textContent = s.name;

      var walkSpan = document.createElement("span");
      walkSpan.className = "r-walk";
      walkSpan.textContent = "↳ " + walk;

      li.appendChild(nameSpan);
      li.appendChild(walkSpan);
      li.style.cursor = "pointer";
      li.addEventListener("click", function () { focusShop(s.id); });
      el.routeList.appendChild(li);
    });
  }

  /* ---------- 篩選 chips ---------- */
  function renderChips(d) {
    el.chips.innerHTML = "";
    var cats = ["全部"];
    d.shops.forEach(function (s) { if (cats.indexOf(s.category) === -1) cats.push(s.category); });
    cats.forEach(function (c) {
      var b = document.createElement("button");
      b.className = "chip" + (c === state.filter ? " is-active" : "");
      b.textContent = c;
      b.addEventListener("click", function () {
        state.filter = c;
        renderChips(d);
        renderCards(d);
      });
      el.chips.appendChild(b);
    });
  }

  /* ---------- 卡片 ---------- */
  function renderCards(d) {
    el.cards.innerHTML = "";
    var shops = getShopsSorted(d).filter(function (s) {
      return state.filter === "全部" || s.category === state.filter;
    });
    shops.forEach(function (s) {
      var card = document.createElement("article");
      card.className = "card fade-in";
      card.id = "card-" + s.id;
      card.dataset.id = s.id;

      var tags = s.highlights.map(function (h) { return "<span>" + h + "</span>"; }).join("");

      card.innerHTML =
        '<div class="card__top">' +
          '<span class="card__num">' + s.order + "</span>" +
          "<div>" +
            '<h4 class="card__title">' + s.name + "</h4>" +
            '<p class="card__title-ko">' + s.nameKo + "</p>" +
          "</div>" +
          '<span class="card__cat">' + s.category + "</span>" +
        "</div>" +
        '<p class="card__style">' + s.style + "</p>" +
        '<div class="card__tags">' + tags + "</div>" +
        '<div class="card__info">' +
          row("📍", "地址", s.address) +
          row("🕒", "營業", s.hours) +
          row("💰", "預算", s.budget + '　<span class="price-tier">' + TIER_LABEL[s.priceTier] + "</span>") +
        "</div>";

      card.addEventListener("mouseenter", function () { highlight(s.id, false); });
      card.addEventListener("mouseleave", function () { if (state.activeShop !== s.id) clearHighlight(); });
      el.cards.appendChild(card);
    });
  }

  function row(ico, k, v) {
    return '<div class="card__row"><span class="ico">' + ico + '</span><span class="k">' + k + '</span><span class="v">' + v + "</span></div>";
  }

  /* ---------- 互動：高亮 / 聚焦 ---------- */
  function highlight(id, scroll) {
    clearHighlight();
    var marker = el.svg.querySelector('.marker[data-id="' + id + '"]');
    if (marker) marker.classList.add("is-active");
    var card = document.getElementById("card-" + id);
    if (card) {
      card.classList.add("is-active");
      if (scroll) card.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }

  function clearHighlight() {
    var prev = el.svg.querySelectorAll(".marker.is-active");
    Array.prototype.forEach.call(prev, function (n) { n.classList.remove("is-active"); });
    var cards = el.cards.querySelectorAll(".card.is-active");
    Array.prototype.forEach.call(cards, function (n) { n.classList.remove("is-active"); });
  }

  function focusShop(id) {
    if (!document.getElementById("card-" + id)) {
      state.filter = "全部";
      renderChips(DISTRICTS[state.district]);
      renderCards(DISTRICTS[state.district]);
    }
    state.activeShop = id;
    highlight(id, true);
  }

  /* ---------- 切換區域 ---------- */
  function setDistrict(key) {
    state.district = key;
    state.filter = "全部";
    state.activeShop = null;
    var d = DISTRICTS[key];

    Array.prototype.forEach.call(el.tabs.querySelectorAll(".tab"), function (t) {
      t.classList.toggle("is-active", t.dataset.district === key);
    });

    renderIntro(d);
    renderMap(d);
    renderRoute(d);
    renderChips(d);
    renderCards(d);
  }

  /* ---------- 初始化 ---------- */
  Array.prototype.forEach.call(el.tabs.querySelectorAll(".tab"), function (t) {
    t.addEventListener("click", function () { setDistrict(t.dataset.district); });
  });

  setDistrict("seongsu");
})();
