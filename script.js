/* 日田AEDナビ - 日田市内の最寄りAEDへの徒歩最短経路
 * データ: AED.xlsx (シート「AED」/ 無ければ先頭シート)
 * 地図: Leaflet + OpenStreetMap / 経路: routing.openstreetmap.de (徒歩)
 */
'use strict';

// ====== 設定 ======
const CONFIG = {
  xlsxPath: 'AED.xlsx',
  // ---- 日田市(大分県)向け設定 ----
  cityName: '日田市',
  defaultCenter: [33.3216, 130.9419], // 初期表示位置: JR日田駅付近
  defaultZoom: 14,
  // 日田市全域を含む概算の矩形 [[南, 西], [北, 東]] (市境そのものではありません)
  cityBounds: [[33.02, 130.70], [33.50, 131.30]],
  // 現在地が取れないとき用の「出発地」候補 (概算位置。必要に応じて修正してください)
  startPoints: [
    { label: 'JR日田駅付近', lat: 33.3216, lon: 130.9419 },
    { label: '豆田町付近',   lat: 33.3205, lon: 130.9447 },
    { label: '日田市役所付近', lat: 33.3220, lon: 130.9411 },
  ],
  candidateCount: 5,                  // 直線距離の近い順に、実際の徒歩経路を調べる件数
  walkSpeedMPerMin: 80,               // 徒歩速度 (m/分) ※一般的な目安
  routeUrl: 'https://routing.openstreetmap.de/routed-foot/route/v1/foot/',
  routeTimeoutMs: 8000,
};

// 列見出しの別名 (Excel側の表記ゆれを吸収)
const COLUMN_ALIASES = {
  no:       ['no', 'No', 'NO', '番号'],
  name:     ['施設名', '名称', '場所名', 'name', 'Name'],
  address:  ['住所', '所在地', 'address', 'Address'],
  lat:      ['緯度', 'lat', 'Lat', 'latitude', 'Latitude'],
  lon:      ['経度', 'lon', 'lng', 'Lon', 'Lng', 'longitude', 'Longitude'],
  place:    ['設置場所', '設置位置', '場所', 'place'],
  hours:    ['利用可能時間', '利用時間', '設置時間', 'hours'],
  note:     ['備考', 'メモ', 'note'],
};

// ====== 状態 ======
const state = {
  map: null,
  aeds: [],
  user: null,          // {lat, lon}
  userMarker: null,
  aedMarkers: new Map(), // aed.id -> marker
  routeLayer: null,
  bestId: null,
  searching: false,
};

// ====== DOM ======
const $ = (id) => document.getElementById(id);
const el = {
  locate: $('btn-locate'),
  search: $('btn-search'),
  chkAvail: $('chk-available'),
  status: $('status'),
  result: $('result'),
  listWrap: $('list-wrap'),
  list: $('nearby-list'),
  startPoints: $('start-points'),
  fallback: $('fallback'),
  fileInput: $('file-input'),
};

// ====== ユーティリティ ======
function setStatus(msg, kind = '') {
  el.status.textContent = msg;
  el.status.className = 'status' + (kind ? ' ' + kind : '');
}

// ハバーシン公式による2点間の直線距離 (m)
function haversine(lat1, lon1, lat2, lon2) {
  const R = 6371008.8;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

function formatDistance(m) {
  return m < 1000 ? `${Math.round(m)} m` : `${(m / 1000).toFixed(2)} km`;
}

function formatMinutes(min) {
  const r = Math.max(1, Math.round(min));
  if (r < 60) return `${r} 分`;
  return `${Math.floor(r / 60)} 時間 ${r % 60} 分`;
}

function pad2(n) { return String(n).padStart(2, '0'); }
function formatClock(d) { return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; }

function withTimeout(promise, ms) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error('timeout')), ms);
    promise.then((v) => { clearTimeout(t); resolve(v); },
                 (e) => { clearTimeout(t); reject(e); });
  });
}

// ====== 利用可能時間の判定 ======
// 「24時間」/ 空欄 → 常時可。「09:00-17:00」「9:00～17:00」形式に対応 (日またぎも可)
function parseHours(text) {
  const s = String(text ?? '').trim()
    .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[：]/g, ':').replace(/[〜～~－―ー–—]/g, '-');
  if (s === '' || /24\s*時間|常時|終日|24h/i.test(s)) return { always: true };
  const m = s.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  if (!m) return { always: true, unknown: true }; // 解釈できない場合は制限なし扱い
  return { always: false, start: (+m[1]) * 60 + (+m[2]), end: (+m[3]) * 60 + (+m[4]) };
}

function isAvailableAt(aed, date) {
  const h = aed.hoursParsed;
  if (h.always) return true;
  const t = date.getHours() * 60 + date.getMinutes();
  if (h.start === h.end) return true;
  return h.start < h.end ? (t >= h.start && t < h.end) : (t >= h.start || t < h.end);
}

// ====== Excel 読み込み ======
function pick(row, key) {
  for (const alias of COLUMN_ALIASES[key]) {
    if (row[alias] !== undefined && row[alias] !== '') return row[alias];
  }
  return '';
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = wb.SheetNames.includes('AED') ? 'AED' : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });

  const aeds = [];
  let skipped = 0;
  rows.forEach((row, i) => {
    const lat = parseFloat(pick(row, 'lat'));
    const lon = parseFloat(pick(row, 'lon'));
    const name = String(pick(row, 'name')).trim();
    const valid = Number.isFinite(lat) && Number.isFinite(lon) &&
                  lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
    if (!valid || !name) { skipped++; return; }
    const hours = String(pick(row, 'hours')).trim();
    aeds.push({
      id: aeds.length,
      no: pick(row, 'no') || i + 1,
      name,
      address: String(pick(row, 'address')).trim(),
      lat, lon,
      place: String(pick(row, 'place')).trim(),
      hours: hours || '24時間',
      hoursParsed: parseHours(hours),
      note: String(pick(row, 'note')).trim(),
    });
  });
  return { aeds, skipped, sheetName };
}

async function loadAedData() {
  try {
    const res = await fetch(CONFIG.xlsxPath, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    applyAedData(parseWorkbook(await res.arrayBuffer()));
  } catch (err) {
    console.error(err);
    el.fallback.hidden = false;
    el.fallback.open = true;
    setStatus('AED.xlsx を自動で読み込めませんでした。下のボタンからファイルを選択してください。', 'error');
  }
}

function applyAedData({ aeds, skipped }) {
  if (aeds.length === 0) {
    setStatus('AEDデータが見つかりません。AED.xlsx の列見出し(施設名/緯度/経度)を確認してください。', 'error');
    return;
  }
  state.aeds = aeds;
  renderAedMarkers();
  const bounds = L.latLngBounds(aeds.map((a) => [a.lat, a.lon]));
  state.map.fitBounds(bounds.pad(0.15));
  el.search.disabled = !state.user;
  const extra = skipped ? `(座標や名称が不正な ${skipped} 行は除外)` : '';
  setStatus(`AED ${aeds.length} 件を読み込みました${extra}。「現在地を取得」を押してください。`, 'ok');
  if (state.user) updateNearbyList();
}

// ====== 地図 ======
function makeIcon(cls, label) {
  return L.divIcon({
    className: '',
    html: `<div class="pin ${cls}"><span>${label}</span></div>`,
    iconSize: [30, 30],
    iconAnchor: [15, 30],
    popupAnchor: [0, -28],
  });
}

function initMap() {
  // 日田市周辺に表示範囲を制限
  state.map = L.map('map', {
    zoomControl: true,
    minZoom: 10,
    maxBounds: L.latLngBounds(CONFIG.cityBounds).pad(0.3),
    maxBoundsViscosity: 0.8,
  }).setView(CONFIG.defaultCenter, CONFIG.defaultZoom);
  // 日田市の範囲を薄く表示
  L.rectangle(CONFIG.cityBounds, { color: '#c0392b', weight: 1, fill: false, dashArray: '4 6', interactive: false })
    .addTo(state.map);
  L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
  }).addTo(state.map);

  // 地図のクリック/タップで現在地を手動指定
  state.map.on('click', (e) => setUserPosition(e.latlng.lat, e.latlng.lng, '地図で指定した地点'));
}

function buildPopup(aed) {
  const wrap = document.createElement('div');
  const title = document.createElement('strong');
  title.textContent = aed.name;
  wrap.appendChild(title);
  const lines = [aed.place && `設置場所: ${aed.place}`, `利用可能時間: ${aed.hours}`, aed.address];
  lines.filter(Boolean).forEach((t) => {
    const p = document.createElement('div');
    p.textContent = t;
    wrap.appendChild(p);
  });
  return wrap;
}

function renderAedMarkers() {
  state.aedMarkers.forEach((m) => m.remove());
  state.aedMarkers.clear();
  const now = new Date();
  state.aeds.forEach((aed) => {
    const open = isAvailableAt(aed, now);
    const marker = L.marker([aed.lat, aed.lon], {
      icon: makeIcon('aed' + (open ? '' : ' closed'), '+'),
      title: aed.name,
    }).addTo(state.map);
    marker.bindPopup(buildPopup(aed));
    state.aedMarkers.set(aed.id, marker);
  });
}

function highlightBest(id) {
  const now = new Date();
  state.aedMarkers.forEach((marker, aedId) => {
    const aed = state.aeds[aedId];
    const open = isAvailableAt(aed, now);
    const best = aedId === id;
    marker.setIcon(makeIcon('aed' + (open ? '' : ' closed') + (best ? ' best' : ''), best ? '★' : '+'));
    marker.setZIndexOffset(best ? 1000 : 0);
  });
  state.bestId = id;
}

function clearRoute() {
  if (state.routeLayer) { state.routeLayer.remove(); state.routeLayer = null; }
  el.result.hidden = true;
  if (state.bestId !== null) highlightBest(null);
}

// ====== 現在地 ======
function isInCity(lat, lon) {
  const [[s, w], [n, e]] = CONFIG.cityBounds;
  return lat >= s && lat <= n && lon >= w && lon <= e;
}

function setUserPosition(lat, lon, label) {
  if (!isInCity(lat, lon)) {
    // 日田市外: 地図の範囲外で位置が見えなくなるのを防ぐため、設定せず案内する
    setStatus(`${label}は${CONFIG.cityName}の範囲外です。${CONFIG.cityName}内の地点を地図で指定するか、下の出発地から選んでください。`, 'warn');
    return;
  }
  state.user = { lat, lon };
  if (state.userMarker) state.userMarker.remove();
  state.userMarker = L.marker([lat, lon], { icon: makeIcon('me', '●'), zIndexOffset: 2000 })
    .addTo(state.map).bindPopup('現在地');
  clearRoute();
  el.search.disabled = state.aeds.length === 0;
  updateNearbyList();
  if (state.aeds.length) {
    setStatus(`${label}を現在地にしました。「最短のAEDへの経路を表示」を押してください。`, 'ok');
  } else {
    setStatus(`${label}を現在地にしました。AEDデータを読み込み中です。`);
  }
}

function locate() {
  if (!('geolocation' in navigator)) {
    setStatus('この端末では位置情報を取得できません。地図をタップして現在地を指定してください。', 'warn');
    return;
  }
  setStatus('現在地を取得しています…');
  el.locate.disabled = true;
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      el.locate.disabled = false;
      const { latitude, longitude } = pos.coords;
      setUserPosition(latitude, longitude, '現在地');
      state.map.setView([latitude, longitude], Math.max(state.map.getZoom(), 16));
    },
    (err) => {
      el.locate.disabled = false;
      const msg = err.code === 1
        ? '位置情報の利用が許可されていません。ブラウザの設定を確認するか、地図をタップして現在地を指定してください。'
        : '現在地を取得できませんでした。地図をタップして現在地を指定してください。';
      setStatus(msg, 'warn');
    },
    { enableHighAccuracy: true, timeout: 12000, maximumAge: 10000 }
  );
}

// ====== 近くのAED一覧 ======
function withStraightDistance() {
  const { lat, lon } = state.user;
  return state.aeds
    .map((aed) => ({ aed, straight: haversine(lat, lon, aed.lat, aed.lon) }))
    .sort((a, b) => a.straight - b.straight);
}

function updateNearbyList() {
  if (!state.user || !state.aeds.length) { el.listWrap.hidden = true; return; }
  const now = new Date();
  const items = withStraightDistance().slice(0, 5);
  el.list.replaceChildren();
  items.forEach(({ aed, straight }) => {
    const li = document.createElement('li');
    const nm = document.createElement('span');
    nm.className = 'nm';
    nm.textContent = aed.name;
    const sub = document.createElement('span');
    sub.className = 'sub';
    const open = isAvailableAt(aed, now) ? '利用可' : '時間外';
    sub.textContent = `直線 ${formatDistance(straight)} ・ ${aed.hours}(${open})`;
    li.append(nm, sub);
    li.addEventListener('click', () => {
      state.map.setView([aed.lat, aed.lon], 17);
      state.aedMarkers.get(aed.id).openPopup();
    });
    el.list.appendChild(li);
  });
  el.listWrap.hidden = false;
}

// ====== 経路探索 ======
async function fetchWalkingRoute(from, to) {
  const url = `${CONFIG.routeUrl}${from.lon},${from.lat};${to.lon},${to.lat}` +
              `?overview=full&geometries=geojson`;
  const res = await withTimeout(fetch(url), CONFIG.routeTimeoutMs);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  if (data.code !== 'Ok' || !data.routes || !data.routes.length) throw new Error(data.code || 'no route');
  const r = data.routes[0];
  return {
    distance: r.distance,
    // GeoJSON は [lon, lat] なので Leaflet 用に [lat, lon] へ
    latlngs: r.geometry.coordinates.map(([lo, la]) => [la, lo]),
    estimated: false,
  };
}

// 経路APIが使えない場合の代替: 直線 (距離は直線の1.3倍で概算)
function fallbackRoute(from, to) {
  const d = haversine(from.lat, from.lon, to.lat, to.lon);
  return { distance: d * 1.3, latlngs: [[from.lat, from.lon], [to.lat, to.lon]], estimated: true };
}

async function findShortest() {
  if (state.searching || !state.user || !state.aeds.length) return;
  state.searching = true;
  el.search.disabled = true;
  clearRoute();
  setStatus('最短のAEDを探しています…');

  try {
    const now = new Date();
    const onlyAvailable = el.chkAvail.checked;
    const sorted = withStraightDistance();

    // 候補: 現在利用可能なものを優先 (チェック時は利用可能のみ)
    let pool = onlyAvailable ? sorted.filter((x) => isAvailableAt(x.aed, now)) : sorted;
    if (pool.length === 0) {
      setStatus('現在利用できるAEDが見つかりません。「今利用できるAEDのみ」のチェックを外して再検索できます。', 'warn');
      return;
    }
    const candidates = pool.slice(0, CONFIG.candidateCount);

    // 候補すべての徒歩経路を並列で取得し、実際の徒歩距離が最短のものを選ぶ
    const settled = await Promise.all(candidates.map(async ({ aed, straight }) => {
      try {
        return { aed, straight, route: await fetchWalkingRoute(state.user, aed) };
      } catch (e) {
        console.warn('route failed:', aed.name, e);
        return { aed, straight, route: null };
      }
    }));

    let usable = settled.filter((x) => x.route);
    let apiFailed = false;
    if (usable.length === 0) {
      apiFailed = true;
      usable = settled.map((x) => ({ ...x, route: fallbackRoute(state.user, x.aed) }));
    }

    // 到着時刻に利用可能かも考慮 (時間内のものを優先)
    usable.forEach((x) => {
      x.minutes = x.route.distance / CONFIG.walkSpeedMPerMin;
      x.arrival = new Date(now.getTime() + x.minutes * 60000);
      x.openOnArrival = isAvailableAt(x.aed, x.arrival);
    });
    const preferred = usable.filter((x) => x.openOnArrival);
    const finalists = preferred.length ? preferred : usable;
    finalists.sort((a, b) => a.route.distance - b.route.distance);
    const best = finalists[0];

    showResult(best, { apiFailed, closedOnArrival: !best.openOnArrival });
  } catch (err) {
    console.error(err);
    setStatus('検索中にエラーが発生しました。もう一度お試しください。', 'error');
  } finally {
    state.searching = false;
    el.search.disabled = false;
  }
}

// ====== 結果表示 ======
function addDt(dl, term, desc) {
  const dt = document.createElement('dt'); dt.textContent = term;
  const dd = document.createElement('dd'); dd.textContent = desc;
  dl.append(dt, dd);
}

function showResult(best, { apiFailed, closedOnArrival }) {
  const { aed, route, minutes, arrival, openOnArrival } = best;

  if (state.routeLayer) state.routeLayer.remove();
  state.routeLayer = L.polyline(route.latlngs, {
    color: route.estimated ? '#7a8794' : '#c0392b',
    weight: 6, opacity: 0.85,
    dashArray: route.estimated ? '8 10' : null,
  }).addTo(state.map);
  highlightBest(aed.id);
  state.map.fitBounds(state.routeLayer.getBounds().pad(0.25));

  // 結果カード (ユーザー/Excel由来の文字列は textContent で挿入)
  el.result.replaceChildren();

  const h2 = document.createElement('h2');
  h2.textContent = `最短: ${aed.name}`;
  const badge = document.createElement('span');
  badge.className = 'badge ' + (openOnArrival ? 'open' : 'closed');
  badge.textContent = openOnArrival ? '到着時に利用可' : '到着時は時間外の可能性';
  h2.append(' ', badge);
  el.result.appendChild(h2);

  if (aed.address) {
    const p = document.createElement('p');
    p.className = 'addr';
    p.textContent = aed.address;
    el.result.appendChild(p);
  }

  const metrics = document.createElement('div');
  metrics.className = 'metrics';
  [[(route.estimated ? '約 ' : '') + formatDistance(route.distance), '徒歩距離'],
   [(route.estimated ? '約 ' : '') + formatMinutes(minutes), '徒歩時間の目安']].forEach(([v, l]) => {
    const box = document.createElement('div'); box.className = 'metric';
    const vv = document.createElement('span'); vv.className = 'v'; vv.textContent = v;
    const ll = document.createElement('span'); ll.className = 'l'; ll.textContent = l;
    box.append(vv, ll);
    metrics.appendChild(box);
  });
  el.result.appendChild(metrics);

  const dl = document.createElement('dl');
  if (aed.place) addDt(dl, '設置場所', aed.place);
  addDt(dl, '利用可能時間', `${aed.hours}(到着予定 ${formatClock(arrival)} 頃)`);
  if (aed.note) addDt(dl, '備考', aed.note);
  el.result.appendChild(dl);

  const a = document.createElement('a');
  a.className = 'gmap-link';
  a.target = '_blank';
  a.rel = 'noopener';
  a.href = 'https://www.google.com/maps/dir/?api=1' +
    `&origin=${state.user.lat},${state.user.lon}` +
    `&destination=${aed.lat},${aed.lon}&travelmode=walking`;
  a.textContent = 'Googleマップで案内を開く';
  el.result.appendChild(a);
  el.result.hidden = false;

  if (apiFailed) {
    setStatus('経路サービスに接続できなかったため、直線距離から概算しました(点線)。', 'warn');
  } else if (closedOnArrival) {
    setStatus('利用可能な時間内に到着できるAEDが見つかりませんでした。最短のAEDを表示しています。時間外でも設置場所で借りられる場合があります。', 'warn');
  } else {
    setStatus('最短経路を表示しました。', 'ok');
  }
}

// ====== 初期化 ======
function init() {
  initMap();
  // 出発地の候補ボタンを生成 (GPSが使えない場合用)
  CONFIG.startPoints.forEach((p) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'chip';
    b.textContent = p.label;
    b.addEventListener('click', () => {
      setUserPosition(p.lat, p.lon, p.label);
      state.map.setView([p.lat, p.lon], 16);
    });
    el.startPoints.appendChild(b);
  });
  el.locate.addEventListener('click', locate);
  el.search.addEventListener('click', findShortest);
  el.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      applyAedData(parseWorkbook(await file.arrayBuffer()));
    } catch (err) {
      console.error(err);
      setStatus('ファイルを読み込めませんでした。AED.xlsx の形式を確認してください。', 'error');
    }
  });
  if (typeof XLSX === 'undefined') {
    setStatus('Excel読み込みライブラリ(SheetJS)を読み込めませんでした。ネットワーク接続を確認してください。', 'error');
    return;
  }
  loadAedData();
}

document.addEventListener('DOMContentLoaded', init);
