const R = window.QUARANTINE_REFERENCE;
const M = window.QuarantineModel;
const A = window.QuarantineAllocation;
const MP = window.QUARANTINE_MODEL_PARAMS;
const $ = id => document.getElementById(id);
const tabs = [...document.querySelectorAll('.workspace-tab')];
const roles = { port: '항만 ID 또는 이름', volume: '물동량', country: '국가', year: '연도', month: '월', period: '연월', inspection: '검사량', previous: '전월 검사량', source: '물동량 출처', species: '외래종 목록', latitude: '위도', longitude: '경도' };
const sources = { measured: '실측', estimated: '추정', uploaded: '업로드', mixed: '혼합', none: '값 없음', 'n/a': '기후 전용' };
const countryNames = new Intl.DisplayNames(['ko'], { type: 'region' });
const suitabilityLoads = new Map();
const weatherCache = new Map();
const DEFAULT_PARAMS = { gamma: A.DEFAULTS.gamma, budget: A.DEFAULTS.budget };
const state = { workbook: null, fileName: '', sheetRows: [], headerRow: 0, records: [], visible: [], selected: '', source: '', sourceTitle: '', issues: {}, assignedPeriod: false, weatherController: null,
  params: { ...DEFAULT_PARAMS }, ctx: null, ctxKey: '', ctxInput: null, suitability: null, result: null, metrics: null, dataVersion: 0, allocToken: 0, prevCache: new Map(), baseWarnings: [], sortMode: 'inspection' };
const format = value => Number(value).toLocaleString('ko-KR', { maximumFractionDigits: 1 });
const dateLabel = date => date ? `${date.slice(0, 4)}년 ${Number(date.slice(5))}월` : '연월 미지정';
function countryLabel(code) { try { return /^[A-Z]{2}$/.test(code) ? countryNames.of(code) : code || '국가 미확인'; } catch { return code || '국가 미확인'; } }
function localMonth() { const now = new Date(); return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`; }
function showView(id) {
  for (const tab of tabs) {
    const active = tab.dataset.view === id;
    tab.classList.toggle('is-active', active);
    tab.setAttribute('aria-selected', String(active));
    tab.tabIndex = active ? 0 : -1;
    $(tab.dataset.view).hidden = !active;
  }
  document.querySelector('.app-shell').classList.toggle('app-shell--params', id === 'params-view');
  if (id === 'params-view') refreshAllocation();
}
tabs.forEach(tab => tab.addEventListener('click', () => { if (!tab.disabled) showView(tab.dataset.view); }));
$('workspace-tabs').addEventListener('keydown', event => {
  if (!['ArrowLeft', 'ArrowRight'].includes(event.key)) return;
  const enabled = tabs.filter(tab => !tab.disabled);
  const index = enabled.indexOf(document.activeElement);
  if (index < 0) return;
  event.preventDefault();
  const next = enabled[(index + (event.key === 'ArrowRight' ? 1 : -1) + enabled.length) % enabled.length];
  showView(next.dataset.view); next.focus();
});
$('replace-file').addEventListener('click', () => showView('upload-view'));
$('open-params').addEventListener('click', () => showView('params-view'));
$('back-to-all').addEventListener('click', () => selectPort(''));

for (const [role, label] of Object.entries(roles)) {
  const field = document.createElement('label'); field.className = 'form-field'; field.htmlFor = `map-${role}`;
  field.append(document.createTextNode(label + (['port', 'volume'].includes(role) ? ' *' : '')));
  const select = document.createElement('select'); select.id = `map-${role}`; select.addEventListener('change', updateReadyState); field.append(select);
  $('column-mapping').append(field);
}
function getMapping() { return Object.fromEntries(Object.keys(roles).map(role => [role, $(`map-${role}`).value === '' ? -1 : Number($(`map-${role}`).value)])); }
function updateReadyState() { const map = getMapping(); $('show-operations').disabled = map.port < 0 || map.volume < 0; }
function setFeedback(message, error = false) { $('upload-feedback').textContent = message; $('upload-feedback').classList.toggle('is-error', error); }
function configureSheet(name) {
  const sheet = state.workbook.Sheets[name];
  state.sheetRows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  state.headerRow = M.detectHeader(state.sheetRows);
  const headers = state.sheetRows[state.headerRow] || [];
  const mapping = M.mappingFor(headers);
  for (const role of Object.keys(roles)) {
    const select = $(`map-${role}`); select.replaceChildren(new Option('연결 안 함', ''));
    headers.forEach((header, index) => select.append(new Option(`${window.XLSX.utils.encode_col(index)} · ${String(header ?? '제목 없음')}`, String(index))));
    select.value = mapping[role] >= 0 ? String(mapping[role]) : '';
  }
  $('file-options').hidden = false;
  $('file-prompt').textContent = state.fileName;
  $('mapping-details').open = mapping.port < 0 || mapping.volume < 0;
  setFeedback(`${name} 시트 · 제목 ${state.headerRow + 1}행 · ${Math.max(0, state.sheetRows.length - state.headerRow - 1)}개 자료 행`);
  updateReadyState();
}
function applyUpload() {
  const map = getMapping();
  if (map.port < 0 || map.volume < 0) { setFeedback('항만과 물동량 열을 연결해 주세요.', true); return; }
  const selected = $('reference-month').value;
  const result = M.readRows(state.sheetRows.slice(state.headerRow + 1), map, { ports: R.ports, defaultPeriod: selected || localMonth(), SSF: window.XLSX.SSF });
  if (!result.records.length) { setFeedback('읽을 수 있는 항만 자료가 없습니다. 항만·연월·물동량 열을 확인해 주세요.', true); return; }
  state.records = result.records; state.issues = result.issues; state.assignedPeriod = result.assignedPeriod;
  state.source = 'upload'; state.sourceTitle = `${state.fileName} / ${$('sheet-select').value}`;
  const dates = [...new Set(result.records.map(row => row.period))].sort();
  $('reference-month').value = dates.includes(selected) ? selected : dates.at(-1);
  state.selected = ''; invalidateAllocation(); $('operations-tab').disabled = false; renderOperations(); showView('params-view');
}
async function loadWorkbook(file) {
  if (!window.XLSX) throw new Error('엑셀 읽기 파일을 찾지 못했습니다. vendor 폴더를 확인해 주세요.');
  if (file.size > 20 * 1024 * 1024) throw new Error('20MB 이하 파일을 선택해 주세요.');
  if (!/\.(xlsx|xls|xlsb|csv)$/i.test(file.name)) throw new Error('엑셀 또는 CSV 파일을 선택해 주세요.');
  setFeedback('파일을 읽고 있습니다…');
  const workbook = window.XLSX.read(await file.arrayBuffer(), { cellDates: true });
  if (!workbook.SheetNames?.length) throw new Error('읽을 수 있는 시트가 없습니다.');
  state.workbook = workbook; state.fileName = file.name;
  $('sheet-select').replaceChildren(...workbook.SheetNames.map(name => new Option(name, name)));
  $('sheet-select').value = workbook.SheetNames.includes('port_month') ? 'port_month' : workbook.SheetNames[0];
  configureSheet($('sheet-select').value);
  if (!$('show-operations').disabled) applyUpload();
}
async function acceptFile(file) { if (!file) return; try { await loadWorkbook(file); } catch (error) { setFeedback(error.message || '파일을 읽지 못했습니다.', true); } }
$('workbook-input').addEventListener('change', async event => { await acceptFile(event.target.files?.[0]); event.target.value = ''; });
$('file-drop').addEventListener('dragover', event => { event.preventDefault(); $('file-drop').classList.add('is-dragging'); });
$('file-drop').addEventListener('dragleave', () => $('file-drop').classList.remove('is-dragging'));
$('file-drop').addEventListener('drop', async event => { event.preventDefault(); $('file-drop').classList.remove('is-dragging'); await acceptFile(event.dataTransfer.files?.[0]); });
$('sheet-select').addEventListener('change', () => configureSheet($('sheet-select').value));
$('show-operations').addEventListener('click', applyUpload);
function openReference(month) {
  state.records = M.referenceRecords(R); state.source = 'reference'; state.sourceTitle = '제공된 월별 기준자료'; state.issues = {}; state.assignedPeriod = false; state.selected = '';
  if (month) $('reference-month').value = month;
  else if (!$('reference-month').value) $('reference-month').value = [...new Set(state.records.map(row => row.period))].sort().at(-1);
  invalidateAllocation(); $('operations-tab').disabled = false; renderOperations(); showView('params-view'); syncMonthButtons();
}
function syncMonthButtons() { document.querySelectorAll('.month-btn').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.month === $('reference-month').value && state.source === 'reference'))); }
$('open-history').addEventListener('click', () => openReference(''));
document.querySelectorAll('.month-btn').forEach(button => button.addEventListener('click', () => openReference(button.dataset.month)));
$('reference-month').addEventListener('change', () => { if (state.records.length) { state.selected = ''; renderOperations(); } else refreshAllocation(); syncMonthButtons(); });

function cell(row, value, className = '') { const td = document.createElement('td'); td.textContent = value; td.className = className; row.append(td); return td; }
function renderOperations() {
  const date = $('reference-month').value;
  state.visible = state.records.filter(row => row.period === date).sort((a, b) => countryLabel(a.country).localeCompare(countryLabel(b.country), 'ko') || a.name.localeCompare(b.name, 'ko'));
  $('manifest-period').textContent = dateLabel(date);
  $('summary-ports').textContent = format(state.visible.length);
  const completeVolumes = state.visible.every(row => row.volume !== null);
  $('summary-volume').textContent = state.visible.length && completeVolumes ? format(state.visible.reduce((sum, row) => sum + row.volume, 0)) : '—';
  dropStaleContext(); renderSummaryInspections();
  const warnings = [];
  if (!state.visible.length) warnings.push(`${dateLabel(date)} 자료가 없습니다. 다른 연월을 선택하거나 해당 월 파일을 등록해 주세요.`);
  if (state.visible.some(row => row.source === 'estimated')) warnings.push('추정 물동량이 포함되어 있습니다.');
  if (state.visible.some(row => row.volume === null)) warnings.push('물동량이 없는 행을 0으로 처리하지 않았습니다.');
  if (state.issues.invalidPeriod) warnings.push(`연월을 읽지 못한 ${state.issues.invalidPeriod}개 행은 제외했습니다.`);
  if (state.issues.missingPort) warnings.push(`항만이 없는 ${state.issues.missingPort}개 행은 제외했습니다.`);
  if (state.issues.unknownPort) warnings.push(`기준목록에 없는 ${state.issues.unknownPort}개 행은 입력 이름으로 표시합니다.`);
  if (state.issues.referenceOnly) warnings.push(`기후 전용·물동량 미제공 ${state.issues.referenceOnly}개 행은 입항 목록에서 제외했습니다.`);
  if (state.assignedPeriod) warnings.push('연월 열이 없는 파일은 등록할 때 선택한 기준월에 배정했습니다.');
  state.baseWarnings = warnings; renderWarnings();
  renderRail(); renderTable(); selectPort('');
  refreshAllocation();
}
function renderRail() {
  const fragment = document.createDocumentFragment();
  for (const item of [{ id: '', name: `전체 항만 · ${state.visible.length}` }, ...sortedVisible()]) {
    const button = document.createElement('button'); button.type = 'button'; button.className = 'port-chip'; button.dataset.port = item.id;
    button.textContent = item.name; button.setAttribute('aria-pressed', String(item.id === state.selected));
    if (item.country) button.title = `${countryLabel(item.country)} · ${item.id}`;
    button.addEventListener('click', () => selectPort(item.id)); fragment.append(button);
  }
  $('port-rail').replaceChildren(fragment); $('port-list-count').textContent = `${state.visible.length}개 항만`;
}
/* 정렬: 기본은 검사량(TEU) 많은 순. 배분 결과가 없으면(계산 전) 국가·항만순으로 둡니다.
   검사량이 같으면 기후 게이트로 제외된 항만(0 TEU)보다 배분 대상 항만, 그다음 물동량 많은 순입니다.
   배분 제외·계산 불가 항만은 맨 아래에 물동량 많은 순으로 둡니다. */
function sortedVisible() {
  const rows = state.visible.slice();
  if (state.sortMode !== 'inspection' || !state.ctx) return rows;
  const rank = id => { const info = allocInfo(id); return info.status === 'ok' ? { tier: 0, I: info.I } : info.status === 'gate' ? { tier: 1, I: 0 } : { tier: 2, I: 0 }; };
  const keyed = new Map(rows.map(row => [row.id, rank(row.id)]));
  const vol = row => (row.volume === null ? -1 : row.volume);
  return rows.sort((a, b) => {
    const ka = keyed.get(a.id), kb = keyed.get(b.id);
    if (ka.tier !== kb.tier && (ka.tier === 2 || kb.tier === 2)) return ka.tier - kb.tier;
    return (kb.I - ka.I) || (ka.tier - kb.tier) || (vol(b) - vol(a)) || a.name.localeCompare(b.name, 'ko');
  });
}
function syncSortButtons() {
  document.querySelectorAll('.sort-btn').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.sort === state.sortMode)));
  const head = $('sort-inspection-head'); if (head) head.textContent = state.sortMode === 'inspection' ? '검사량 (TEU) ▼' : '검사량 (TEU)';
}
document.querySelectorAll('.sort-btn').forEach(button => button.addEventListener('click', () => {
  state.sortMode = button.dataset.sort; syncSortButtons(); renderRail(); renderTable();
}));
function renderTable() {
  const fragment = document.createDocumentFragment();
  for (const item of sortedVisible()) {
    const row = document.createElement('tr'); cell(row, countryLabel(item.country));
    const nameCell = cell(row, ''); const button = document.createElement('button'); button.type = 'button'; button.className = 'port-select'; button.textContent = item.name;
    button.addEventListener('click', () => selectPort(item.id)); nameCell.append(button);
    const code = document.createElement('small'); code.className = 'port-code'; code.textContent = item.id.startsWith('custom:') ? '기준목록 외 항만' : item.id; nameCell.append(code);
    const info = allocInfo(item.id);
    cell(row, item.volume === null ? '값 없음' : `${format(item.volume)} TEU`, 'numeric');
    cell(row, info.status === 'ok' ? pct(info.share) : '—', info.status === 'ok' ? 'numeric' : 'numeric pending-value');
    renderInspectionCell(row, item, info);
    const source = cell(row, ''); const badge = document.createElement('span'); badge.className = `source-label source-label--${item.source}`; badge.textContent = sources[item.source] || '업로드'; source.append(badge);
    fragment.append(row);
  }
  if (!state.visible.length) { const row = document.createElement('tr'); const td = cell(row, '해당 연월에 표시할 입항 자료가 없습니다.', 'empty-table'); td.colSpan = 6; fragment.append(row); }
  $('manifest-body').replaceChildren(fragment);
}
function selectPort(id) {
  state.selected = id;
  document.querySelectorAll('.port-chip').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.port === id)));
  $('manifest-card').hidden = Boolean(id); $('port-detail').hidden = !id;
  state.weatherController?.abort();
  if (!id) return;
  const item = state.visible.find(row => row.id === id); if (!item) return;
  $('detail-title').textContent = item.name;
  $('detail-country').textContent = `${countryLabel(item.country)} · ${item.id.startsWith('custom:') ? '기준목록 외 항만' : item.id} · ${dateLabel(item.period)} · ${item.volume === null ? '물동량 값 없음' : format(item.volume) + ' TEU'}`;
  $('detail-status').textContent = detailStatus(item);
  renderSpecies(item); renderChange(item); loadWeather(item); resetForecast(item); renderClimate(item);
  $('port-detail').scrollIntoView({ block: 'start', behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'instant' : 'smooth' });
}
function suitabilityFor(date) {
  if (window.QUARANTINE_S?.[date]) return Promise.resolve(window.QUARANTINE_S[date]);
  if (!R.months[date]) return Promise.resolve(null);
  if (!suitabilityLoads.has(date)) suitabilityLoads.set(date, new Promise((resolve, reject) => {
    const script = document.createElement('script'); script.src = `data/suitability/${date}.js`;
    script.onload = () => resolve(window.QUARANTINE_S?.[date] || null);
    script.onerror = () => { suitabilityLoads.delete(date); script.remove(); reject(new Error('서식 적합도 파일을 읽지 못했습니다.')); };
    document.head.append(script);
  }));
  return suitabilityLoads.get(date);
}
async function renderSpecies(item) {
  const climateDate = M.shiftMonth(item.period, -12);
  $('species-list').textContent = '서식 적합도를 불러오고 있습니다…';
  $('species-basis').textContent = `${dateLabel(climateDate)} 기후 기반`;
  try {
    const month = await suitabilityFor(climateDate);
    if (state.selected !== item.id || $('reference-month').value !== item.period) return;
    const scores = month?.[item.id];
    if (!scores) {
      $('species-list').textContent = `이 항만의 ${dateLabel(climateDate)} 서식 적합도 자료가 없습니다.`;
      if (item.species.length) { const p = document.createElement('p'); p.className = 'uploaded-species'; p.textContent = `업로드 종 목록: ${item.species.join(', ')}`; $('species-list').append(p); }
      return;
    }
    const rows = R.species.map((species, index) => ({ ...species, score: scores[index] })).filter(row => Number.isFinite(row.score)).sort((a, b) => b.score - a.score);
    const table = document.createElement('table'); table.className = 'species-table';
    const head = document.createElement('thead'); const tr = document.createElement('tr');
    for (const name of ['종 정보', '서식 적합도']) { const th = document.createElement('th'); th.scope = 'col'; th.textContent = name; tr.append(th); } head.append(tr); table.append(head);
    const body = document.createElement('tbody');
    for (const species of rows) {
      const row = document.createElement('tr'); const name = cell(row, '');
      const strong = document.createElement('strong'); strong.textContent = species.name || species.scientific;
      const scientific = document.createElement('small'); scientific.textContent = species.scientific;
      const gbif = document.createElement('small'); const records = R.gbif[item.id]?.[species.key]; gbif.textContent = `GBIF 100km 기록 ${records ? format(records) + '건' : '—'}`;
      name.append(strong, scientific, gbif);
      const score = cell(row, '', 'numeric'); const text = document.createElement('strong'); text.textContent = species.score < .001 ? species.score.toExponential(2) : species.score.toFixed(3);
      const bar = document.createElement('span'); bar.className = 'score-track'; const fill = document.createElement('span'); fill.style.width = `${species.score * 100}%`; bar.append(fill); score.append(text, bar);
      body.append(row);
    }
    table.append(body); $('species-list').replaceChildren(table);
    $('species-basis').textContent = `${dateLabel(climateDate)} 기후 · ${rows.length}종`;
  } catch (error) { if (state.selected === item.id && $('reference-month').value === item.period) $('species-list').textContent = error.message; }
}
async function renderChange(item) {
  const previousDate = M.shiftMonth(item.period, -1);
  const out = $('inspection-change');
  const stale = () => state.selected !== item.id || $('reference-month').value !== item.period;
  const show = (value, percentBase, caption) => {
    const difference = value.current - value.previous;
    const percent = percentBase > 0 ? ` (${difference > 0 ? '+' : ''}${(difference / percentBase * 100).toFixed(1)}%)` : '';
    out.className = 'change-value'; out.textContent = `${difference > 0 ? '+' : ''}${format(difference)}${percent}`;
    out.classList.toggle('is-positive', difference > 0); out.classList.toggle('is-negative', difference < 0);
    $('change-caption').textContent = caption;
  };
  const none = () => { out.className = 'change-value'; out.textContent = '비교 자료 없음'; $('change-caption').textContent = `${dateLabel(previousDate)}와 이번 달 검사량이 모두 있어야 비교할 수 있습니다.`; };
  const info = allocInfo(item.id);
  const current = info.status === 'ok' ? info.I : info.status === 'gate' ? 0 : null;
  if (current !== null) {
    const prev = await previousAllocation(item.period);
    if (stale()) return;
    const index = prev?.index.get(item.id);
    if (index !== undefined) {
      const previous = prev.ctx.gate[index] ? prev.result.Iint[index] : 0;
      show({ current, previous }, previous, `이번 달 ${format(current)} · 전월 ${format(previous)} TEU (현재 γ·B로 계산한 값)`); return;
    }
  }
  const previous = item.previousInspection ?? state.records.find(row => row.period === previousDate && row.id === item.id)?.inspection ?? null;
  if (item.inspection === null || previous === null) { none(); return; }
  show({ current: item.inspection, previous }, previous, `이번 달 ${format(item.inspection)} · 전월 ${format(previous)} (업로드 원본값)`);
}
function weatherDescription(code) {
  if (code === 0) return '맑음'; if (code === 1) return '대체로 맑음'; if (code === 2) return '구름 조금'; if (code === 3) return '흐림';
  if ([45, 48].includes(code)) return '안개'; if ([51, 53, 55, 56, 57].includes(code)) return '이슬비';
  if ([61, 63, 65, 66, 67, 80, 81, 82].includes(code)) return '비'; if ([71, 73, 75, 77, 85, 86].includes(code)) return '눈';
  return [95, 96, 99].includes(code) ? '뇌우' : '현재 기상';
}
const WEATHER_URL = 'https://api.open-meteo.com/v1/forecast';
const WEATHER_FULL = 'temperature_2m,relative_humidity_2m,precipitation,wind_speed_10m,weather_code';
function weatherFailure(error) {
  if (error?.status === 429) return '요청이 많아 잠시 제한되었습니다(HTTP 429). 1~2분 뒤 다시 선택해 보세요.';
  if (error?.status) return `기상 서비스 응답 오류(HTTP ${error.status})입니다.`;
  if (error?.timedOut) return '응답 시간이 초과되었습니다. 네트워크가 느리거나 기상 서버로 연결이 막혀 있을 수 있습니다.';
  if (navigator.onLine === false) return '인터넷에 연결되어 있지 않습니다.';
  return '기상 서버(api.open-meteo.com)에 연결하지 못했습니다. 방화벽·학교/회사 망 차단, 광고·개인정보 차단 확장 프로그램을 확인해 주세요.';
}
async function fetchWeatherJson(query, controller) {
  let timer = 0;
  try {
    const timeout = new Promise((_, reject) => { timer = setTimeout(() => { const e = new Error('timeout'); e.timedOut = true; reject(e); controller.abort(); }, 12000); });
    const request = (async () => {
      const response = await fetch(`${WEATHER_URL}?${new URLSearchParams(query)}`, { signal: controller.signal, cache: 'no-store' });
      if (!response.ok) { const e = new Error('http'); e.status = response.status; throw e; }
      return response.json();
    })();
    return await Promise.race([request, timeout]);
  } finally { clearTimeout(timer); }
}
function normalizeWeather(data) {
  if (data?.current && Number.isFinite(data.current.temperature_2m)) return data.current;
  const legacy = data?.current_weather;
  if (legacy && Number.isFinite(legacy.temperature)) return { time: legacy.time, temperature_2m: legacy.temperature, wind_speed_10m: legacy.windspeed, weather_code: legacy.weathercode, relative_humidity_2m: null, precipitation: null };
  throw new Error('기상 값 없음');
}
async function loadWeather(item) {
  $('weather-time').textContent = '현재 시점'; $('weather-main').textContent = '불러오는 중'; $('weather-caption').textContent = '항만 좌표의 현재 기상을 조회합니다.';
  if (item.lat === null || item.lon === null) { $('weather-main').textContent = '좌표 없음'; $('weather-caption').textContent = '위도·경도 열을 연결하면 조회할 수 있습니다.'; return; }
  const controller = new AbortController(); state.weatherController = controller;
  const key = `${item.lat},${item.lon}`, base = { latitude: String(item.lat), longitude: String(item.lon), timezone: 'auto' };
  const stale = () => state.selected !== item.id || state.weatherController !== controller;
  try {
    let current = weatherCache.get(key);
    if (!current) {
      // 1) 전체 항목 → 2) 같은 요청 재시도 → 3) 구형 current_weather 형식 (서버 일시 오류·형식 차이 대비)
      const attempts = [{ ...base, current: WEATHER_FULL }, { ...base, current: WEATHER_FULL }, { ...base, current_weather: 'true' }];
      let lastError = null;
      for (let i = 0; i < attempts.length && !current; i += 1) {
        try { current = normalizeWeather(await fetchWeatherJson(attempts[i], controller)); }
        catch (error) {
          if (stale()) return; lastError = error;
          if (i < attempts.length - 1) await new Promise(resolve => setTimeout(resolve, error?.status === 429 ? 2500 : 700));
          if (stale()) return;
          if (error?.timedOut) { return fail(weatherFailure(error)); }
        }
      }
      if (!current) throw lastError;
      weatherCache.set(key, current);
    }
    if (stale()) return;
    $('weather-main').textContent = `${Math.round(current.temperature_2m)}°C · ${weatherDescription(current.weather_code)}`;
    $('weather-time').textContent = String(current.time || '').replace('T', ' ') || '현재 시점';
    const parts = [];
    if (current.relative_humidity_2m !== null && current.relative_humidity_2m !== undefined) parts.push(`습도 ${current.relative_humidity_2m}%`);
    if (current.wind_speed_10m !== undefined) parts.push(`풍속 ${current.wind_speed_10m} km/h`);
    if (current.precipitation !== null && current.precipitation !== undefined) parts.push(`강수 ${current.precipitation} mm`);
    $('weather-caption').textContent = parts.join(' · ');
  } catch (error) {
    if (!stale()) fail(weatherFailure(error));
  }
  function fail(message) {
    if (stale()) return;
    $('weather-main').textContent = '현재 기상 조회 불가'; $('weather-caption').textContent = message;
  }
}
function renderClimate(item) {
  const date = M.shiftMonth(item.period, -12), row = R.months[date]?.find(row => row[0] === item.id);
  const mean = row?.[3 + R.climateFields.indexOf('x_tmean_c')];
  $('climate-caption').textContent = mean === undefined || mean === null ? '모형 기준 기후 자료 없음' : `모형 기준 기후: ${dateLabel(date)} 평균 ${format(mean)}°C`;
}
function resetForecast(item) {
  const ceiling = Math.max(100000, Math.ceil((item.volume || 0) * 2 / 10000) * 10000);
  $('forecast-period').textContent = dateLabel(M.shiftMonth(item.period, 1));
  $('forecast-range').max = String(ceiling); $('forecast-range').value = String(item.volume || 0); $('forecast-range').disabled = false;
  $('forecast-value').value = String(item.volume || 0); $('forecast-range-max').textContent = `${format(ceiling)} TEU`;
  $('forecast-value').removeAttribute('aria-invalid'); $('forecast-feedback').textContent = '';
  updateForecast();
}
$('forecast-range').addEventListener('input', event => { $('forecast-value').value = event.target.value; $('forecast-value').removeAttribute('aria-invalid'); $('forecast-feedback').textContent = ''; updateForecast(); });
$('forecast-value').addEventListener('input', event => {
  const value = M.number(event.target.value);
  $('forecast-value').setAttribute('aria-invalid', String(value === null)); $('forecast-feedback').textContent = value === null ? '0 이상의 물동량을 입력해 주세요.' : '';
  if (value === null) { updateForecast(); return; }
  if (value > Number($('forecast-range').max)) { const max = Math.ceil(value / 10000) * 10000; $('forecast-range').max = String(max); $('forecast-range-max').textContent = `${format(max)} TEU`; }
  $('forecast-range').value = String(value);
  updateForecast();
});
$('history-caption').textContent = `185개 항만 · 분석 74종 · 물동량 2016.01–2026.09`;

/* ===================== 배분 계산 · 파라미터 탭 (설계서 5~9장) ===================== */
const MIN_BUDGET = A.DEFAULTS.minBudget, GAMMA_MAX = 2;
const portById = new Map(R.ports.map(port => [port.id, port]));
let referenceCache = null, refreshQueued = false;
const pct = (value, digits = 1) => `${(value * 100).toFixed(digits)}%`;
const fixed = (value, digits = 2) => Number(value).toFixed(digits);
const portName = id => portById.get(id)?.name || id;
function namesOf(ids) { const names = ids.map(portName); return names.length > 3 ? `${names.slice(0, 3).join(', ')} 외 ${names.length - 3}개 항만` : names.join(', '); }

function currentInput() {
  const date = $('reference-month').value;
  if (state.records.length) return { loaded: true, period: date, records: state.records.filter(row => row.period === date), label: state.sourceTitle };
  referenceCache = referenceCache || M.referenceRecords(R);
  const period = date || Object.keys(R.months).sort().at(-1);
  return { loaded: false, period, records: referenceCache.filter(row => row.period === period), label: '제공된 월별 기준자료' };
}
const keyFor = input => `${state.dataVersion}|${input.loaded ? 'upload' : 'reference'}|${input.period}`;
function invalidateAllocation() { state.dataVersion++; state.ctx = null; state.result = null; state.metrics = null; state.prevCache.clear(); }
function dropStaleContext() { if (state.ctx && state.ctxKey !== keyFor(currentInput())) { state.ctx = null; state.result = null; state.metrics = null; } }

async function refreshAllocation() {
  const token = ++state.allocToken;
  const input = currentInput();
  if (!input.records.length) { state.ctx = null; state.result = null; state.metrics = null; afterAllocation(); return; }
  if (!state.ctx || state.ctxKey !== keyFor(input)) {
    let suitability = null;
    try { suitability = await suitabilityFor(M.shiftMonth(input.period, -12)); } catch { suitability = null; }
    if (token !== state.allocToken) return;
    state.suitability = suitability;
    state.ctx = A.buildContext({ records: input.records, period: input.period, ref: R, model: MP, suitability });
    state.ctxKey = keyFor(input); state.ctxInput = input;
  }
  state.result = A.allocate(state.ctx, state.params.gamma, state.params.budget);
  state.metrics = A.metrics(state.ctx, state.result);
  afterAllocation();
}
function queueRefresh() { if (refreshQueued) return; refreshQueued = true; requestAnimationFrame(() => { refreshQueued = false; refreshAllocation(); }); }
function afterAllocation() {
  syncParamControls(); renderSummaryInspections(); renderWarnings();
  if (!$('operations-view').hidden || state.visible.length) { renderRail(); renderTable(); refreshSelectedDetail(); }
  renderParams();
}

function allocInfo(id) {
  const ctx = state.ctx;
  if (!ctx) return { status: 'pending' };
  const index = ctx.ids.indexOf(id);
  const reason = ctx.excluded.find(entry => entry.id === id)?.reason;
  if (reason && reason !== 'gate') return { status: 'skipped', reason };
  if (index < 0) return { status: 'skipped', reason: 'no-climate' };
  if (!ctx.gate[index]) return { status: 'gate', reason: 'gate', I: 0 };
  if (!state.result) return { status: 'skipped', reason: 'no-climate' };
  return { status: 'ok', I: state.result.Iint[index], c: state.result.c[index], share: ctx.C[index] };
}
function renderInspectionCell(row, item, info) {
  const td = cell(row, '', 'numeric');
  const main = document.createElement('span');
  if (info.status === 'pending') { main.textContent = '계산 중'; td.classList.add('pending-value'); }
  else main.textContent = info.status === 'ok' ? format(info.I) : info.status === 'gate' ? '0' : '—';
  td.append(main);
  if (info.status === 'gate' || info.status === 'skipped') {
    const note = document.createElement('small'); note.className = 'cell-note'; note.textContent = A.REASONS[info.reason]; note.title = A.REASON_TEXT[info.reason]; td.append(note);
  }
  if (item.inspection !== null) { const note = document.createElement('small'); note.className = 'cell-note'; note.textContent = `업로드 값 ${format(item.inspection)}`; td.append(note); }
}
function detailStatus(item) {
  const info = allocInfo(item.id);
  if (info.status === 'ok') return `검사량 ${format(info.I)} TEU`;
  if (info.status === 'gate') return '검사 대상 제외 (기후 게이트)';
  if (info.status === 'skipped') return `배분 제외 · ${A.REASONS[info.reason]}`;
  return '검사량 계산 중';
}
function refreshSelectedDetail() {
  const item = state.visible.find(row => row.id === state.selected);
  if (!item || item.period !== $('reference-month').value) return;
  $('detail-status').textContent = detailStatus(item); renderChange(item); updateForecast();
}
function renderSummaryInspections() {
  const value = $('summary-inspections');
  if (!state.visible.length || !state.result) { value.textContent = state.visible.length && !state.ctx ? '…' : '—'; return; }
  value.textContent = format(state.result.Iint.reduce((total, count) => total + count, 0));
}
function renderWarnings() {
  const notes = [...state.baseWarnings], ctx = state.ctx;
  if (ctx && state.visible.length) {
    const by = reason => ctx.excluded.filter(entry => entry.reason === reason).map(entry => entry.id);
    if (by('gate').length) notes.push(`기후 게이트로 검사 대상에서 제외한 항만: ${namesOf(by('gate'))}.`);
    if (by('no-climate').length) notes.push(`전년 동월 기후 자료가 없는 ${by('no-climate').length}개 항만은 배분에서 제외했습니다.`);
    if (by('below-min').length) notes.push(`물동량 1 TEU 미만 ${by('below-min').length}개 항만은 배분에서 제외했습니다.`);
    if (state.result?.capped) notes.push(`월 검사량이 배분 대상 총 물동량보다 커서 ${state.result.budget} TEU로 계산했습니다.`);
    if (state.params.budget < MIN_BUDGET) notes.push(lowBudgetText());
  }
  $('data-warning').hidden = !notes.length; $('data-warning').textContent = notes.join(' ');
}

async function previousAllocation(period) {
  const previousPeriod = M.shiftMonth(period, -1);
  const records = state.records.filter(row => row.period === previousPeriod);
  if (!records.length) return null;
  const key = `${state.dataVersion}|${previousPeriod}`;
  let ctx = state.prevCache.get(key);
  if (!ctx) {
    let suitability = null;
    try { suitability = await suitabilityFor(M.shiftMonth(previousPeriod, -12)); } catch { suitability = null; }
    ctx = A.buildContext({ records, period: previousPeriod, ref: R, model: MP, suitability });
    state.prevCache.set(key, ctx);
  }
  const result = A.allocate(ctx, state.params.gamma, state.params.budget);
  return result ? { ctx, result, index: new Map(ctx.ids.map((id, i) => [id, i])) } : null;
}
function updateForecast() {
  const out = $('forecast-inspection');
  const item = state.visible.find(row => row.id === state.selected);
  if (!item) return;
  const value = M.number($('forecast-value').value);
  if (value === null) { out.textContent = '—'; return; }
  if (!state.ctx || !state.ctxInput || state.ctx.period !== item.period) { out.textContent = '계산 중…'; return; }
  const records = state.ctxInput.records.map(row => row.id === item.id ? { ...row, volume: value } : row);
  const ctx = A.buildContext({ records, period: state.ctx.period, ref: R, model: MP, suitability: state.suitability });
  const reason = ctx.excluded.find(entry => entry.id === item.id)?.reason;
  if (reason === 'gate') { out.textContent = '0 TEU · 기후 게이트 제외'; return; }
  if (reason) { out.textContent = `배분 제외 · ${A.REASONS[reason]}`; return; }
  const result = A.allocate(ctx, state.params.gamma, state.params.budget);
  out.textContent = result ? `${format(result.Iint[ctx.ids.indexOf(item.id)])} TEU` : '계산 불가';
}

/* ---- 파라미터 입력 ---- */
function syncParamControls(force = false) {
  const { gamma, budget } = state.params;
  if (force || document.activeElement !== $('gamma-value')) $('gamma-value').value = String(Number(gamma.toFixed(2)));
  if (force || document.activeElement !== $('budget-value')) $('budget-value').value = String(budget);
  const base = budgetBase(), pctInput = $('budget-pct');
  pctInput.disabled = !base;
  if (force || document.activeElement !== pctInput) pctInput.value = base ? String(Number((Math.min(1, budget / base) * 100).toFixed(2))) : '';
  if (state.params.budget < MIN_BUDGET && !$('budget-value').hasAttribute('aria-invalid')) $('budget-feedback').textContent = '설계 최소(' + MIN_BUDGET + ' TEU) 미만입니다. 계산은 되지만 포착이 현행보다 낮을 수 있습니다.';
  else if (!$('budget-value').hasAttribute('aria-invalid') && !$('budget-pct').hasAttribute('aria-invalid')) $('budget-feedback').textContent = '';
  $('budget-range-max').textContent = String(budgetAxisMax());
  for (const slider of Object.values(sliders)) slider.update();
}
function setValid(input, feedback, message) { input.toggleAttribute('aria-invalid', Boolean(message)); if (message) input.setAttribute('aria-invalid', 'true'); feedback.textContent = message; }
function applyGamma(value) { state.params.gamma = Math.round(value * 1000) / 1000; setValid($('gamma-value'), $('gamma-feedback'), ''); syncParamControls(); queueRefresh(); }
function applyBudget(value) { state.params.budget = value; setValid($('budget-value'), $('budget-feedback'), ''); syncParamControls(); queueRefresh(); }
$('gamma-value').addEventListener('input', event => {
  const value = event.target.valueAsNumber;
  if (!Number.isFinite(value) || value < 0 || value > GAMMA_MAX) { setValid($('gamma-value'), $('gamma-feedback'), `0 이상 ${GAMMA_MAX} 이하의 값을 입력해 주세요.`); return; }
  applyGamma(value);
});
$('gamma-value').addEventListener('change', () => { setValid($('gamma-value'), $('gamma-feedback'), ''); syncParamControls(true); });
const lowBudgetText = () => `월 검사량 ${state.params.budget} TEU는 설계 최소(${MIN_BUDGET} TEU)보다 적습니다. 계산은 하지만 검사량이 적으면 위험 항만에 나눠 담을 여지가 작아 현행 무작위보다 포착이 낮을 수 있습니다(설계서 10.4절).`;
const budgetBase = () => state.result?.totalN ?? state.ctx?.totalN ?? 0; // 배분 대상 총물동량(TEU)
$('budget-value').addEventListener('input', event => {
  const value = event.target.valueAsNumber;
  if (!Number.isFinite(value) || value < 1) { setValid($('budget-value'), $('budget-feedback'), '1 TEU 이상의 값을 입력해 주세요.'); return; }
  applyBudget(Math.round(value));
});
$('budget-value').addEventListener('change', () => { setValid($('budget-value'), $('budget-feedback'), ''); syncParamControls(true); });
$('budget-pct').addEventListener('input', event => {
  const value = event.target.valueAsNumber, base = budgetBase();
  if (!base) { setValid($('budget-pct'), $('budget-feedback'), '총물동량을 알 수 없어 %로 입력할 수 없습니다. 자료를 먼저 불러와 주세요.'); return; }
  if (!Number.isFinite(value) || value <= 0 || value > 100) { setValid($('budget-pct'), $('budget-feedback'), '0 초과 100 이하의 %를 입력해 주세요.'); return; }
  setValid($('budget-pct'), $('budget-feedback'), '');
  applyBudget(Math.max(1, Math.round(base * value / 100)));
});
$('budget-pct').addEventListener('change', () => { setValid($('budget-pct'), $('budget-feedback'), ''); syncParamControls(true); });
$('reset-params').addEventListener('click', () => { state.params = { ...DEFAULT_PARAMS }; setValid($('gamma-value'), $('gamma-feedback'), ''); setValid($('budget-value'), $('budget-feedback'), ''); syncParamControls(true); queueRefresh(); });

/* ---- 파라미터 탭 화면 ---- */
const CHART_COLORS = { proposed: '#82384e', baseline: '#7a8799', tv: '#315fae' };
let chartFitKey = '';
function fitParamCharts() {
  const view = $('params-view'); if (view.hidden) return false;
  if (window.innerWidth <= 900) { view.style.removeProperty('--chart-h'); chartFitKey = ''; return false; }
  const key = [window.innerWidth, window.innerHeight, $('params-warning').hidden].join('|');
  if (key === chartFitKey) return false;
  const doc = document.documentElement, fits = h => { view.style.setProperty('--chart-h', h + 'px'); return doc.scrollHeight <= window.innerHeight; };
  let lo = 150, hi = Math.max(150, window.innerHeight);
  if (fits(hi)) lo = hi; else while (hi - lo > 1) { const mid = (lo + hi) >> 1; if (fits(mid)) lo = mid; else hi = mid; }
  view.style.setProperty('--chart-h', lo + 'px'); chartFitKey = key; return true;
}
function renderParams() { renderParamsOnce(); if (fitParamCharts()) renderParamsOnce(); }
function renderParamsOnce() {
  if ($('params-view').hidden) return;
  syncParamControls();
  const input = currentInput(), ctx = state.ctx, warning = $('params-warning');
  const blank = message => {
    warning.hidden = !message; warning.textContent = message || '';
    for (const id of ['kpi-capture', 'kpi-ratio', 'kpi-tv', 'kpi-ports']) { $(id).textContent = '—'; $(`${id}-note`).textContent = ''; }
    for (const id of ['chart-gamma-tv', 'chart-budget-capture']) $(id).replaceChildren();
    for (const id of ['readout-gamma-tv', 'readout-budget-capture']) $(id).textContent = '';
  };
  if (!input.records.length) return blank(`${dateLabel(input.period)} 자료가 없어 변화를 계산할 수 없습니다. 조회 기준 연월을 바꾸거나 자료를 등록해 주세요.`);
  if (!ctx) return blank('');
  if (!state.result) return blank(ctx.ids.length ? '기후 게이트를 통과한 항만이 없어 검사량을 계산할 수 없습니다.' : `${dateLabel(input.period)}에는 전년 동월 기후 자료가 있는 항만이 없어 검사량을 계산할 수 없습니다.`);
  const notes = [];
  if (!ctx.hasTruth) notes.push('서식 적합도 파일을 읽지 못해 위험 포착 비율과 배분 일치도는 표시하지 않습니다.');
  if (state.result.capped) notes.push(`월 검사량이 배분 대상 총 물동량(${format(state.result.totalN)} TEU)보다 커서 ${state.result.budget} TEU로 계산했습니다.`);
  if (state.params.budget < MIN_BUDGET) notes.push(lowBudgetText());
  warning.hidden = !notes.length; warning.textContent = notes.join(' ');
  renderKpis(); renderGammaCharts(); renderBudgetCharts();
}
function renderKpis() {
  const m = state.metrics, ctx = state.ctx;
  const gateOut = ctx.excluded.filter(entry => entry.reason === 'gate').length;
  $('kpi-capture').textContent = m.capture === null ? '—' : pct(m.capture, m.capture < 0.01 ? 2 : 1);
  $('kpi-capture-note').textContent = m.capture === null ? '' : `현행 무작위 ${pct(m.baseline, m.baseline < 0.01 ? 2 : 1)} · 월 ${state.result.budget}개`;
  const gain = m.ratio === null ? null : (m.ratio - 1) * 100;
  $('kpi-ratio').textContent = gain === null ? '—' : `${gain >= 0 ? '+' : '−'}${Math.abs(gain).toFixed(Math.abs(gain) < 10 ? 1 : 0)}% (${fixed(m.ratio)}배)`;
  $('kpi-ratio-note').textContent = m.ratio === null ? '' : m.ratio < 1 ? '현행보다 낮음 · 검사량이 적으면 나눠 담을 여지가 작습니다' : `최선 배분 대비 ${fixed(m.foh, 3)} (현행 ${fixed(m.baselineFoh, 3)})`;
  $('kpi-tv').textContent = m.tv === null ? '—' : fixed(1 - m.tv);
  $('kpi-tv-note').textContent = m.tv === null ? '' : '검사 배분이 위험 분포와 겹치는 비율';
  $('kpi-ports').textContent = `${m.inspected} / ${m.ports}`;
  $('kpi-ports-note').textContent = `게이트 통과 ${ctx.gateN}개 중${gateOut ? ` · 제외 ${gateOut}개` : ''}`;
}
const niceCeil = value => { if (!(value > 0)) return 1; const unit = Math.pow(10, Math.floor(Math.log10(value))), f = value / unit; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * unit; };
function lineChart(host, readout, cfg) {
  const W = Math.max(280, host.clientWidth || 480), H = Math.max(150, host.clientHeight || 240), m = { l: 56, r: 16, t: 10, b: 42 }, iw = W - m.l - m.r, ih = H - m.t - m.b;
  const values = cfg.series.flatMap(item => item.points.map(point => point[1])).filter(Number.isFinite);
  const yMax = niceCeil(Math.max(...values, 1e-9) * 1.1 / 4) * 4;
  const sx = x => m.l + (x - cfg.xMin) / (cfg.xMax - cfg.xMin) * iw, sy = y => m.t + ih - y / yMax * ih;
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${cfg.label}" preserveAspectRatio="xMidYMid meet">`;
  if (cfg.shade) svg += `<rect class="chart-shade" x="${sx(cfg.shade.from)}" y="${m.t}" width="${sx(cfg.shade.to) - sx(cfg.shade.from)}" height="${ih}"></rect><text class="chart-shade-label" x="${sx(cfg.shade.from) + 4}" y="${m.t + 11}">${cfg.shade.label}</text>`;
  for (let i = 0; i <= 4; i++) { const y = yMax * i / 4; svg += `<line class="chart-gridline" x1="${m.l}" x2="${W - m.r}" y1="${sy(y)}" y2="${sy(y)}"></line><text class="chart-tick" x="${m.l - 8}" y="${sy(y) + 4.5}" text-anchor="end">${cfg.yFormat(y, yMax)}</text>`; }
  for (const x of cfg.xTicks) svg += `<line class="chart-axis" x1="${sx(x)}" x2="${sx(x)}" y1="${m.t + ih}" y2="${m.t + ih + 4}"></line><text class="chart-tick" x="${sx(x)}" y="${m.t + ih + 18}" text-anchor="middle">${cfg.xFormat(x)}</text>`;
  svg += `<text class="chart-title" x="${m.l + iw / 2}" y="${H - 6}" text-anchor="middle">${cfg.xTitle}</text><line class="chart-axis" x1="${m.l}" x2="${W - m.r}" y1="${m.t + ih}" y2="${m.t + ih}"></line>`;
  cfg.series.forEach((item, index) => {
    svg += `<polyline fill="none" stroke="${item.color}" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round" ${item.dash ? 'stroke-dasharray="6 4"' : ''} points="${item.points.map(point => `${sx(point[0]).toFixed(1)},${sy(point[1]).toFixed(1)}`).join(' ')}"></polyline>`;
    if (cfg.series.length > 1) svg += `<line x1="${m.l + 8 + index * 120}" x2="${m.l + 26 + index * 120}" y1="${m.t - 12}" y2="${m.t - 12}" stroke="${item.color}" stroke-width="2.2" ${item.dash ? 'stroke-dasharray="5 3"' : ''}></line><text class="chart-legend" x="${m.l + 31 + index * 120}" y="${m.t - 8}">${item.name}</text>`;
  });
  svg += `<line class="chart-marker" x1="${sx(cfg.marker.x)}" x2="${sx(cfg.marker.x)}" y1="${m.t}" y2="${m.t + ih}"></line>`;
  cfg.series.forEach((item, index) => { if (Number.isFinite(cfg.marker.values[index])) svg += `<circle cx="${sx(cfg.marker.x)}" cy="${sy(cfg.marker.values[index])}" r="4.5" fill="#fff" stroke="${item.color}" stroke-width="2.4"></circle>`; });
  svg += `<line class="chart-guide" x1="0" x2="0" y1="${m.t}" y2="${m.t + ih}" hidden></line></svg>`;
  host.innerHTML = svg; readout.textContent = cfg.defaultText;
  const el = host.querySelector('svg'), guide = el.querySelector('.chart-guide'), points = cfg.series[0].points;
  const move = event => {
    const box = el.getBoundingClientRect(); const x = cfg.xMin + ((event.clientX - box.left) / box.width * W - m.l) / iw * (cfg.xMax - cfg.xMin);
    let best = 0; points.forEach((point, index) => { if (Math.abs(point[0] - x) < Math.abs(points[best][0] - x)) best = index; });
    guide.removeAttribute('hidden'); guide.setAttribute('x1', sx(points[best][0])); guide.setAttribute('x2', sx(points[best][0])); readout.textContent = cfg.hoverText(best);
  };
  el.addEventListener('pointermove', move); el.addEventListener('pointerdown', move);
  el.addEventListener('pointerleave', () => { guide.setAttribute('hidden', ''); readout.textContent = cfg.defaultText; });
}
function renderGammaCharts() {
  const ctx = state.ctx, m = state.metrics, budget = state.result.budget, gamma = state.params.gamma;
  if (m.tv === null) { $('chart-gamma-tv').replaceChildren(); $('readout-gamma-tv').textContent = ''; return; }
  const sweep = A.sweepGamma(ctx, budget, Array.from({ length: 41 }, (_, i) => i * 0.05));
  lineChart($('chart-gamma-tv'), $('readout-gamma-tv'), { label: 'γ에 따른 배분 일치도 변화', xMin: 0, xMax: GAMMA_MAX, xTicks: [0, 0.5, 1, 1.5, 2], xFormat: x => String(x), xTitle: '위험 집중도 γ',
    series: [{ name: '배분 일치도', color: CHART_COLORS.tv, points: sweep.map(row => [row.gamma, 1 - row.tv]) }], yFormat: value => value.toFixed(2), marker: { x: gamma, values: [1 - m.tv] },
    defaultText: `γ ${fixed(gamma)} → ${fixed(1 - m.tv)}`, hoverText: i => `γ ${fixed(sweep[i].gamma)} → ${fixed(1 - sweep[i].tv)}` });
}
function budgetAxisMax() { return Math.max(8, Math.min(state.result ? Math.floor(state.result.totalN) : 100, Math.max(100, Math.ceil(state.params.budget * 1.25 / 10) * 10))); }
function renderBudgetCharts() {
  const ctx = state.ctx, m = state.metrics, budget = state.result.budget, gamma = state.params.gamma;
  if (m.capture === null) { $('chart-budget-capture').replaceChildren(); $('readout-budget-capture').textContent = ''; return; }
  const xMax = Math.max(budgetAxisMax(), budget), xMin = Math.max(1, Math.min(5, budget));
  const sweep = A.sweepBudget(ctx, gamma, Array.from({ length: xMax - xMin + 1 }, (_, i) => xMin + i)).filter(row => row.capture !== undefined);
  const percent = value => `${(value * 100).toFixed(value < 0.01 ? 2 : 1)}%`;
  lineChart($('chart-budget-capture'), $('readout-budget-capture'), { label: '월 검사량에 따른 위험 포착 비율 변화', xMin, xMax, xTicks: xMax > 150 ? [...new Set([xMin, ...[1, 2, 3, 4].map(i => Math.round(xMax * i / 4)), xMax])] : [...new Set([xMin, 5, ...[15, 30, 45, 60, 75, 90, 105, 120, 135, 150].filter(x => x <= xMax)])].filter(x => x >= xMin), xFormat: x => String(x), xTitle: '월 검사량 B (TEU)',
    shade: { from: xMin, to: Math.min(MIN_BUDGET, xMax), label: '설계 최소 미만(계산됨)' },
    series: [{ name: '위험 포착 비율', color: CHART_COLORS.proposed, points: sweep.map(row => [row.budget, row.capture * 100]) }],
    yFormat: (value, max) => `${value.toFixed(max < 1 ? 2 : 1)}%`, marker: { x: budget, values: [m.capture * 100] },
    defaultText: `B ${budget} → ${percent(m.capture)}`, hoverText: i => `B ${sweep[i].budget} → ${percent(sweep[i].capture)}` });
}

/* ---- 원통형 바 + 원형 버튼 슬라이더 ---- */
const sliders = {};
function setupSlider(id, cfg) {
  const el = $(id), knob = el.querySelector('.pill-knob'), fill = el.querySelector('.pill-fill');
  const update = () => {
    const { min, max } = cfg.range(), value = Math.min(max, Math.max(min, cfg.get()));
    knob.style.left = fill.style.width = `${(value - min) / (max - min) * 100}%`;
    el.setAttribute('aria-valuemin', String(min)); el.setAttribute('aria-valuemax', String(max)); el.setAttribute('aria-valuenow', String(value));
  };
  const pick = event => {
    const box = el.getBoundingClientRect(), { min, max, lo = min } = cfg.range();
    let value = min + (event.clientX - box.left) / box.width * (max - min);
    value = Math.min(max, Math.max(lo, Math.round(value / cfg.step) * cfg.step));
    cfg.set(Number(value.toFixed(3)));
  };
  el.addEventListener('pointerdown', event => { el.setPointerCapture(event.pointerId); el.classList.add('is-dragging'); el.focus(); pick(event); });
  el.addEventListener('pointermove', event => { if (el.hasPointerCapture(event.pointerId)) pick(event); });
  for (const type of ['pointerup', 'pointercancel']) el.addEventListener(type, event => { el.classList.remove('is-dragging'); if (el.hasPointerCapture(event.pointerId)) el.releasePointerCapture(event.pointerId); });
  el.addEventListener('keydown', event => {
    const { min, max, lo = min } = cfg.range(); const keys = { ArrowLeft: -1, ArrowDown: -1, ArrowRight: 1, ArrowUp: 1, PageDown: -10, PageUp: 10 };
    let value = cfg.get();
    if (event.key in keys) value += keys[event.key] * cfg.step; else if (event.key === 'Home') value = lo; else if (event.key === 'End') value = max; else return;
    event.preventDefault(); cfg.set(Number(Math.min(max, Math.max(lo, value)).toFixed(3)));
  });
  sliders[id] = { update };
}
setupSlider('gamma-slider', { step: 0.05, range: () => ({ min: 0, max: GAMMA_MAX }), get: () => state.params.gamma, set: applyGamma });
setupSlider('budget-slider', { step: 1, range: () => ({ min: 5, max: budgetAxisMax() }), get: () => state.params.budget, set: applyBudget });
window.addEventListener('resize', () => { if (!$('params-view').hidden) renderParams(); });
syncParamControls(true);
