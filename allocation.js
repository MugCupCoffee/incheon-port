/*
 * 검역 자원 배분 계산 (설계서 v5: 5·6·8·9장)
 *
 *   게이트   W_g = Σ_s H_s · W_gs ,  m_g = 1[W_g ≥ τ]
 *   위험비중 P̂_g = C_g (74종 연보 가중이동평균 × 물동량^α, 게이트 통과 항만만, 5% 균등 바닥)
 *   배분     c_g = min(1, θ·(P̂_g / N_g)^γ),  Σ N_g c_g = B,  최대잔여법 정수화
 *   평가     가상 정답 p*(서식 적합도·기후 유사도·물동량)에 대한 유입 차단율(포착)·배분 TV·FoH
 *
 * 적발 표본이 없는 운영 상태(n = 0)이므로 P̂_g = C_g 입니다.
 * 브라우저(window.QuarantineAllocation)와 node(require) 양쪽에서 쓸 수 있습니다.
 */
(function (root) {
  const DEFAULTS = { alpha: 0.3, tau: 0.35, floor: 0.05, gamma: 0.5, budget: 25, minBudget: 15, minVolume: 1 };
  const REASONS = {
    domestic: '국내 항만', 'no-volume': '물동량 없음', 'below-min': '1 TEU 미만', 'no-climate': '기후 자료 없음',
    'no-suit': '적합도 자료 없음', gate: '기후 게이트 제외'
  };
  const REASON_TEXT = {
    domestic: '국내 항만은 배분 대상이 아닙니다.', 'no-volume': '물동량 값이 없어 계산하지 않았습니다.',
    'below-min': '물동량 1 TEU 미만 항만은 계산하지 않았습니다.',
    'no-climate': '전년 동월 기후 자료가 없어 기후 게이트·위험도를 계산할 수 없습니다.',
    'no-suit': '전년 동월 서식 적합도 자료가 없습니다.',
    gate: '인천항과 기후가 크게 달라(기후 유사도가 기준 미만) 검사 대상에서 제외했습니다.'
  };

  function shiftMonth(value, offset) {
    const [year, month] = value.split('-').map(Number);
    const date = new Date(Date.UTC(year, month - 1 + offset, 1));
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
  }
  const sum = (array, pick = value => value) => array.reduce((total, item, index) => total + pick(item, index), 0);

  /**
   * records: [{ id, country, volume }]  — 한 달치 항만 목록
   * ref: QUARANTINE_REFERENCE, model: QUARANTINE_MODEL_PARAMS
   * suitability: { 항만ID: [종별 S] } | null (null이면 평가 지표만 계산하지 않음)
   */
  function buildContext({ records, period, ref, model, suitability = null, options = {} }) {
    const opt = { ...DEFAULTS, ...options };
    const climateMonth = shiftMonth(period, -12);
    const domestic = new Set(ref.ports.filter(port => port.domestic).map(port => port.id));
    const climateRows = new Map((ref.months[climateMonth] || []).map(row => [row[0], row.slice(3, 3 + model.fields.length)]));
    const incheon = model.krinc[climateMonth] || null;
    const excluded = [];
    const eligible = [];
    for (const record of records) {
      let reason = '';
      if (domestic.has(record.id) || record.country === 'KR') reason = 'domestic';
      else if (record.volume === null || record.volume === undefined) reason = 'no-volume';
      else if (record.volume < opt.minVolume) reason = 'below-min';
      else {
        const climate = climateRows.get(record.id);
        if (!incheon || !climate || climate.some(value => value === null || value === undefined)) reason = 'no-climate';
        else if (suitability && !suitability[record.id]) reason = 'no-suit';
      }
      if (reason) excluded.push({ id: record.id, reason }); else eligible.push({ record, climate: climateRows.get(record.id) });
    }
    const n = eligible.length;
    const ctx = {
      period, climateMonth, options: opt, ids: eligible.map(item => item.record.id), N: eligible.map(item => item.record.volume),
      country: eligible.map(item => item.record.country), excluded, hasTruth: Boolean(suitability), ready: false,
      Wg: [], gate: [], C: [], Ps: null, PsPre: null, totalN: 0, gateN: 0, yearbookYears: []
    };
    ctx.totalN = sum(ctx.N);
    if (!n) return ctx;

    // 1) 정규화 기후 φ (이번 달 항만 + 인천항 전체 Min-Max)
    const K = model.fields.length;
    const all = [...eligible.map(item => item.climate), incheon];
    const min = Array.from({ length: K }, (_, k) => Math.min(...all.map(row => row[k])));
    const max = Array.from({ length: K }, (_, k) => Math.max(...all.map(row => row[k])));
    const range = min.map((value, k) => max[k] - value);
    const phi = row => row.map((value, k) => range[k] > 0 ? (value - min[k]) / range[k] : 0);
    const phiI = phi(incheon), phiG = eligible.map(item => phi(item.climate));

    // 2) 종별 기후 유사도 W_gs, 항만 유사도 W_g, 게이트
    const weights = ref.species.map(species => species.weight);
    const totalWeight = sum(weights);
    const H = weights.map(weight => weight / totalWeight);
    const coefficient = model.beta.map((betaRow, s) => {
      const scaled = betaRow.map((beta, k) => Math.abs(beta * range[k] / model.trainRange[s][k]));
      const total = sum(scaled);
      return scaled.map(value => total > 0 ? value / total : 1 / K);
    });
    const Wgs = phiG.map(row => coefficient.map(a => Math.min(1, Math.max(0, 1 - sum(a, (weight, k) => weight * Math.abs(row[k] - phiI[k]))))));
    ctx.Wg = Wgs.map(row => sum(row, (value, s) => H[s] * value));
    ctx.gate = ctx.Wg.map(value => value >= opt.tau ? 1 : 0);
    ctx.gateN = sum(ctx.gate);
    ctx.ids.forEach((id, i) => { if (!ctx.gate[i]) ctx.excluded.push({ id, reason: 'gate' }); });

    // 3) 연보 사전 C_g
    const years = Object.keys(model.yearbook).map(Number).filter(year => year >= 2016 && year < Number(period.slice(0, 4))).sort((a, b) => a - b);
    ctx.yearbookYears = years;
    const countries = [...new Set(ctx.country)];
    const q = Object.fromEntries(countries.map(code => [code, 0]));
    for (const year of years) {
      const counts = model.yearbook[String(year)], total = sum(Object.values(counts));
      for (const code of countries) q[code] += (year - 2015) * ((counts[code] || 0) / total);
    }
    const noPrior = !years.length;
    const powN = ctx.N.map(value => Math.pow(value, opt.alpha));
    const countrySum = {};
    ctx.ids.forEach((id, i) => { countrySum[ctx.country[i]] = (countrySum[ctx.country[i]] || 0) + powN[i]; });
    const C0 = ctx.ids.map((id, i) => (noPrior ? 1 : q[ctx.country[i]]) * powN[i] / countrySum[ctx.country[i]]);
    const gatedC0 = sum(C0, (value, i) => ctx.gate[i] * value);
    ctx.C = ctx.gate.length && ctx.gateN
      ? C0.map((value, i) => ctx.gate[i] ? opt.floor / ctx.gateN + (gatedC0 > 0 ? (1 - opt.floor) * value / gatedC0 : (1 - opt.floor) / ctx.gateN) : 0)
      : C0.map(() => 0);

    // 4) 가상 정답 p* (평가용; 운영 위험도와 재료가 다름)
    if (suitability && ctx.gateN) {
      const S = ctx.ids.map(id => suitability[id]);
      const species = H.length;
      const share = Array.from({ length: n }, () => new Array(species).fill(0));
      for (let s = 0; s < species; s++) {
        const raw = ctx.ids.map((id, i) => powN[i] * Wgs[i][s] * S[i][s]);
        const total = sum(raw);
        if (total > 0) raw.forEach((value, i) => { share[i][s] = value / total; });
      }
      const score = share.map(row => sum(row, (value, s) => H[s] * value));
      const scoreTotal = sum(score), gatedTotal = sum(score, (value, i) => ctx.gate[i] * value);
      ctx.PsPre = scoreTotal > 0 ? score.map(value => value / scoreTotal) : null;
      ctx.Ps = gatedTotal > 0 ? score.map((value, i) => ctx.gate[i] * value / gatedTotal) : null;
    }
    ctx.ready = ctx.gateN > 0;
    return ctx;
  }

  /** γ, B 로 검사 비율·검사량 계산. 정수 검사량은 최대잔여법. */
  function allocate(ctx, gamma, budget) {
    const n = ctx.ids.length;
    const active = [];
    for (let i = 0; i < n; i++) if (ctx.gate[i] && ctx.C[i] > 0) active.push(i);
    if (!ctx.ready || !active.length) return null;
    const totalN = sum(active, i => ctx.N[i]);
    const B = Math.max(0, Math.min(Math.round(budget), Math.floor(totalN + 1e-9)));
    const density = new Array(n).fill(0);
    for (const i of active) density[i] = Math.pow(ctx.C[i] / ctx.N[i], gamma);
    const dmin = Math.min(...active.map(i => density[i]));
    let lo = 0, hi = 1 / dmin * (1 + 1e-9);
    for (let step = 0; step < 200; step++) {
      const theta = (lo + hi) / 2;
      let used = 0;
      for (const i of active) used += ctx.N[i] * Math.min(1, theta * density[i]);
      if (used > B) hi = theta; else lo = theta;
    }
    const I = new Array(n).fill(0), c = new Array(n).fill(0);
    for (const i of active) { c[i] = Math.min(1, lo * density[i]); I[i] = ctx.N[i] * c[i]; }
    const cap = ctx.N.map(Math.ceil);
    const Iint = I.map((value, i) => Math.min(Math.floor(value + 1e-9), cap[i]));
    let remainder = B - sum(Iint);
    const order = active.slice().sort((a, b) => (I[b] - Math.floor(I[b])) - (I[a] - Math.floor(I[a])));
    for (const i of order) {
      if (remainder <= 0) break;
      if (Iint[i] < cap[i] && I[i] - Math.floor(I[i]) > 1e-9) { Iint[i]++; remainder--; }
    }
    for (const i of order) { // 소수 부분이 없는 항만까지 써도 합계가 모자라면 남은 칸에 채움
      if (remainder <= 0) break;
      if (Iint[i] < cap[i]) { Iint[i]++; remainder--; }
    }
    return { gamma, budget: B, requested: budget, theta: lo, c, I, Iint, capped: B < Math.round(budget), totalN };
  }

  /** 평가 지표 (정답 p* 필요). capture = 유입 차단율, foh = 최선 배분 대비, tv = 배분 오차 */
  function metrics(ctx, result) {
    if (!result) return null;
    const total = sum(result.Iint);
    const base = { inspected: result.Iint.filter(value => value > 0).length, total, ports: result.Iint.filter((value, i) => ctx.gate[i]).length };
    if (!ctx.Ps) return { ...base, tv: null, capture: null, best: null, foh: null, baseline: null, baselineFoh: null, ratio: null };
    const Ps = ctx.Ps, N = ctx.N;
    const tv = 0.5 * sum(Ps, (value, i) => Math.abs((total ? result.Iint[i] / total : 0) - value));
    const capture = sum(Ps, (value, i) => Math.min(1, result.Iint[i] / N[i]) * value);
    let left = result.budget, best = 0;
    for (const i of Ps.map((value, i) => i).sort((a, b) => Ps[b] / N[b] - Ps[a] / N[a])) {
      if (left <= 0) break;
      const take = Math.min(N[i], left); best += take / N[i] * Ps[i]; left -= take;
    }
    const baseline = Math.min(1, result.budget / ctx.totalN); // 현행: 전체 물동량에서 무작위 추출
    return { ...base, tv, capture, best, foh: best > 0 ? capture / best : null, baseline, baselineFoh: best > 0 ? baseline / best : null, ratio: baseline > 0 ? capture / baseline : null };
  }

  function sweepGamma(ctx, budget, gammas) {
    return gammas.map(gamma => { const result = allocate(ctx, gamma, budget); return { gamma, ...(metrics(ctx, result) || {}) }; });
  }
  function sweepBudget(ctx, gamma, budgets) {
    return budgets.map(budget => { const result = allocate(ctx, gamma, budget); return result ? { budget, ...metrics(ctx, result), capped: result.capped } : { budget }; });
  }

  const api = { DEFAULTS, REASONS, REASON_TEXT, shiftMonth, buildContext, allocate, metrics, sweepGamma, sweepBudget };
  root.QuarantineAllocation = api;
  if (typeof module === 'object' && module.exports) module.exports = api;
})(globalThis);
