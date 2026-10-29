import { useEffect, useState } from 'react';
import { period, useData } from '@/model';
import { Detail, million, rub, ratingNames } from './review-shared';
const ZONES = ['Ниже 90%','90–100%','100–110%','110% и выше'];
/** Axis: four equal steps from a round number below the start up to at least the end.
 * The start stays at least half a step above the axis, so the first bar is visible. */
function axis(start:number, end:number) {
 const k0=Math.floor(Math.log10(Math.max(end,1)))-6;
 for (let k=k0;k<k0+12;k++) for (const m of [1,2,2.5,4,5]) {
  const step=m*10**k;
  let min=Math.floor((start-1e-6)/step)*step;
  if (start-min<step/2) min-=step;
  if (min>=0 && min+4*step>=end) return {min,step,ticks:[0,1,2,3,4].map(i=>min+i*step)};
 }
 const step=end/4||1;
 return {min:0,step,ticks:[0,1,2,3,4].map(i=>i*step)};
}
function useNarrow() {
 const query = '(max-width: 600px)';
 const [narrow,setNarrow]=useState(()=>window.matchMedia(query).matches);
 useEffect(()=>{const media=window.matchMedia(query);const update=()=>setNarrow(media.matches);media.addEventListener('change',update);return()=>media.removeEventListener('change',update);},[]);
 return narrow;
}
export function Waterfall() {
 const data = useData();
 const p = period(data);
 const current = data.budget.payroll_before;
 const merit = data.merit.in_year;
 const inYear = current+merit;
 const rest = data.budget.change_total_annual-data.budget.change_total_in_year;
 const full = inYear+rest;
 const steps = [
  {name:'ФОТ со взносами сейчас',short:['Сейчас','со взн.'],amount:current,kind:'total',start:0,end:current},
  {name:`Повышение, ${p.span}`,short:['Повышение'],amount:merit,kind:'merit',start:current,end:inYear},
  {name:`ФОТ со взносами ${p.year}`,short:[String(p.year),'со взн.'],amount:inYear,kind:'total',start:0,end:inYear},
  {name:'Эффект полного года',short:['Эффект','полного года'],amount:rest,kind:'annual',start:inYear,end:full},
  {name:`ФОТ со взносами ${p.next}`,short:[String(p.next),'со взн.'],amount:full,kind:'total',start:0,end:full},
 ];
 const ax = axis(current, full)!;
 const y=(n:number)=>225-(Math.max(n,ax.min)-ax.min)/(4*ax.step)*175;
 const narrow = useNarrow();
 // phone: the same five bars, narrower, so the chart fits the screen without scrolling
 const g = narrow ? {w:332,x0:22,step:64,bar:40,brk:2} : {w:800,x0:44,step:156,bar:84,brk:15};
 // the unit is in the subtitle, so the bars carry bare numbers
 const value=(s:typeof steps[number])=>`${s.kind==='total'?'':'+'}${(s.amount/1e6).toFixed(2).replace('.',',')}`;
 const label=(s:typeof steps[number])=>s.short;
 return <section className="section waterfall-section"><div className="section-head"><div><h2>Рост ФОТ со взносами: {p.year} и эффект полного года в {p.next}</h2><p className="subtitle">Млн ₽. В {p.next} — только переходящий эффект пересмотра {p.year}. Пересмотр {p.next} не учтён.</p></div><Detail>Здесь ФОТ — оклады и переменная часть плюс страховые взносы работодателя. Повышение считается с месяца пересмотра: {p.span}, {p.monthsText}. Эффект полного года — те же повышения за месяцы, не попавшие в {p.year}. ФОТ {p.next} — текущий ФОТ плюс этот пересмотр за 12 месяцев. Пересмотр {p.next} не учтён.</Detail></div>
 <div className="chart-surface">
 <div className="waterfall-scroll"><svg className={`waterfall ${narrow?'narrow':''}`} viewBox={`0 22 ${g.w} 261`} aria-label={`Каскад: ФОТ со взносами растёт с ${million(current)} до ${million(inYear)} в ${p.year} и до ${million(full)} в ${p.next}`} role="img">
 <line className="chart-base" x1={g.brk-1} x2={g.w-2} y1="225" y2="225"/>
 {/* the scale starts above zero: the break mark says the base of the bars is cut */}
 <g className="axis-break"><rect x={g.brk+2} y="218" width="9" height="14"/><line x1={g.brk} x2={g.brk+7} y1="231" y2="219"/><line x1={g.brk+6} x2={g.brk+13} y1="231" y2="219"/></g>
 {steps.map((s,i)=>{const x=g.x0+i*g.step;const c=x+g.bar/2;const end=y(s.end);const bottom=y(s.start);return <g key={s.name} aria-label={`${s.name}: ${rub(s.amount)}`}>
 {i<steps.length-1&&<line className="connector" x1={x+g.bar} x2={x+g.step} y1={end} y2={end}/>}
 <rect className={`bar-${s.kind}`} x={x} y={end} width={g.bar} height={Math.max(3,bottom-end)} rx="1" />
 <text className="bar-value" textAnchor="middle" x={c} y={end-9}>{value(s)}</text>
 {label(s).map((l,j)=><text key={j} className="bar-label" textAnchor="middle" x={c} y={246+j*12}>{l}</text>)}
 </g>;})}</svg></div>
 </div></section>;
}
export function MeritMatrix() {
 const data = useData();
 const [active,setActive]=useState<{rating:number;zone:number;n:number;rub:number;pct:number}|null>(null);
 return <section className="section matrix-section"><div className="section-head"><div><h2>Матрица повышений: среднее повышение участников, % от оклада</h2></div><Detail>Среднее повышение в % от оклада до пересмотра. В каждой клетке — только участники пересмотра. «—» — участников в клетке нет; 0% — участники в клетке повышения не получают.</Detail></div>
 <div className="chart-surface">{active&&<div className="chart-popover" role="status"><strong>{ratingNames[active.rating]} · {ZONES[active.zone]}</strong><div className="detail-line"><span>Среднее повышение</span><b>{active.pct.toFixed(1).replace('.',',')}%</b></div><div className="detail-line"><span>Участников</span><b>{active.n}</b></div><div className="detail-line"><span>Прибавка к окладам за год</span><b>{million(active.rub)}</b></div></div>}
 <div className="matrix-wrap"><table className="matrix"><thead><tr><th>Рейтинг / CR до пересмотра</th>{ZONES.map(z=><th key={z}>{z}</th>)}</tr></thead><tbody>{[...data.matrix.rows].reverse().map(row=><tr key={row.rating}><th>{ratingNames[row.rating]}</th>{row.cells.map((cell,i)=>{const heat=cell.pct===0?0:cell.pct<3?1:cell.pct<8?2:cell.pct<16?3:cell.pct<22?4:5;return <td key={i} className={cell.n===0?'heat-empty':heat===0?'heat-zero':`heat-${heat}`} tabIndex={cell.n?0:undefined} aria-label={`${ratingNames[row.rating]}, ${ZONES[i]}, участников: ${cell.n}, ${cell.pct.toFixed(1).replace('.',',')}%`} onMouseEnter={()=>cell.n&&setActive({rating:row.rating,zone:i,...cell})} onMouseLeave={()=>setActive(null)} onFocus={()=>setActive({rating:row.rating,zone:i,...cell})} onBlur={()=>setActive(null)} onClick={()=>cell.n&&setActive(active?.rating===row.rating&&active.zone===i?null:{rating:row.rating,zone:i,...cell})}>{cell.n>0?(heat===0?'0%':`${cell.pct.toFixed(1).replace('.',',')}%`):'—'}</td>;})}</tr>)}</tbody></table></div>
 <div className="chart-footer matrix-footer"><div className="heat-legend section-note"><span>0%</span>{[1,2,3,4,5].map(i=><i key={i} className={`heat-${i}`}/>)}<span>25%</span></div><span className="section-note">— нет участников</span></div></div></section>;
}
