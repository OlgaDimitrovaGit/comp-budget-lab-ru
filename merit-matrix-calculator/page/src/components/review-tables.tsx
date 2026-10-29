import { useData, type Data } from '@/model';
import { ALL_EMPLOYEES } from '@/engine.js';
import { Detail, rub } from './review-shared';
type Cat = Data['categories'][number];
const cr = (n:number) => `${Math.floor(n+0.5)}%`;
function Name({c}:{c:Cat}) { return <span className="category-name" tabIndex={0}>{c.name}<span className="category-tooltip" role="tooltip"><b>{c.name}</b>Сотрудников: {c.n}<br/>Прибавка к окладам за год: {rub(c.merit_cost_base)}{c.small&&<><br/><em>Меньше 10 человек: один сотрудник заметно сдвигает CR.</em></>}</span></span>; }
export function CategoryTables() {
 const data = useData();
 const order = [...data.categories].sort((a,b)=>a.cr_before-b.cr_before);
 const single = data.categories.length===1 && data.categories[0].name===ALL_EMPLOYEES;
 const low = data.cr_all.low_cr_pct;
 return <div className="tables-grid single"><section className="section"><div className="section-head"><div><h2>CR по подразделениям: до и после пересмотра</h2><p className="subtitle">Оклад в % от середины вилки</p></div><Detail title="CR (compa-ratio)">Оклад, делённый на середину вилки, по подразделению — взвешенно по ставке. Вилки считаются актуальными. Цифры по подразделению показывают положение в вилках, а не рекомендации по конкретным окладам.</Detail></div><table className="category-table"><thead><tr><th>Подразделение</th><th>CR сейчас</th><th>CR после<br/>пересмотра</th><th>Ниже {low}%<br/>сейчас → после</th><th>Прибавка к окладам,<br/>за год</th></tr></thead><tbody>{!single&&order.map(c=><tr key={c.name}><th scope="row"><Name c={c}/></th><td>{cr(c.cr_before)}</td><td>{cr(c.cr_after_merit)}</td><td>{c.below_low_before} → {c.below_low_after}</td><td>{rub(c.merit_cost_base)}</td></tr>)}<tr className="total"><th>{single?'Все сотрудники':'Итого'}</th><td>{cr(data.cr_all.cr_before)}</td><td>{cr(data.cr_all.cr_after_merit)}</td><td>{data.cr_all.below_low_before} → {data.cr_all.below_low_after}</td><td>{rub(data.merit.base)}</td></tr></tbody></table></section></div>;
}
export function ReviewPeople() {
 const data = useData();
 const f = data.flags;
 const stat = (n:string|number, text:string, warn=false, zero=false) => <div className={`people-stat ${warn?'warning':''} ${zero?'is-zero':''}`}><b>{n}</b><span>{text}</span></div>;
 return <section className="section people-section"><div className="section-head"><div><h2>Кого посмотреть отдельно</h2></div><Detail title="Кого посмотреть отдельно">Демо-данные: табельные номера — в Excel-модели, лист «Расчёт».</Detail></div>
 <div className="people-strip">{stat(`${data.cr_all.below_low_before} → ${data.cr_all.below_low_after}`, `ниже ${data.cr_all.low_cr_pct}% середины вилки: сейчас → после пересмотра`)}{stat(f.red_circled, 'на максимуме вилки и выше: без повышения', false, f.red_circled===0)}{stat(f.below_min_after_merit, 'ниже минимума вилки после пересмотра', f.below_min_after_merit>0, f.below_min_after_merit===0)}{stat(f.not_eligible, 'не участвуют: стаж или недавнее изменение оклада', false, f.not_eligible===0)}</div></section>;
}
