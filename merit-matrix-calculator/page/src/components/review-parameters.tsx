import { useEffect, useRef } from 'react';
import { ChevronDown } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { MAX_INCREASE, MONTH_STEPS, readNumber, REVIEW_DATES, type Params } from '@/model';
import { Detail, Field, ratingNames } from './review-shared';
export function ReviewParameters({ params, setParams }: { params: Params; setParams: (f: (p: Params) => Params) => void }) {
 const ref = useRef<HTMLDetailsElement>(null);
 useEffect(() => { const media = window.matchMedia('(min-width: 768px)'); const update=()=> {if(ref.current) ref.current.open=media.matches;}; update(); media.addEventListener('change',update); return()=>media.removeEventListener('change',update); },[]);
 const set = (k: 'budget'|'date'|'target'|'max'|'service'|'change') => (v:string) => setParams(x=>({...x,[k]:v}));
 const setRating = (i:number, k:'target'|'min', v:string) => setParams(x=>({...x,ratings:x.ratings.map((r,j)=>j===i?{...r,[k]:v}:r)}));
 return <aside className="parameters"><details open ref={ref}><summary className="parameter-toggle"><span>Параметры пересмотра<span className="toggle-summary">Введите свои значения</span></span><ChevronDown /></summary><p className="panel-subtitle">введите свои значения</p>
 <Field label="Целевой бюджет повышения" hint="% от окладов участников" value={params.budget} onChange={set('budget')} />
 <Field label="Дата пересмотра" detail="Новые оклады действуют с 1-го числа этого месяца. Право на пересмотр считается на эту дату." value={params.date} onChange={set('date')} options={REVIEW_DATES} />
 <Field label="Политика оплаты" hint="% от середины вилки" value={params.target} onChange={set('target')} />
 <Field label="Потолок повышения" detail="% от оклада до пересмотра. Больше не дадим, даже если оклад останется ниже минимума вилки." value={params.max} onChange={set('max')} options={MAX_INCREASE} suffix="%" />
 <Field label="Минимальный стаж на дату пересмотра" value={params.service} onChange={set('service')} options={MONTH_STEPS} suffix=" мес." />
 <Field label="Без изменения оклада не менее" detail="Сколько месяцев прошло с последнего индивидуального изменения оклада, на дату пересмотра." value={params.change} onChange={set('change')} options={MONTH_STEPS} suffix=" мес." />
 <div className="segment-row"><span>Не выше максимума вилки</span><div className="segment" role="group" aria-label="Не выше максимума вилки">{[true,false].map(v=><Button variant="ghost" key={String(v)} aria-pressed={params.clip===v} onClick={()=>setParams(x=>({...x,clip:v}))}>{v?'Да':'Нет'}</Button>)}</div></div>
 <div className="parameter-section"><div className="parameter-head"><h3>Целевой CR по рейтингу</h3><Detail label="целевой CR по рейтингу"><p>CR (compa-ratio) — оклад к середине вилки. Целевой CR задан в % от политики оплаты, а не от середины вилки.</p><p className="mt-2">Целевой оклад = середина вилки × политика оплаты × целевой CR рейтинга.</p><p className="mt-2 text-muted-foreground">Пример: середина 3 000 000 ₽, политика {params.target.replace('.',',')}%, «Значительно превосходит» {params.ratings[0].target.replace('.',',')||'0'}% → целевой оклад {Math.round(3000000*readNumber(params.target)/100*readNumber(params.ratings[0].target,0)/100).toLocaleString('ru-RU')} ₽.</p></Detail></div><table className="rating-input-table"><thead><tr><th>Рейтинг</th><th>Целевой CR<small>% от политики<br/>оплаты</small></th><th>Мин. повышение<small>% от оклада</small></th></tr></thead><tbody>{params.ratings.map((r,i)=><tr key={r.rating}><td>{ratingNames[r.rating]}</td><td><input aria-label={`${ratingNames[r.rating]}: целевой CR`} type="text" inputMode="decimal" value={r.target} onChange={e=>setRating(i,'target',e.target.value)} /></td><td><input aria-label={`${ratingNames[r.rating]}: минимальное повышение`} type="text" inputMode="decimal" value={`${r.min}%`} onChange={e=>setRating(i,'min',e.target.value.replace('%',''))} /></td></tr>)}</tbody></table><p className="field-hint">Пустой целевой CR — без повышения при этом рейтинге</p></div>
 </details></aside>;
}
