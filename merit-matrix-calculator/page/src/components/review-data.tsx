import { useRef, useState } from 'react';
import { ChevronDown, Download, FileSpreadsheet, RotateCcw, Upload } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { parseCsv, EMPLOYEE_COLUMNS, BAND_COLUMNS } from '@/engine.js';
import { compute, period, toSettings, useData, type Params, type Source } from '@/model';
import { Field } from './review-shared';
const csvHref = (text: string) => `data:text/csv;charset=utf-8,${encodeURIComponent(text)}`;
function method(y: number) {
 return [
  { title: 'Повышение', items: [
   'Целевой оклад = середина вилки × политика оплаты × целевой CR рейтинга. Пример: середина 3 000 000 ₽, политика 100%, «Значительно превосходит» 110% — целевой оклад 3 300 000 ₽.',
   'Участнику с окладом ниже целевого оклад поднимается до целевого. Оклад выше целевого не снижается.',
   'Затем по порядку действуют три ограничения: не меньше минимального повышения для рейтинга; не больше потолка повышения; не выше максимума вилки, если это ограничение включено. Сотрудник на максимуме вилки и выше повышения не получает.',
   'Модель не подгоняет повышения под целевой бюджет. Карточки показывают разницу. Чтобы приблизиться к цели, меняйте целевой CR по рейтингам.',
   'Целевая стоимость в рублях: целевой процент от оклада каждого участника плюс переменная часть и страховые взносы на это повышение.',
  ] },
  { title: 'Сроки', items: [
   `Новые оклады действуют с даты пересмотра. В ${y} учтены только месяцы с этой даты.`,
   `12 месяцев ${y + 1}: этот пересмотр за полный год. Пересмотр ${y + 1} не учтён.`,
  ] },
  { title: 'Допущения', items: [
   'ФОТ — оклады и переменная часть (премия в % от оклада). Там, где сказано «ФОТ со взносами», добавлены страховые взносы работодателя. Суммы «на оклады» — без переменной части и взносов.',
   `Тарифы взносов 30% и 15,1% заданы НК РФ; предельная база на ${y}–${y + 1} годы ещё не утверждена — взята база 2026 года. Взносы на травматизм не включены.`,
   'Вилки считаются актуальными. Коллективные договоры и районные коэффициенты не моделируются.',
  ] },
 ];
}
export function ReviewData({open,onToggle,source,demo,setSource,params,setParams}:{open:boolean;onToggle:()=>void;source:Source;demo:Source;setSource:(s:Source)=>void;params:Params;setParams:(f:(p:Params)=>Params)=>void}) {
 const data=useData();
 const p=period(data);
 const input=useRef<HTMLInputElement>(null);
 const [notice,setNotice]=useState('');const [error,setError]=useState(false);
 const change=(key:keyof Params['contrib'])=>(v:string)=>setParams(x=>({...x,contrib:{...x.contrib,[key]:v}}));
 async function load(files:FileList|null){
  if(!files||!files.length)return;setError(false);
  try{
   const next={...source};const names:string[]=[];let ownEmployees=false;let bandsLoaded=false;
   for(const file of Array.from(files)){
    if(file.size>10*1024*1024)throw new Error(`${file.name}: выберите CSV меньше 10 МБ.`);
    const text=await file.text();const head=(parseCsv(text)[0]??[]).map(h=>h.trim().toLowerCase());
    if(head.includes('mid')&&head.includes('max')){next.bands=text;bandsLoaded=true;}
    else if(head.includes('base_salary')){next.employees=text;ownEmployees=true;}
    else throw new Error(`${file.name}: это не файл сотрудников (столбец base_salary) и не файл вилок (столбцы mid и max). Формат — как в образцах CSV.`);
    names.push(file.name);
   }
   if(ownEmployees&&!bandsLoaded&&next.bands===demo.bands)throw new Error('Загрузите вилки вместе с файлом сотрудников: выберите оба файла сразу.');
   next.name=names.join(' + ');
   const settings=toSettings(params);
   const result=settings?compute(next,settings):null;
   setSource(next);
   setNotice(result?`${next.name} · загружено сотрудников: ${result.headcount}.`:`${next.name} загружен.`);
  }catch(e){setError(true);setNotice(e instanceof Error?e.message:'Не удалось прочитать файл.');}
 }
 return <section className="data-section" id="data-section"><Button variant="ghost" className="data-toggle" aria-expanded={open} aria-controls="data-content" onClick={onToggle}><span className="flex min-w-0 items-center gap-3"><FileSpreadsheet className="shrink-0 text-muted-foreground"/><span className="data-title">Свои данные, Excel-модель, взносы и метод</span></span><ChevronDown className={`shrink-0 transition-transform ${open?'rotate-180':''}`}/></Button>
 {open&&<div className="data-body" id="data-content"><div className="data-columns"><div><h3>Посчитать на своих данных</h3><ol className="data-steps">
 <li>Скачайте два образца: сотрудники и вилки.<div className="data-tools"><Button variant="outline" size="sm" asChild><a href={csvHref(demo.employees)} download="sample-employees.csv"><Download/>Образец: сотрудники</a></Button><Button variant="outline" size="sm" asChild><a href={csvHref(demo.bands)} download="sample-salary-ranges.csv"><Download/>Образец: вилки</a></Button></div></li>
 <li>Замените строки своими данными. Названия столбцов не меняйте.<p>Сотрудники: {EMPLOYEE_COLUMNS.join(', ')}. Оклад — годовой, в рублях; bonus_pct — переменная часть в % от оклада; fte — ставка (1 — полная); rating — от 1 до 5; даты — ГГГГ-ММ-ДД. category — подразделение или другая группа, для которой заданы вилки. Без столбца category вилки подбираются только по грейду.</p><p>Вилки: {BAND_COLUMNS.join(', ')} — годовые суммы в рублях.</p></li>
 <li>Нажмите «Загрузить CSV» и выберите оба файла сразу.<div className="data-tools"><Button size="sm" onClick={()=>input.current?.click()}><Upload/>Загрузить CSV</Button></div></li>
 <li>Страница пересчитается на ваших данных, имена файлов появятся вверху. Параметры пересмотра меняйте как нужно.</li>
</ol><input className="hidden" ref={input} type="file" multiple accept=".csv,text/csv" aria-label="Загрузить CSV сотрудников и вилок" onChange={e=>{void load(e.target.files);e.target.value='';}}/><p>Файлы остаются в вашем браузере. Страница никуда их не отправляет.</p>{notice&&<div role="status" className={`import-result ${error?'error':''}`}>{notice}</div>}<div className="data-tools"><Button variant="outline" size="sm" onClick={()=>{setSource(demo);setError(false);setNotice(`Возвращены демо-данные · смоделированных сотрудников: ${parseCsv(demo.employees).length-1}.`);}}><RotateCcw/>Вернуть демо-данные</Button></div>
 <h3 className="mt-6">Проверить расчёт в Excel</h3><p>Excel-модель — тот же калькулятор в формулах, строка на сотрудника: каждую цифру можно проследить. В ней демо-данные и параметры по умолчанию. Чтобы посчитать своих сотрудников, вставьте их на лист «Сотрудники», а вилки — на лист «Вилки»: формулы рассчитаны на 2 000 сотрудников, подразделения подтянутся из данных. Ваши файлы и значения, введённые на странице, в модель не попадают.</p><Button variant="outline" size="sm" asChild><a href="merit-model.xlsx" download><Download/>Скачать Excel-модель</a></Button></div>
 <div><h3>Страховые взносы работодателя</h3><p>В демо — общие ставки России на 2026 год. Свои ставки введите ниже.</p><div className="contribution-grid"><Field label="Ставка до предельной базы, %" value={params.contrib.rate_below} onChange={change('rate_below')}/><Field label="Ставка сверх предельной базы, %" value={params.contrib.rate_above} onChange={change('rate_above')}/><Field label="Предельная база за год, ₽" value={params.contrib.ceiling} onChange={change('ceiling')}/></div><p>Предельная база действует на каждого сотрудника отдельно: на его годовой оклад с переменной частью. Пониженные тарифы (МСП, IT-компании) — введите свои ставки.</p><p className="section-note">Россия, 2026: 30% до 2 979 000 ₽ в год на человека, 15,1% сверх. Источник: постановление Правительства РФ от 31.10.2025 № 1705 (предельная база), ст. 425 НК РФ (тарифы). Взносы на травматизм не включены. Ставки на {p.year} и {p.next} ещё не утверждены — взяты ставки 2026.</p></div></div>
 <div className="method">{method(p.year).map(s=><div key={s.title}><h3>{s.title}</h3><ul>{s.items.map(t=><li key={t}>{t}</li>)}</ul></div>)}</div></div>}</section>;
}
