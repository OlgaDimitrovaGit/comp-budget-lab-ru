import { useState, type ReactNode } from 'react';
import { Info } from 'lucide-react';
import { Button } from '@/components/ui/button';
const MINUS = '−';
export const rub = (n: number) => new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 }).format(Math.round(n)).replace('-', MINUS);
export const million = (n: number) => `${n < 0 ? MINUS : ''}${(Math.abs(n) / 1e6).toFixed(2).replace('.', ',')} млн ₽`;
/** Explicit sign: +1,50 млн ₽, −1,50 млн ₽; zero without a sign. */
export const signed = (s: string, n: number) => (n > 0 ? '+' : '') + s;
export function Detail({ title, label, children }: { title?: string; label?: string; children: ReactNode }) {
 const [open, setOpen] = useState(false);
 return <span className={`info ${open ? 'is-open' : ''}`} onKeyDown={e => { if (e.key === 'Escape') setOpen(false); }}><Button variant="ghost" size="icon" aria-label={`Подробнее: ${title ?? label ?? 'пояснение'}`} aria-expanded={open} onClick={() => setOpen(v => !v)} onBlur={() => setOpen(false)}><Info /></Button><span className="info-box" role="tooltip">{title && <strong>{title}</strong>}{children}</span></span>;
}
export function Field({ label, hint, detail, value, onChange, options, suffix }: { label: string; hint?: string; detail?: string; value: string | number; onChange: (value: string) => void; options?: (string | number)[]; suffix?: string }) {
 const control = options ? <select aria-label={label} value={value} onChange={e => onChange(e.target.value)}>{options.map(o => <option key={o} value={o}>{o}{typeof o === 'number' ? suffix : ''}</option>)}</select> : <input aria-label={label} type="text" inputMode="decimal" value={value} onChange={e => onChange(e.target.value)} />;
 const hintLine = hint && <span className="field-hint">{hint}</span>;
 // a label around the (i) button would pass its clicks to the button, so a field with details is a div
 if (detail) return <div className="field"><span className="field-label">{label}<Detail label={label}>{detail}</Detail></span>{control}{hintLine}</div>;
 return <label className="field"><span className="field-label">{label}</span>{control}{hintLine}</label>;
}
export const ratingNames: Record<number,string> = {5:'Значительно превосходит',4:'Превосходит',3:'Соответствует',2:'Частично соответствует',1:'Не соответствует'};
