'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  BookOpen,
  Check,
  Clipboard,
  Download,
  FileText,
  Printer,
  ShieldCheck,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Checkbox } from '@/components/ui/checkbox';
import { NativeSelect, NativeSelectOption } from '@/components/ui/native-select';
import { Textarea } from '@/components/ui/textarea';
import { viduiItems } from './vidui-data';

type Entry = { selected: number[]; custom: string; kavanot: number[] };
type Notes = Record<string, Entry>;
type View = 'work' | 'summary';

const emptyEntry = (): Entry => ({ selected: [], custom: '', kavanot: [] });

function hasContent(entry?: Entry) {
  return Boolean(entry && (entry.selected.length || entry.kavanot.length || entry.custom.trim()));
}

function compileDocument(notes: Notes) {
  const sections = viduiItems.flatMap((item, index) => {
    const entry = notes[item.id];
    if (!hasContent(entry)) return [];

    const kavanot = entry.kavanot.map((choice) => item.kavanot[choice]).filter(Boolean);
    const examples = entry.selected.map((choice) => item.examples[choice]).filter(Boolean);
    const lines = [
      `${index + 1}. ${item.hebrew} — ${item.translit}`,
      item.translation,
      ...(kavanot.length ? ['', 'О чём я хочу думать:', ...kavanot.map((line) => `• ${line}`)] : []),
      ...(examples.length || entry.custom.trim()
        ? ['', 'Моя личная конкретизация:', ...examples.map((line) => `• ${line}`), ...(entry.custom.trim() ? [`• ${entry.custom.trim()}`] : [])]
        : []),
    ];
    return [lines.join('\n')];
  });

  return [
    'МОЙ ВИДУЙ',
    'Личная подготовка к Йом-Кипуру',
    '',
    ...sections.flatMap((section) => [section, '', '—', '']),
    'Формулировки для личного размышления; текст не заменяет махзор и консультацию с раввином.',
  ].join('\n');
}

export default function Home() {
  const [active, setActive] = useState(0);
  const [notes, setNotes] = useState<Notes>({});
  const [hydrated, setHydrated] = useState(false);
  const [view, setView] = useState<View>('work');
  const [message, setMessage] = useState('');
  const notesRef = useRef(notes);
  const current = viduiItems[active];

  useEffect(() => {
    try {
      const saved = window.localStorage.getItem('vidui-notes-v1');
      if (saved) setNotes(JSON.parse(saved));
    } catch {
      // The workbook still works when private browser storage is unavailable.
    } finally {
      setHydrated(true);
    }
  }, []);

  useEffect(() => {
    notesRef.current = notes;
    if (!hydrated) return;
    try {
      window.localStorage.setItem('vidui-notes-v1', JSON.stringify(notes));
    } catch {
      // Keep the current session usable without persistent browser storage.
    }
  }, [hydrated, notes]);

  useEffect(() => {
    const context = document.modelContext;
    if (!context?.registerTool) return;
    const lifecycle = new AbortController();
    const register = (tool: WebMcpTool) => {
      try {
        void Promise.resolve(context.registerTool(tool, { signal: lifecycle.signal })).catch(() => undefined);
      } catch {
        // WebMCP is an optional enhancement; the visible interface remains primary.
      }
    };

    register({
      name: 'open_vidui_item',
      title: 'Открыть часть видуя',
      description: 'Open one of the 24 Vidui sections in the visible workbook.',
      inputSchema: {
        type: 'object',
        properties: { item_id: { type: 'string', enum: viduiItems.map((item) => item.id) } },
        required: ['item_id'],
        additionalProperties: false,
      },
      annotations: { readOnlyHint: true, untrustedContentHint: false },
      execute(input) {
        const itemId = (input as { item_id?: unknown })?.item_id;
        const index = viduiItems.findIndex((item) => item.id === itemId);
        if (index < 0) throw new Error('Unknown Vidui item');
        setActive(index);
        setView('work');
        return { item_id: itemId, position: index + 1 };
      },
    });

    register({
      name: 'set_vidui_entry',
      title: 'Заполнить личную запись',
      description: 'Set selected reflections and a private personal note for one Vidui section.',
      inputSchema: {
        type: 'object',
        properties: {
          item_id: { type: 'string', enum: viduiItems.map((item) => item.id) },
          kavana_indexes: { type: 'array', items: { type: 'integer', minimum: 0, maximum: 2 }, maxItems: 3 },
          example_indexes: { type: 'array', items: { type: 'integer', minimum: 0, maximum: 1 }, maxItems: 2 },
          custom: { type: 'string', maxLength: 2000 },
        },
        required: ['item_id'],
        additionalProperties: false,
      },
      annotations: { readOnlyHint: false, untrustedContentHint: true },
      execute(input) {
        const value = input as { item_id?: unknown; kavana_indexes?: unknown; example_indexes?: unknown; custom?: unknown };
        const item = viduiItems.find((candidate) => candidate.id === value.item_id);
        if (!item) throw new Error('Unknown Vidui item');
        const validateIndexes = (candidate: unknown, max: number) => {
          if (candidate === undefined) return undefined;
          if (!Array.isArray(candidate) || candidate.some((entry) => !Number.isInteger(entry) || entry < 0 || entry > max)) {
            throw new Error('Invalid selection indexes');
          }
          return [...new Set(candidate as number[])];
        };
        const kavanot = validateIndexes(value.kavana_indexes, item.kavanot.length - 1);
        const selected = validateIndexes(value.example_indexes, item.examples.length - 1);
        if (value.custom !== undefined && typeof value.custom !== 'string') throw new Error('Custom note must be text');

        setNotes((previous) => {
          const prior = previous[item.id] ?? emptyEntry();
          return {
            ...previous,
            [item.id]: {
              kavanot: kavanot ?? prior.kavanot,
              selected: selected ?? prior.selected,
              custom: typeof value.custom === 'string' ? value.custom : prior.custom,
            },
          };
        });
        return { item_id: item.id, saved: true };
      },
    });

    register({
      name: 'compile_vidui_document',
      title: 'Собрать личный видуй',
      description: 'Compile all filled Vidui reflections into one plain-text personal document.',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      annotations: { readOnlyHint: true, untrustedContentHint: true },
      execute() {
        const filled_count = viduiItems.filter((item) => hasContent(notesRef.current[item.id])).length;
        return { filled_count, document_text: compileDocument(notesRef.current) };
      },
    });

    return () => lifecycle.abort();
  }, []);

  const entry = notes[current.id] ?? emptyEntry();
  const completed = useMemo(() => viduiItems.filter((item) => hasContent(notes[item.id])).length, [notes]);
  const filledItems = useMemo(() => viduiItems.filter((item) => hasContent(notes[item.id])), [notes]);

  function update(patch: Partial<Entry>) {
    setNotes((previous) => ({ ...previous, [current.id]: { ...(previous[current.id] ?? emptyEntry()), ...patch } }));
  }

  function toggle(list: number[], index: number) {
    return list.includes(index) ? list.filter((value) => value !== index) : [...list, index];
  }

  function goTo(index: number) {
    setActive(index);
    setView('work');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function copyDocument() {
    await navigator.clipboard.writeText(compileDocument(notes));
    setMessage('Текст скопирован');
    window.setTimeout(() => setMessage(''), 2200);
  }

  function downloadDocument() {
    const blob = new Blob([compileDocument(notes)], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'moy-viduy.txt';
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="site-header sticky top-0 z-50 border-b border-navy-100 bg-white/95 backdrop-blur">
        <div className="mx-auto flex max-w-[1400px] items-center justify-between gap-3 px-3 py-3 sm:px-8 sm:py-4">
          <button type="button" className="flex min-w-0 items-center gap-2.5 text-left sm:gap-3" onClick={() => setView('work')}>
            <span className="grid size-9 shrink-0 place-items-center rounded-full bg-primary text-primary-foreground sm:size-10">
              <BookOpen className="size-4.5 sm:size-5" aria-hidden="true" />
            </span>
            <span className="min-w-0">
              <span className="block truncate font-serif text-lg font-semibold leading-none sm:text-xl">Мой видуй</span>
              <span className="mt-1 block truncate text-xs text-muted-foreground sm:text-sm">Подготовка к Йом-Кипуру</span>
            </span>
          </button>
          <div className="flex items-center gap-3">
            <div className="hidden items-center gap-2 text-sm text-muted-foreground md:flex">
              <ShieldCheck className="size-4 text-teal-700" aria-hidden="true" />
              Записи остаются на этом устройстве
            </div>
            <Button variant={view === 'summary' ? 'default' : 'outline'} size="sm" onClick={() => setView('summary')}>
              <FileText className="size-4" /> <span className="hidden sm:inline">Итоговый документ</span><span className="sm:hidden">Итог</span>
            </Button>
          </div>
        </div>
      </header>

      {view === 'work' ? (
        <div className="mx-auto grid max-w-[1400px] gap-0 lg:grid-cols-[300px_minmax(0,1fr)]">
          <aside className="sticky top-[61px] z-40 border-b border-navy-100 bg-white/95 px-3 py-3 shadow-sm backdrop-blur sm:top-[73px] sm:px-5 lg:hidden">
            <div className="flex items-center justify-between gap-4">
              <div>
                <p className="text-sm font-semibold text-primary">Часть {active + 1} из 24</p>
                <p className="text-xs text-muted-foreground">{completed} заполнено</p>
              </div>
              <span lang="he" dir="rtl" className="font-hebrew text-2xl font-semibold text-primary">{current.hebrew}</span>
            </div>
            <ViduiProgress current={active + 1} className="mt-2.5" />
            <NativeSelect className="mt-3 w-full [&_[data-slot=native-select]]:h-11 [&_[data-slot=native-select]]:text-base" value={String(active)} onChange={(event) => goTo(Number(event.target.value))} aria-label="Выберите часть видуя">
              {viduiItems.map((item, index) => (
                <NativeSelectOption key={item.id} value={String(index)}>
                  {index + 1}. {item.translit} — {item.translation}{hasContent(notes[item.id]) ? ' ✓' : ''}
                </NativeSelectOption>
              ))}
            </NativeSelect>
          </aside>

          <aside className="hidden border-r border-navy-100 bg-white px-6 py-8 lg:sticky lg:top-[73px] lg:block lg:h-[calc(100vh-73px)]">
            <div className="mb-5 flex items-center justify-between">
              <div>
                <p className="text-xs font-bold uppercase tracking-[0.15em] text-teal-700">Короткий видуй</p>
                <p className="mt-1 text-sm text-muted-foreground">{completed} из 24 заполнено</p>
              </div>
              <span className="font-serif text-2xl text-primary">{String(active + 1).padStart(2, '0')}</span>
            </div>
            <ViduiProgress current={active + 1} />
            <nav aria-label="Части видуя" className="mt-5 grid max-h-[calc(100vh-150px)] gap-2 overflow-y-auto pr-1">
              {viduiItems.map((item, index) => {
                const itemHasContent = hasContent(notes[item.id]);
                return (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => goTo(index)}
                    aria-current={active === index ? 'step' : undefined}
                    className={`group flex min-w-0 items-center gap-3 rounded-xl px-3 py-2.5 text-left transition ${
                      active === index ? 'bg-primary text-primary-foreground shadow-sm' : 'hover:bg-navy-50'
                    }`}
                  >
                    <span className={`grid size-7 shrink-0 place-items-center rounded-full text-xs font-bold ${active === index ? 'bg-white/15' : 'bg-navy-50 text-primary'}`}>
                      {itemHasContent ? <Check className="size-4" /> : index + 1}
                    </span>
                    <span className="min-w-0">
                      <span className="block truncate font-medium">{item.translit}</span>
                      <span className={`block truncate text-xs ${active === index ? 'text-white/70' : 'text-muted-foreground'}`}>{item.translation}</span>
                    </span>
                  </button>
                );
              })}
            </nav>
          </aside>

          <section className="px-3 pb-28 pt-5 sm:px-8 sm:pt-7 lg:px-12 lg:py-10">
            <div className="mx-auto max-w-4xl">
              <div className="mb-5 flex items-start justify-between gap-4 sm:mb-7 sm:gap-5">
                <div>
                  <p className="mb-2 hidden text-sm font-semibold text-teal-700 lg:block">{active + 1} / 24</p>
                  <div className="flex flex-wrap items-end gap-x-4 gap-y-1 sm:gap-x-5">
                    <h1 lang="he" dir="rtl" className="font-hebrew text-4xl font-semibold leading-none text-primary sm:text-6xl">{current.hebrew}</h1>
                    <p className="font-serif text-xl text-navy-700 sm:text-2xl">{current.translit}</p>
                  </div>
                  <p className="mt-2 text-lg font-semibold sm:mt-3 sm:text-xl">{current.translation}</p>
                </div>
                <span className="hidden rounded-full border border-gold/40 bg-gold/10 px-3 py-1.5 text-xs font-semibold text-amber-900 sm:block">по материалам Толдот</span>
              </div>

              <div className="grid gap-5 xl:grid-cols-[1.05fr_.95fr]">
                <div className="space-y-5">
                  <article className="paper-card p-4 sm:p-6">
                    <p className="section-label">Краткий смысл</p>
                    <p className="mt-3 text-[1.05rem] leading-7 text-navy-900">{current.explanation}</p>
                  </article>

                  <article className="paper-card p-4 sm:p-6">
                    <p className="section-label">О чём подумать</p>
                    <p className="mt-1 text-sm text-muted-foreground">Отметьте мысли, которые помогают вашей каване.</p>
                    <div className="mt-5 space-y-3">
                      {current.kavanot.map((kavana, index) => (
                        <label key={kavana} className="choice-row">
                          <Checkbox checked={entry.kavanot.includes(index)} onCheckedChange={() => update({ kavanot: toggle(entry.kavanot, index) })} />
                          <span>{kavana}</span>
                        </label>
                      ))}
                    </div>
                  </article>
                </div>

                <article className="paper-card p-4 sm:p-6">
                  <p className="section-label">Моя личная конкретизация</p>
                  <p className="mt-1 text-sm leading-6 text-muted-foreground">Выберите только то, что относится к вам, или напишите своими словами.</p>
                  <div className="mt-5 space-y-3">
                    {current.examples.map((example, index) => (
                      <label key={example} className="choice-row items-start">
                        <Checkbox checked={entry.selected.includes(index)} onCheckedChange={() => update({ selected: toggle(entry.selected, index) })} />
                        <span>{example}</span>
                      </label>
                    ))}
                  </div>
                  <label className="mt-5 block">
                    <span className="text-sm font-semibold">Сформулировать своё</span>
                    <Textarea
                      value={entry.custom}
                      onChange={(event) => update({ custom: event.target.value })}
                      className="mt-2 min-h-28 resize-y bg-white text-base sm:min-h-32"
                      placeholder="Я хочу признать, что…"
                    />
                  </label>
                  <p className="mt-3 text-xs leading-5 text-muted-foreground">Не записывайте имена других людей. Если проступок был перед человеком, одной исповеди недостаточно: нужно также исправить вред и попросить прощения.</p>
                </article>
              </div>

              <div className="mt-7 hidden items-center justify-between gap-3 lg:flex">
                <Button variant="outline" disabled={active === 0} onClick={() => goTo(active - 1)}>
                  <ArrowLeft className="size-4" /> Назад
                </Button>
                {active === viduiItems.length - 1 ? (
                  <Button onClick={() => setView('summary')}>
                    Собрать документ <FileText className="size-4" />
                  </Button>
                ) : (
                  <Button onClick={() => goTo(active + 1)}>
                    Следующая часть <ArrowRight className="size-4" />
                  </Button>
                )}
              </div>

              <Sources compact />
            </div>
          </section>

          <div className="print-hidden fixed inset-x-0 bottom-0 z-50 border-t border-navy-100 bg-white/95 px-3 py-2.5 shadow-[0_-10px_30px_rgba(23,42,69,0.08)] backdrop-blur lg:hidden">
            <div className="mx-auto grid max-w-lg grid-cols-[1fr_auto_1fr] items-center gap-2">
              <Button variant="outline" className="min-h-11 justify-self-stretch" disabled={active === 0} onClick={() => goTo(active - 1)}>
                <ArrowLeft className="size-4" /> Назад
              </Button>
              <span className="min-w-12 text-center text-sm font-semibold tabular-nums text-muted-foreground">{active + 1}/24</span>
              {active === viduiItems.length - 1 ? (
                <Button className="min-h-11 justify-self-stretch" onClick={() => setView('summary')}>
                  Итог <FileText className="size-4" />
                </Button>
              ) : (
                <Button className="min-h-11 justify-self-stretch" onClick={() => goTo(active + 1)}>
                  Дальше <ArrowRight className="size-4" />
                </Button>
              )}
            </div>
          </div>
        </div>
      ) : (
        <Summary
          notes={notes}
          filledItems={filledItems}
          message={message}
          onBack={() => setView('work')}
          onEdit={(id) => goTo(viduiItems.findIndex((item) => item.id === id))}
          onCopy={copyDocument}
          onDownload={downloadDocument}
        />
      )}
    </main>
  );
}

function ViduiProgress({ current, className = '' }: { current: number; className?: string }) {
  return (
    <div
      role="progressbar"
      aria-label={`Часть ${current} из 24`}
      aria-valuemin={1}
      aria-valuemax={24}
      aria-valuenow={current}
      className={`h-1.5 overflow-hidden rounded-full bg-navy-50 ${className}`}
    >
      <span className="block h-full rounded-full bg-gold transition-[width]" style={{ width: `${(current / 24) * 100}%` }} />
    </div>
  );
}

function Summary({
  notes,
  filledItems,
  message,
  onBack,
  onEdit,
  onCopy,
  onDownload,
}: {
  notes: Notes;
  filledItems: typeof viduiItems;
  message: string;
  onBack: () => void;
  onEdit: (id: string) => void;
  onCopy: () => void;
  onDownload: () => void;
}) {
  return (
    <section className="summary-page mx-auto max-w-4xl px-3 py-5 sm:px-8 sm:py-12">
      <div className="print-hidden mb-5 grid gap-3 sm:mb-8 sm:flex sm:flex-wrap sm:items-center sm:justify-between sm:gap-4">
        <Button variant="outline" className="min-h-11 w-full sm:w-auto" onClick={onBack}><ArrowLeft className="size-4" /> Вернуться к частям</Button>
        <div className="grid grid-cols-2 gap-2 sm:flex sm:flex-wrap">
          <Button variant="outline" className="min-h-11" onClick={onCopy}><Clipboard className="size-4" /> {message || 'Копировать'}</Button>
          <Button variant="outline" className="min-h-11" onClick={onDownload}><Download className="size-4" /> Скачать</Button>
          <Button className="col-span-2 min-h-11" onClick={() => window.print()}><Printer className="size-4" /> Печать / PDF</Button>
        </div>
      </div>

      <div className="document-sheet rounded-2xl border border-navy-100 bg-white px-4 py-7 shadow-[0_18px_50px_rgba(23,42,69,0.08)] sm:px-12 sm:py-12">
        <div className="border-b border-navy-100 pb-6 text-center sm:pb-8">
          <p className="text-xs font-bold uppercase tracking-[0.18em] text-teal-700">Личная подготовка к Йом-Кипуру</p>
          <h1 className="mt-3 font-serif text-3xl font-semibold text-primary sm:text-4xl">Мой видуй</h1>
          <p className="mx-auto mt-3 max-w-xl text-sm leading-6 text-muted-foreground">Здесь собраны только выбранные вами направления мысли и личные формулировки. Произносите лишь то, что правдиво и относится к вам.</p>
        </div>

        {filledItems.length ? (
          <div className="mt-8 space-y-7">
            {filledItems.map((item) => {
              const entry = notes[item.id] ?? emptyEntry();
              const number = viduiItems.findIndex((candidate) => candidate.id === item.id) + 1;
              return (
                <article key={item.id} className="summary-entry break-inside-avoid border-b border-navy-100 pb-7 last:border-0">
                  <div className="flex items-start justify-between gap-4">
                    <div>
                      <p className="text-xs font-bold text-teal-700">{number} / 24</p>
                      <div className="mt-2 flex flex-wrap items-baseline gap-x-3 gap-y-1">
                        <h2 lang="he" dir="rtl" className="font-hebrew text-3xl font-semibold text-primary">{item.hebrew}</h2>
                        <p className="font-serif text-xl font-semibold">{item.translit}</p>
                      </div>
                      <p className="mt-1 text-sm font-semibold text-navy-700">{item.translation}</p>
                    </div>
                    <button type="button" onClick={() => onEdit(item.id)} className="print-hidden text-sm font-semibold text-teal-700 hover:underline">Изменить</button>
                  </div>

                  {entry.kavanot.length > 0 && (
                    <div className="mt-5">
                      <p className="section-label">О чём я хочу думать</p>
                      <ul className="mt-2 space-y-1.5 text-[0.96rem] leading-6">
                        {entry.kavanot.map((index) => <li key={index} className="flex gap-2"><span className="text-gold">•</span><span>{item.kavanot[index]}</span></li>)}
                      </ul>
                    </div>
                  )}

                  {(entry.selected.length > 0 || entry.custom.trim()) && (
                    <div className="mt-5 rounded-xl bg-navy-50/60 px-4 py-4">
                      <p className="section-label">Моя личная конкретизация</p>
                      <ul className="mt-2 space-y-2 text-[0.96rem] leading-6">
                        {entry.selected.map((index) => <li key={index} className="flex gap-2"><span className="text-gold">•</span><span>{item.examples[index]}</span></li>)}
                        {entry.custom.trim() && <li className="flex gap-2"><span className="text-gold">•</span><span>{entry.custom.trim()}</span></li>}
                      </ul>
                    </div>
                  )}
                </article>
              );
            })}
          </div>
        ) : (
          <div className="py-16 text-center">
            <FileText className="mx-auto size-10 text-navy-200" aria-hidden="true" />
            <h2 className="mt-4 font-serif text-2xl font-semibold">Документ пока пуст</h2>
            <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-muted-foreground">Вернитесь к частям видуя, отметьте подходящие мысли или добавьте собственную формулировку.</p>
            <Button className="print-hidden mt-5" onClick={onBack}>Начать заполнять</Button>
          </div>
        )}

        <Sources />
      </div>
    </section>
  );
}

function Sources({ compact = false }: { compact?: boolean }) {
  return (
    <footer className={`${compact ? 'mt-10 rounded-xl border border-navy-100 bg-white px-5 py-4' : 'mt-10 border-t border-navy-100 pt-7'} text-xs leading-5 text-muted-foreground`}>
      <p><strong className="text-navy-700">Об источниках.</strong> Перевод заголовков и направления для размышления сокращены и адаптированы по материалам Толдот. Личные примеры составлены как подсказки и не являются готовой исповедью за человека.</p>
      <p className="mt-2 print-links">
        <a className="text-teal-700 underline underline-offset-2" href="https://toldot.com/pdf/ekrupnik-mahzor_na_Iom_Kipur_Ashkenaz.pdf" target="_blank" rel="noreferrer">Махзор с переводом и объяснениями</a>
        <span aria-hidden="true"> · </span>
        <a className="text-teal-700 underline underline-offset-2" href="https://toldot.com/audio/lessons/lessons_37205.html" target="_blank" rel="noreferrer">Разбор молитвы «Видуй»</a>
        <span aria-hidden="true"> · </span>
        <a className="text-teal-700 underline underline-offset-2" href="https://toldot.com/articles/articles_2216.html" target="_blank" rel="noreferrer">Законы и традиции Йом-Кипура</a>
      </p>
      <p className="mt-2">Это помощник для самоанализа, а не замена махзору или решению раввина. В вопросах практического закона обратитесь к своему раввину.</p>
    </footer>
  );
}
