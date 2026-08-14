import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import * as XLSX from 'npm:xlsx@0.18.5'

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

const supabaseUrl = Deno.env.get('SUPABASE_URL')!
const supabaseAnonKey = Deno.env.get('SUPABASE_PUBLISHABLE_KEY') || Deno.env.get('SUPABASE_ANON_KEY')!
const geminiKey = Deno.env.get('GEMINI_API_KEY')!
const geminiModel = Deno.env.get('GEMINI_MODEL') || 'gemini-3.6-flash'

const DAY_ORDER = ['Понедельник','Вторник','Среда','Четверг','Пятница','Суббота','Воскресенье']
const LATIN_TO_CYR = new Map([
  ['A','А'],['B','В'],['C','С'],['E','Е'],['H','Н'],['K','К'],['M','М'],['O','О'],['P','Р'],['T','Т'],['X','Х'],['Y','У'],
])

function normalizeClassName(value: unknown): string | null {
  let s = String(value ?? '').trim().toUpperCase()
  s = s.replace(/\bКЛАСС\b/giu, '').replace(/[\s._\-–—:№]+/g, '')
  for (const [latin, cyr] of LATIN_TO_CYR) s = s.replaceAll(latin, cyr)
  const match = s.match(/^(5|6|7|8|9|10|11)([АБВГ])$/u)
  return match ? `${match[1]}${match[2]}` : null
}

function normalizeDay(value: unknown): string {
  const raw = String(value ?? '').trim().toLowerCase().replace(/ё/g, 'е')
  const aliases: Record<string, string> = {
    'пн':'Понедельник','пон':'Понедельник','понедельник':'Понедельник',
    'вт':'Вторник','вторник':'Вторник',
    'ср':'Среда','среда':'Среда',
    'чт':'Четверг','четверг':'Четверг',
    'пт':'Пятница','пятница':'Пятница',
    'сб':'Суббота','суббота':'Суббота',
    'вс':'Воскресенье','воскресенье':'Воскресенье',
  }
  return aliases[raw] || String(value ?? '').trim() || 'Понедельник'
}

function normalizeLesson(raw: any, fallbackIndex: number): any | null {
  if (!raw || typeof raw !== 'object') return null
  const subject = String(raw.subject ?? raw.name ?? raw.title ?? '').trim()
  if (!subject) return null
  const lessonRaw = Number(raw.lesson ?? raw.number ?? fallbackIndex + 1)
  const lesson = Number.isFinite(lessonRaw) && lessonRaw > 0 ? Math.floor(lessonRaw) : fallbackIndex + 1
  return {
    day: normalizeDay(raw.day),
    lesson,
    subject,
    time: String(raw.time ?? '').trim(),
    room: String(raw.room ?? raw.cabinet ?? raw.classroom ?? '').trim(),
  }
}

function normalizeLessons(raw: unknown): any[] {
  if (!Array.isArray(raw)) return []
  const result: any[] = []
  raw.forEach((item, i) => {
    const lesson = normalizeLesson(item, i)
    if (lesson) result.push(lesson)
  })
  result.sort((a,b) => {
    const d = DAY_ORDER.indexOf(a.day) - DAY_ORDER.indexOf(b.day)
    return d || a.lesson - b.lesson || a.subject.localeCompare(b.subject, 'ru')
  })
  const seen = new Set<string>()
  return result.filter(x => {
    const key = `${x.day}|${x.lesson}|${x.subject}|${x.time}|${x.room}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function normalizeSchedules(input: any) {
  const out: Record<string, any[]> = {}
  const ignoredClasses: string[] = []
  const source = Array.isArray(input?.schedules)
    ? input.schedules
    : input?.schedules && typeof input.schedules === 'object'
      ? Object.entries(input.schedules).map(([class_name, lessons]) => ({ class_name, lessons }))
      : Array.isArray(input)
        ? input
        : []

  for (const item of source) {
    const rawClass = item?.class_name ?? item?.className ?? item?.class ?? item?.name
    const className = normalizeClassName(rawClass)
    if (!className) {
      if (String(rawClass ?? '').trim()) ignoredClasses.push(String(rawClass).trim())
      continue
    }
    const lessons = normalizeLessons(item?.lessons)
    if (!out[className]) out[className] = []
    out[className].push(...lessons)
  }

  for (const className of Object.keys(out)) out[className] = normalizeLessons(out[className])
  return { schedules: out, ignoredClasses: [...new Set(ignoredClasses)] }
}


const responseSchema = {
  type: 'object',
  properties: {
    schedules: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          class_name: { type: 'string' },
          lessons: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                day: { type: 'string' },
                lesson: { type: 'integer' },
                subject: { type: 'string' },
                time: { type: 'string' },
                room: { type: 'string' },
              },
              required: ['day','lesson','subject','time','room'],
            },
          },
        },
        required: ['class_name','lessons'],
      },
    },
  },
  required: ['schedules'],
}

const SUPPORTED_EXTENSIONS = new Set(['xls','xlsx','ods','csv'])

function base64ToBytes(value: string): Uint8Array {
  const clean = value.includes(',') ? value.split(',').pop()! : value
  const binary = atob(clean)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value.toLocaleString('ru-RU')
  return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim()
}

function scoreScheduleSheet(rows: string[][]): number {
  const text = rows.map(row => row.join(' | ')).join('\n')
  const classHits = (text.match(/\b(?:5|6|7|8|9|10|11)\s*[АБВГAB]\b/giu) || []).length
  const dayHits = (text.match(/\b(?:понедельник|вторник|среда|четверг|пятница|суббота)\b/giu) || []).length
  const timeHits = (text.match(/\b\d{1,2}[:.]\d{2}\s*[-–—]\s*\d{1,2}[:.]\d{2}\b/g) || []).length
  const lessonHits = rows.reduce((n, row) => n + row.slice(0, 4).filter(x => /^\d{1,2}$/.test(x.trim())).length, 0)
  const nonEmpty = rows.reduce((n, row) => n + row.filter(Boolean).length, 0)
  return classHits * 20 + dayHits * 8 + timeHits * 6 + lessonHits * 2 + Math.min(nonEmpty, 500) / 100
}

function sheetToCsvWithMergedCells(sheet: XLSX.WorkSheet): { rows: string[][], csv: string } {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
    blankrows: true,
  }) as unknown[][]
  const grid = rows.map(row => Array.isArray(row) ? row.map(normalizeCell) : [])

  for (const merge of sheet['!merges'] || []) {
    const top = grid[merge.s.r]?.[merge.s.c] ?? ''
    if (!top) continue
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      if (!grid[r]) grid[r] = []
      for (let c = merge.s.c; c <= merge.e.c; c++) grid[r][c] = top
    }
  }

  const width = grid.reduce((m, row) => Math.max(m, row.length), 0)
  for (const row of grid) while (row.length < width) row.push('')

  const csv = grid.map(row => XLSX.utils.sheet_to_csv(
    XLSX.utils.aoa_to_sheet([row]),
    { FS: ',', RS: '\n', blankrows: true }
  ).trimEnd()).join('\n')
  return { rows: grid, csv }
}

function workbookToCleanCsv(bytes: Uint8Array, extension: string): string {
  let workbook: XLSX.WorkBook
  try {
    workbook = XLSX.read(bytes, {
      type: 'array',
      cellDates: true,
      raw: false,
      dense: false,
      WTF: false,
    })
  } catch (err) {
    throw new Error(`Не удалось прочитать таблицу .${extension}: ${err instanceof Error ? err.message : String(err)}`)
  }

  const candidates = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name]
    const converted = sheetToCsvWithMergedCells(sheet)
    return { name, ...converted, score: scoreScheduleSheet(converted.rows) }
  }).sort((a,b) => b.score - a.score)

  const best = candidates[0]
  if (!best || !best.csv.trim()) throw new Error('В таблице не найдено ни одного непустого листа.')

  // Обычно расписание находится на одном листе. Передаём только наиболее похожий
  // на расписание лист, чтобы не засорять контекст мета-информацией из других листов.
  return `ЛИСТ: ${best.name}\n\n${best.csv}`
}

const SCHEDULE_PROMPT = `Ты — строгий парсер школьного расписания. На входе НЕ фотография и НЕ PDF, а очищенный CSV-текст электронной таблицы. Твоя задача — восстановить только основную сетку расписания для классов 5А–11Б.

КЛЮЧЕВОЕ ПРАВИЛО: В таблице могут быть служебные блоки сверху и снизу: дежурство, ВПР, замены, объявления, примечания, подписи, даты, телефоны и прочая мета-информация. ПОЛНОСТЬЮ ИГНОРИРУЙ их. Ищи ТОЛЬКО НИЖНЮЮ ОСНОВНУЮ СЕТКУ РАСПИСАНИЯ. Она может начинаться не с первой строки — например, с 17-й, 20-й или другой строки. Начало сетки определяется по сочетанию заголовков классов, дней/номеров уроков, времени и предметов.

1. КЛАССЫ
- Ищи и обрабатывай классы от 5А до 11Б включительно. Если в таблице есть 5В, 10Г и т.п., НЕ добавляй их: в этом задании нужны только 5А–11Б.
- Нормализуй 8а/8А/8 а/8-А/8A в 8А; 8Б сохраняй как 8Б.
- Латинская B в заголовке может означать кириллическую В, но не превращай явную Б в В.
- Не придумывай класс, которого нет в таблице.

2. ГРАНИЦЫ ОСНОВНОЙ СЕТКИ
- Сначала найди строку/строки заголовков классов.
- Затем найди основную сетку под этими заголовками и читай её ДО КОНЦА ВНИЗ.
- Верхние служебные блоки, даже если там встречаются названия классов, не считать расписанием.
- Нижние примечания/дежурство/ВПР после основной сетки не считать расписанием.
- Особое внимание строкам 6, 7, 8 и 9 уроков и вообще последним строкам основной сетки: они не должны потеряться.

3. ЧТЕНИЕ СТРОГО ПО ВЕРТИКАЛИ
- Для КАЖДОГО найденного класса читай его столбец строго сверху вниз, строка за строкой.
- Не смешивай соседние столбцы классов.
- Если расписание разделено на дни, переноси день вниз на все строки соответствующего блока до следующего дня.
- Номер урока и время бери из ЛЕВОЙ части той же строки. Сохраняй фактический номер урока и фактическое время, не заменяй их индексом массива.
- Если номер урока или время видны, обязательно сохрани их.
- Если в ячейке стоит ровно прочерк "-", запиши subject как "-". Не превращай его в пустую строку и не удаляй такой урок.
- Пустая ячейка означает отсутствие урока, КРОМЕ СЛУЧАЯ ОБЪЕДИНЁННОЙ ЯЧЕЙКИ.

4. ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ И ПУСТОТЫ
- Excel/ODS может представить объединённую ячейку как значение только в верхней левой клетке, а остальные клетки той же объединённой области будут пустыми. При конвертации часть таких значений уже продублирована по диапазону объединения, но ты всё равно обязан учитывать исходную логику таблицы.
- Если видно, что один урок/предмет относится сразу к нескольким классам или группам, полностью продублируй этот текст в каждой соответствующей колонке класса.
- Если значение стоит в верхней клетке вертикально объединённого блока и ниже идут пустые клетки в том же логическом блоке, наследуй значение вниз ТОЛЬКО если структура таблицы явно показывает объединение/общий урок. Не заполняй обычные случайные пустоты догадками.
- Никогда не теряй урок из-за того, что его ячейка после конвертации CSV стала пустой вследствие объединения.

5. ПРЕДМЕТ И КАБИНЕТ
- subject должен содержать название предмета/текст урока так, как он указан в ячейке, включая сокращения.
- room должен содержать кабинет, если он указан отдельно или явно читается в той же ячейке. Не выдумывай кабинет.
- Если предмет и кабинет записаны вместе, раздели их разумно: название предмета в subject, номер/обозначение кабинета в room.
- Если предмет реально отсутствует, не выдумывай его.

6. JSON
- Верни ТОЛЬКО JSON по заданной схеме.
- В JSON не должно быть OCR-комментариев, мета-информации, названий файлов, confidence, source, notes и других полей.
- Каждый class_name содержит только один канонический класс.
- lessons содержит только day, lesson, subject, time, room. День нужен приложению для правильного показа расписания по дням.
- Не добавляй уроки из служебных блоков.
- Лучше оставить реально пустое поле пустым, чем придумать значение.

ПЕРЕД ОТВЕТОМ СДЕЛАЙ КОНТРОЛЬ:
а) проверь каждый класс 5А–11Б отдельно;
б) проверь основную сетку от её первой строки до последней;
в) отдельно перепроверь нижнюю часть таблицы;
г) убедись, что прочерки "-" сохранены;
д) убедись, что объединённые уроки продублированы для всех нужных классов;
е) убедись, что старые/служебные блоки не попали в JSON.`

async function callGemini(textPrompt: string) {
  const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(geminiModel)}:generateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': geminiKey },
    body: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: textPrompt }] }],
      generationConfig: {
        responseMimeType: 'application/json',
        responseSchema,
        temperature: 0,
      },
    }),
  })
  const data = await r.json()
  if (!r.ok) throw new Error(data?.error?.message || `Gemini HTTP ${r.status}`)
  let raw = ''
  for (const p of data?.candidates?.[0]?.content?.parts || []) if (p.text) raw += p.text
  raw = raw.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim()
  if (!raw) throw new Error('Gemini не вернул JSON.')
  try { return JSON.parse(raw) } catch { throw new Error('Gemini вернул невалидный JSON.') }
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: cors })
  if (req.method !== 'POST') return json({ error: 'Method not allowed' }, 405)

  try {
    if (!geminiKey) return json({ error: 'На сервере не задан GEMINI_API_KEY.' }, 500)

    const auth = req.headers.get('Authorization') || ''
    const supabase = createClient(supabaseUrl, supabaseAnonKey, {
      global: { headers: { Authorization: auth } },
    })
    const { data: { user } } = await supabase.auth.getUser()
    if (!user) return json({ error: 'Необходим вход в аккаунт.' }, 401)

    const { data: profile, error: profileError } = await supabase
      .from('profiles').select('role').eq('id', user.id).single()
    if (profileError || profile?.role !== 'admin') return json({ error: 'Доступ только для администратора.' }, 403)

    const body = await req.json()
    const fileBase64 = String(body.fileBase64 || '')
    const fileName = String(body.fileName || 'schedule.xlsx')
    const mimeType = String(body.mimeType || 'application/octet-stream').split(';')[0]
    const extension = (fileName.split('.').pop() || '').toLowerCase()

    if (!fileBase64) return json({ error: 'Файл не передан.' }, 400)
    if (!SUPPORTED_EXTENSIONS.has(extension)) return json({ error: 'Поддерживаются только .xls, .xlsx, .ods и .csv.' }, 400)
    if (fileBase64.length > 18_000_000) return json({ error: 'Файл слишком большой. Максимум 12 МБ.' }, 413)

    let csvText: string
    try {
      csvText = workbookToCleanCsv(base64ToBytes(fileBase64), extension)
    } catch (err) {
      return json({ error: err instanceof Error ? err.message : String(err) }, 422)
    }

    if (csvText.length > 180_000) {
      return json({ error: 'После очистки таблица слишком большая для обработки одной операцией. Удали лишние листы/служебные данные и загрузи только лист с расписанием.' }, 413)
    }

    const prompt = `${SCHEDULE_PROMPT}

ОЧИЩЕННЫЙ CSV:
${csvText}`

    let firstPass: any
    try {
      firstPass = await callGemini(prompt)
    } catch (err) {
      return json({ error: err instanceof Error ? err.message : String(err) }, 502)
    }

    const normalized = normalizeSchedules(firstPass)
    const entries = Object.entries(normalized.schedules)
    if (!entries.length) {
      return json({
        error: 'Не удалось распознать ни одного поддерживаемого класса 5А–11Б.',
        ignoredClasses: normalized.ignoredClasses,
      }, 422)
    }

    // Важное изменение: сохранение выполняется одной атомарной SQL-операцией.
    // Она СНАЧАЛА удаляет ВСЕ старые строки schedules, затем вставляет только
    // результат текущего файла. Поэтому старые дни, классы и алиасы никогда
    // не смешиваются с новым расписанием.
    const { error: replaceError } = await supabase.rpc('replace_schedules', {
      p_schedules: normalized.schedules,
    })
    if (replaceError) {
      return json({ error: `Не удалось заменить расписание в базе: ${replaceError.message}` }, 500)
    }

    return json({
      schedules: normalized.schedules,
      ignoredClasses: normalized.ignoredClasses,
      model: geminiModel,
    })
  } catch (err) {
    return json({ error: err instanceof Error ? err.message : String(err) }, 500)
  }
})

function json(data: unknown, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...cors, 'Content-Type': 'application/json' },
  })
}
