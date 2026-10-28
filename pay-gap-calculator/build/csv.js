/* ============================================================================
   Part C — CSV import and validation.
   Public interface: CSV.parse(text) → ParseResult.

   Plain ES2020, a global object, no modules and no external dependencies.
   The main function is pure: a string in, an object out.
   The only part that touches the DOM is the optional CSV.readFile helper
   (FileReader).
   ========================================================================== */

var CSV = (function () {
  'use strict';

  // --- Constants -----------------------------------------------------------

  // Required columns.
  var REQUIRED = [
    'category',
    'gender',
    'base_salary',
    'variable_pay',
    'tenure_years',
    'grade'
  ];

  // Optional columns with silent defaults.
  var DEFAULTS = { fte: 1.0, months_worked: 12 };

  // Share of broken rows above which the file is rejected outright.
  var MAX_BAD_ROW_SHARE = 0.20;

  // Synonyms for headers found in typical HR exports.
  // The key is an already-normalised header (lower case, underscores); the
  // value is the canonical column name.
  var SYNONYMS = {
    // base_salary
    salary: 'base_salary',
    annual_salary: 'base_salary',
    base_pay: 'base_salary',
    basic_salary: 'base_salary',
    base: 'base_salary',
    fixed_pay: 'base_salary',
    gross_salary: 'base_salary',
    // variable_pay
    bonus: 'variable_pay',
    variable: 'variable_pay',
    variable_compensation: 'variable_pay',
    variable_salary: 'variable_pay',
    bonus_pay: 'variable_pay',
    incentive: 'variable_pay',
    // gender
    sex: 'gender',
    // category
    job_category: 'category',
    job_family: 'category',
    category_of_work: 'category',
    work_category: 'category',
    job_group: 'category',
    // tenure_years
    tenure: 'tenure_years',
    years_of_service: 'tenure_years',
    seniority: 'tenure_years',
    service_years: 'tenure_years',
    // grade
    level: 'grade',
    job_level: 'grade',
    band: 'grade',
    pay_grade: 'grade',
    // fte
    full_time_equivalent: 'fte',
    fte_ratio: 'fte',
    workload: 'fte',
    part_time_factor: 'fte',
    // months_worked
    months: 'months_worked',
    months_in_year: 'months_worked',
    worked_months: 'months_worked',
    // id
    employee_id: 'id',
    emp_id: 'id',
    employee_number: 'id',
    staff_id: 'id',
    personnel_number: 'id'
  };

  // Recognised gender values → 'F' | 'M'.
  var GENDER_MAP = {
    f: 'F', female: 'F', woman: 'F', women: 'F', w: 'F',
    'ж': 'F', 'жен': 'F', 'женщина': 'F', 'женский': 'F',
    m: 'M', male: 'M', man: 'M', men: 'M',
    'м': 'M', 'муж': 'M', 'мужчина': 'M', 'мужской': 'M'
  };

  // --- Low-level CSV parser ------------------------------------------------

  /**
   * Strips the BOM (Excel writes one almost always) and normalises line breaks.
   * CRLF and a lone CR both become LF, so the rest of the code deals with one
   * form only.
   */
  function stripBomAndNormalise(text) {
    var s = String(text == null ? '' : text);
    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
    return s.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  }

  /**
   * Detects the delimiter: comma or semicolon.
   * Occurrences OUTSIDE quotes are counted over the first non-empty lines, and
   * the more frequent one wins. On a tie the semicolon is chosen: European
   * exports are the more common source of the situations where a tie is
   * possible at all (a decimal comma inside numbers).
   */
  function detectDelimiter(text) {
    var counts = { ',': 0, ';': 0 };
    var inQuotes = false;
    var lines = 0;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (ch === '"') {
        // A doubled quote inside a field is not a change of mode.
        if (inQuotes && text[i + 1] === '"') { i++; continue; }
        inQuotes = !inQuotes;
        continue;
      }
      if (!inQuotes) {
        if (ch === ',' || ch === ';') counts[ch]++;
        else if (ch === '\n') {
          lines++;
          if (lines >= 5) break; // five lines are enough
        }
      }
    }
    if (counts[';'] > counts[',']) return ';';
    if (counts[','] > counts[';']) return ',';
    return ';';
  }

  /**
   * Parses the text into a matrix of rows. Supports quotes, doubled quotes
   * inside a field ("" → "), and line breaks inside quotes.
   * Empty rows are skipped, trailing ones included.
   */
  function splitRows(text, delim) {
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    var fieldWasQuoted = false;

    function pushField() {
      row.push(fieldWasQuoted ? field : field.trim());
      field = '';
      fieldWasQuoted = false;
    }
    function pushRow() {
      pushField();
      // A row counts as empty when every one of its fields is empty.
      var empty = row.every(function (c) { return c === ''; });
      if (!empty) rows.push(row);
      row = [];
    }

    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else {
          field += ch;
        }
      } else {
        if (ch === '"') { inQuotes = true; fieldWasQuoted = true; }
        else if (ch === delim) pushField();
        else if (ch === '\n') pushRow();
        else field += ch;
      }
    }
    // The tail of the last row, when it has no closing line break.
    if (field !== '' || row.length > 0) pushRow();
    return rows;
  }

  // --- Header normalisation ------------------------------------------------

  /**
   * "Base Salary", "base-salary", "BASE_SALARY" → base_salary.
   * Lower case, trimmed, spaces/hyphens/dots → underscore, repeats collapsed
   * and underscores trimmed from both ends.
   */
  function normaliseHeader(h) {
    return String(h == null ? '' : h)
      .replace(/^﻿/, '')
      .trim()
      .toLowerCase()
      .replace(/[\s\-.]+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '');
  }

  /** Canonical column name: normalisation plus the synonym dictionary. */
  function canonicalHeader(h) {
    var n = normaliseHeader(h);
    return Object.prototype.hasOwnProperty.call(SYNONYMS, n) ? SYNONYMS[n] : n;
  }

  // --- Value parsing --------------------------------------------------------

  /**
   * Numbers in both European and plain format.
   * Accepts "52.000,50", "52 000,50", "52000.50", "52 000,50 ₽", "1 234 руб."
   * Currency symbols and spaces (non-breaking ones included) are discarded as
   * thousands separators.
   * The format is decided by the position of the LAST separator:
   *   - if 1-2 digits follow it and then the string ends, it is decimal;
   *   - otherwise both kinds of separator are thousands separators.
   * Returns a number, or null when it cannot be recognised.
   */
  function parseNumber(raw) {
    if (raw == null) return null;
    var s = String(raw).trim();
    if (s === '') return null;

    // Strip currency, spaces (ordinary, non-breaking, narrow) and apostrophe separators.
    s = s.replace(/руб\.?|р\./gi, '');
    s = s.replace(/[€$£¥₽\s   ']/g, '');
    if (s === '') return null;

    // Sign.
    var negative = false;
    if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); } // (1 234) — accounting-style minus
    if (s[0] === '+') s = s.slice(1);
    else if (s[0] === '-') { negative = true; s = s.slice(1); }

    if (!/^[0-9.,]*$/.test(s) || !/[0-9]/.test(s)) return null;

    var lastComma = s.lastIndexOf(',');
    var lastDot = s.lastIndexOf('.');
    var lastSep = Math.max(lastComma, lastDot);

    var intPart, fracPart = '';
    if (lastSep === -1) {
      intPart = s;
    } else {
      var sepChar = s[lastSep];
      var tail = s.slice(lastSep + 1);
      // Decimal separator: 1-2 digits in the tail, and it is the only one of its kind.
      var sameCharCount = s.split(sepChar).length - 1;
      var isDecimal = /^[0-9]{1,2}$/.test(tail) && sameCharCount === 1;
      // "1.234" is ambiguous: a dot with three digits means thousands, hence 1-2 digits.
      if (isDecimal) {
        intPart = s.slice(0, lastSep);
        fracPart = tail;
      } else {
        intPart = s;
      }
    }
    intPart = intPart.replace(/[.,]/g, '');
    if (intPart === '') intPart = '0';
    if (!/^[0-9]+$/.test(intPart) || (fracPart !== '' && !/^[0-9]+$/.test(fracPart))) return null;

    var value = Number(intPart + (fracPart ? '.' + fracPart : ''));
    if (!isFinite(value)) return null;
    return negative ? -value : value;
  }

  /** Gender: F/M/female/male/woman/man in any case → 'F' | 'M' | null. */
  function parseGender(raw) {
    if (raw == null) return null;
    var s = String(raw).trim().toLowerCase().replace(/[.\s]+$/, '');
    if (s === '') return null;
    return Object.prototype.hasOwnProperty.call(GENDER_MAP, s) ? GENDER_MAP[s] : null;
  }

  /** Sequential ids of the form E001, E002, … when the file has no id column. */
  function makeId(n) {
    var s = String(n);
    while (s.length < 3) s = '0' + s;
    return 'E' + s;
  }

  /** Plural agreement: 1 row / 2 rows. */
  // Russian plural of "строка": 1 строка, 2 строки, 5 строк.
  function rowWord(n) {
    var m10 = n % 10, m100 = n % 100;
    if (m10 === 1 && m100 !== 11) return 'строка';
    if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return 'строки';
    return 'строк';
  }

  // --- Public parse ---------------------------------------------------------

  /**
   * CSV.parse(text) → { ok, employees, errors, warnings }
   * A single malformed row does not sink the file: it goes into warnings and is
   * skipped. If more than 20% of rows are broken, the file itself is at fault.
   */
  function parse(text) {
    var errors = [];
    var warnings = [];
    var fail = function () {
      return { ok: false, employees: [], errors: errors, warnings: warnings };
    };

    var normalised = stripBomAndNormalise(text);
    if (normalised.trim() === '') {
      errors.push('Файл пустой. Выгрузите CSV со строкой заголовков и одной строкой на сотрудника.');
      return fail();
    }

    var delim = detectDelimiter(normalised);
    var rows = splitRows(normalised, delim);

    if (rows.length === 0) {
      errors.push('В файле нет читаемых строк. Выгрузите CSV со строкой заголовков и одной строкой на сотрудника.');
      return fail();
    }

    // --- Headers ---
    var rawHeaders = rows[0];
    var headers = rawHeaders.map(canonicalHeader);
    var present = {};
    headers.forEach(function (h, i) {
      if (h !== '' && !Object.prototype.hasOwnProperty.call(present, h)) present[h] = i;
    });

    var foundList = headers.filter(function (h) { return h !== ''; }).join(', ');
    var missing = REQUIRED.filter(function (c) {
      return !Object.prototype.hasOwnProperty.call(present, c);
    });

    if (missing.length > 0) {
      // The message must name what is missing and show what was found, so the
      // user can see straight away what to fix in their export.
      var what = missing.length === 1
        ? 'Нет обязательного столбца "' + missing[0] + '".'
        : 'Нет обязательных столбцов: ' + missing.map(function (m) { return '"' + m + '"'; }).join(', ') + '.';
      errors.push(
        what +
        ' В файле есть столбцы: ' + (foundList || '(нет)') + '.' +
        ' Переименуйте подходящий столбец в выгрузке или добавьте его.' +
        ' Обязательные столбцы: ' + REQUIRED.join(', ') + '.' +
        ' Необязательные: fte (по умолчанию 1,0), months_worked (по умолчанию 12), id (создаётся, если нет).'
      );
      return fail();
    }

    if (rows.length < 2) {
      errors.push('В файле есть заголовки, но нет строк с данными. Добавьте под заголовками по строке на сотрудника.');
      return fail();
    }

    // --- Rows ---
    var employees = [];
    var badRows = [];              // messages about skipped rows
    var dataRowCount = rows.length - 1;
    var defaultsUsed = { fte: 0, months_worked: 0 };
    var idSeen = {};
    var autoIdCounter = 0;
    var hasIdColumn = Object.prototype.hasOwnProperty.call(present, 'id');
    var genderSeen = { F: 0, M: 0 };
    var duplicateIdCount = 0;

    // Cell value by canonical column name.
    function cell(row, name) {
      var idx = present[name];
      if (idx === undefined) return '';
      var v = row[idx];
      return v === undefined || v === null ? '' : String(v).trim();
    }

    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      var lineNo = r + 1; // the line number in the file as the user sees it
      var problem = null;

      // Gender.
      var genderRaw = cell(row, 'gender');
      var gender = parseGender(genderRaw);
      if (gender === null) {
        problem = genderRaw === ''
          ? 'gender пустой'
          : 'значение gender "' + genderRaw + '" не распознано (используйте Ж/М или F/M)';
      }

      // Category.
      var category = cell(row, 'category');
      if (problem === null && category === '') problem = 'category пустой';

      // Numeric fields.
      var base_salary = null, variable_pay = null, tenure_years = null, grade = null;
      var fte = DEFAULTS.fte, months_worked = DEFAULTS.months_worked;

      if (problem === null) {
        var baseRaw = cell(row, 'base_salary');
        base_salary = parseNumber(baseRaw);
        if (base_salary === null) {
          problem = baseRaw === ''
            ? 'base_salary пустой'
            : 'base_salary "' + baseRaw + '" — не число';
        } else if (base_salary < 0) {
          problem = 'base_salary отрицательный (' + baseRaw + ')';
        }
      }

      if (problem === null) {
        // An empty variable_pay cell means 0. That is normal, not a warning.
        var varRaw = cell(row, 'variable_pay');
        if (varRaw === '') variable_pay = 0;
        else {
          variable_pay = parseNumber(varRaw);
          if (variable_pay === null) problem = 'variable_pay "' + varRaw + '" — не число';
          else if (variable_pay < 0) problem = 'variable_pay отрицательный (' + varRaw + ')';
        }
      }

      if (problem === null) {
        var tenRaw = cell(row, 'tenure_years');
        tenure_years = parseNumber(tenRaw);
        if (tenure_years === null) {
          problem = tenRaw === ''
            ? 'tenure_years пустой'
            : 'tenure_years "' + tenRaw + '" — не число';
        } else if (tenure_years < 0) {
          problem = 'tenure_years отрицательный (' + tenRaw + ')';
        }
      }

      if (problem === null) {
        var gradeRaw = cell(row, 'grade');
        grade = parseNumber(gradeRaw);
        if (grade === null) {
          problem = gradeRaw === ''
            ? 'grade пустой'
            : 'grade "' + gradeRaw + '" — не число';
        } else {
          grade = Math.round(grade); // grade is an ordinal integer level
        }
      }

      // fte: empty → default 1.0 (silently, but summarised in warnings).
      if (problem === null) {
        var fteRaw = cell(row, 'fte');
        if (fteRaw === '') { fte = DEFAULTS.fte; defaultsUsed.fte++; }
        else {
          fte = parseNumber(fteRaw);
          if (fte === null) problem = 'fte "' + fteRaw + '" — не число';
          else if (!(fte > 0 && fte <= 1)) problem = 'fte вне допустимого диапазона 0 < fte <= 1 (' + fteRaw + ')';
        }
      }

      // months_worked: empty → default 12.
      if (problem === null) {
        var mwRaw = cell(row, 'months_worked');
        if (mwRaw === '') { months_worked = DEFAULTS.months_worked; defaultsUsed.months_worked++; }
        else {
          months_worked = parseNumber(mwRaw);
          if (months_worked === null) problem = 'months_worked "' + mwRaw + '" — не число';
          else if (!(months_worked > 0 && months_worked <= 12)) {
            problem = 'months_worked вне допустимого диапазона 0 < months_worked <= 12 (' + mwRaw + ')';
          }
        }
      }

      if (problem !== null) {
        badRows.push('Строка ' + lineNo + ' пропущена: ' + problem + '.');
        continue;
      }

      // id: from the column, or generated.
      var id = hasIdColumn ? cell(row, 'id') : '';
      if (id === '') {
        do { autoIdCounter++; id = makeId(autoIdCounter); } while (idSeen[id]);
      }
      if (idSeen[id]) {
        // A duplicate id is no reason to lose the row: make it unique and count it.
        duplicateIdCount++;
        var suffix = 2;
        while (idSeen[id + '-' + suffix]) suffix++;
        id = id + '-' + suffix;
      }
      idSeen[id] = true;

      genderSeen[gender]++;

      employees.push({
        id: id,
        category: category,
        gender: gender,
        base_salary: base_salary,
        variable_pay: variable_pay,
        tenure_years: tenure_years,
        grade: grade,
        fte: fte,
        months_worked: months_worked
      });
    }

    // --- Summaries in warnings ---
    if (defaultsUsed.fte > 0) {
      warnings.push('Без fte: ' + defaultsUsed.fte + ' ' + rowWord(defaultsUsed.fte) + ', принято 1,0.');
    }
    if (defaultsUsed.months_worked > 0) {
      warnings.push('Без months_worked: ' + defaultsUsed.months_worked + ' ' + rowWord(defaultsUsed.months_worked) +
        ', принято 12.');
    }
    if (duplicateIdCount > 0) {
      warnings.push('Повторяющийся id: ' + duplicateIdCount + ' ' + rowWord(duplicateIdCount) +
        '; к id добавлен суффикс, чтобы все были уникальны.');
    }

    // --- Robustness: share of broken rows ---
    var badShare = dataRowCount > 0 ? badRows.length / dataRowCount : 0;
    if (badRows.length > 0 && badShare > MAX_BAD_ROW_SHARE) {
      var pct = Math.round(badShare * 100);
      errors.push(
        'Не удалось прочитать ' + pct + '% строк с данными (' + badRows.length + ' из ' + dataRowCount +
        '), слишком много, чтобы доверять результату. Обычно это другой разделитель столбцов, ' +
        'сдвинутая строка заголовков или не та выгрузка. Первые проблемы — ниже.'
      );
      errors = errors.concat(badRows.slice(0, 10));
      if (badRows.length > 10) errors.push('… и ещё ' + (badRows.length - 10) + ' ' + rowWord(badRows.length - 10) + '.');
      return fail();
    }

    // A small number of problem rows becomes warnings carrying the line numbers.
    if (badRows.length > 0) {
      warnings.push('Пропущено из-за нечитаемых значений: ' + badRows.length + ' ' + rowWord(badRows.length) + '.');
      warnings = warnings.concat(badRows.slice(0, 20));
      if (badRows.length > 20) warnings.push('… и ещё ' + (badRows.length - 20) + ' ' + rowWord(badRows.length - 20) + ' пропущено.');
    }

    if (employees.length === 0) {
      errors.push('В файле не нашлось ни одной пригодной строки сотрудника. Проверьте, что данные идут сразу под строкой заголовков.');
      return fail();
    }

    // A gap cannot be computed without both genders.
    if (genderSeen.F === 0 || genderSeen.M === 0) {
      var have = genderSeen.F > 0 ? 'F' : 'M';
      errors.push(
        'В файле распознан только один пол (' + have + '). Для разрыва нужны и женщины, и мужчины, ' +
        'сравнивать не с чем. Проверьте столбец gender: подходят Ж/М, жен/муж, F/M, female/male.'
      );
      return fail();
    }

    return { ok: true, employees: employees, errors: [], warnings: warnings };
  }

  // --- Optional browser helper ---------------------------------------------

  /**
   * Reads a File from <input type="file"> and hands a ParseResult to a callback.
   * The only place where part C touches a browser API. No network is involved:
   * FileReader reads a local file.
   */
  function readFile(file, callback) {
    if (typeof FileReader === 'undefined') {
      callback({ ok: false, employees: [], errors: ['Этот браузер не умеет читать файлы.'], warnings: [] });
      return;
    }
    var reader = new FileReader();
    reader.onload = function () { callback(parse(reader.result)); };
    reader.onerror = function () {
      callback({ ok: false, employees: [], errors: ['Файл не удалось прочитать. Сохраните его заново как CSV и загрузите ещё раз.'], warnings: [] });
    };
    reader.readAsText(file, 'utf-8');
  }

  return {
    parse: parse,
    readFile: readFile,
    REQUIRED_COLUMNS: REQUIRED.slice(),
    DEFAULTS: { fte: DEFAULTS.fte, months_worked: DEFAULTS.months_worked },
    // Exposed for tests and the Method section; the UI need not reach in here.
    _internals: {
      normaliseHeader: normaliseHeader,
      canonicalHeader: canonicalHeader,
      parseNumber: parseNumber,
      parseGender: parseGender,
      detectDelimiter: detectDelimiter,
      SYNONYMS: SYNONYMS
    }
  };
})();

// Export for Node — only so the tests can run outside a browser.
// In the built page this branch does not exist: typeof module === 'undefined'.
if (typeof module !== 'undefined' && module.exports) module.exports = CSV;
