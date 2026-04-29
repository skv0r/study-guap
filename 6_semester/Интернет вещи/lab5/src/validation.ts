/**
 * ЛР5 (пособие Шевяков, Савельева): проверка входящих данных на стороне системы.
 * Аналог try/except + приведение типов из Python; для строк — словарь допустимых значений.
 */

/** Допустимый вид числовой температуры в строке (точка или запятая как разделитель). */
export const TEMPERATURE_STRING_RE = /^[-+]?\d+([.,]\d+)?$/;

/** Словарь допустимых строк для замка (пособие: проверка вхождения в набор строк). */
const LOCK_STRING_TO_BOOL = new Map<string, boolean>([
    ["true", true],
    ["1", true],
    ["on", true],
    ["закрыт", true],
    ["closed", true],
    ["false", false],
    ["0", false],
    ["off", false],
    ["открыт", false],
    ["open", false]
]);

/**
 * Два строковых формата температуры: «21.5» и «21,5».
 * Сначала проверка регулярным выражением, затем безопасное приведение к числу.
 */
export function parseTemperatureCelsius(raw: string): { ok: true; value: number } | { ok: false; error: string } {
    const s = raw.trim();
    if (s === "") {
        return { ok: false, error: "Пустое значение температуры" };
    }
    if (!TEMPERATURE_STRING_RE.test(s)) {
        return {
            ok: false,
            error:
                "Строка не соответствует формату числа (допустимы цифры, опционально «.» или «,» как разделитель дробной части)"
        };
    }
    try {
        const normalized = s.replace(",", ".");
        const value = Number(normalized);
        if (Number.isNaN(value)) {
            return { ok: false, error: "Не удалось привести значение к числу (ожидался числовой формат)" };
        }
        if (value < -60 || value > 80) {
            return { ok: false, error: "Температура вне допустимого диапазона для модели (-60…80 °C)" };
        }
        return { ok: true, value };
    } catch {
        return { ok: false, error: "Ошибка при разборе температуры" };
    }
}

/** Яркость: целое 0–100, строка из цифр (второй по смыслу формат — целая строка). */
export const BRIGHTNESS_DIGITS_RE = /^\d{1,3}$/;

export function parseBrightnessPercent(raw: string): { ok: true; value: number } | { ok: false; error: string } {
    const s = raw.trim();
    if (!BRIGHTNESS_DIGITS_RE.test(s)) {
        return { ok: false, error: "Яркость должна быть строкой из 1–3 десятичных цифр (0–100)" };
    }
    try {
        const n = Number.parseInt(s, 10);
        if (n < 0 || n > 100) {
            return { ok: false, error: "Яркость должна быть в диапазоне 0–100" };
        }
        return { ok: true, value: n };
    } catch {
        return { ok: false, error: "Не удалось разобрать яркость" };
    }
}

/**
 * Строковый режим замка: проверка по словарю допустимых строк (как в пособии для месяцев).
 */
export function parseLockedFromString(raw: string): { ok: true; locked: boolean } | { ok: false; error: string } {
    const key = raw.trim().toLowerCase();
    if (key === "") {
        return { ok: false, error: "Пустое значение состояния замка" };
    }
    const locked = LOCK_STRING_TO_BOOL.get(key);
    if (locked === undefined) {
        return {
            ok: false,
            error:
                "Недопустимая строка: используйте true/false, 1/0, on/off, да/нет, закрыт/открыт (без учёта регистра)"
        };
    }
    return { ok: true, locked };
}

/** Режим света «строкой»: on/off/true/false — проверка регулярным выражением. */
export const LIGHT_ENABLED_WORD_RE = /^(on|off|true|false|1|0|yes|no|да|нет)$/i;

export function parseLightEnabledString(raw: string): { ok: true; enabled: boolean } | { ok: false; error: string } {
    const s = raw.trim();
    if (!LIGHT_ENABLED_WORD_RE.test(s)) {
        return {
            ok: false,
            error:
                "Параметр enabled/on: допустимы только слова on, off, true, false, 1, 0, yes, no, да, нет (регулярная проверка)"
        };
    }
    const key = s.toLowerCase();
    const on = key === "on" || key === "true" || key === "1" || key === "yes" || key === "да";
    const off = key === "off" || key === "false" || key === "0" || key === "no" || key === "нет";
    if (on) return { ok: true, enabled: true };
    if (off) return { ok: true, enabled: false };
    return { ok: false, error: "Не удалось интерпретировать состояние света" };
}
