```
1. Упорядочить символы по частоте появления.
2. Повторять шаги деления списка символов на две половины до тех пор, пока каждая половина не будет содержать ровно один символ.
3. Присваивать кодовые значения "0" левой половине и "1" правой.
4. Продолжать процедуру до полного формирования индивидуальных кодов для каждого символа.
```



const text: string = "За двумя зайцами погонишься — ни одного кабана не поймаешь".toLowerCase();
const k: number = 3; // длина блока
const elementarySignals: number = 2; // бинарное кодирование

interface BlockProbability {
    block: string;
    probability: number;
    count: number;
}

interface BlockCode {
    block: string;
    probability: number;
    code: string;
    codeLength: number;
}

// Функция для разбиения текста на блоки длины k
function splitIntoBlocks(text: string, k: number): string[] {
    // Убираем пробелы и дефисы для кодирования
    const cleanText = text.replace(/\s/g, "").replace(/—/g, "").replace(/-/g, "");
    const blocks: string[] = [];
    
    for (let i = 0; i < cleanText.length; i += k) {
        const block = cleanText.slice(i, i + k);
        if (block.length === k) {
            blocks.push(block);
        }
    }
    
    return blocks;
}

// Функция для подсчета вероятностей блоков
function calculateBlockProbabilities(blocks: string[]): BlockProbability[] {
    const blockCounts = new Map<string, number>();
    const totalBlocks = blocks.length;
    
    // Подсчитываем количество каждого блока
    for (const block of blocks) {
        blockCounts.set(block, (blockCounts.get(block) || 0) + 1);
    }
    
    // Преобразуем в массив с вероятностями
    const result: BlockProbability[] = [];
    blockCounts.forEach((count, block) => {
        result.push({
            block,
            count,
            probability: count / totalBlocks
        });
    });
    
    // Сортируем по убыванию вероятности
    return result.sort((a, b) => b.probability - a.probability);
}

// Функция для кодирования методом Шеннона-Фано
function shannonFanoEncode(blockProbs: BlockProbability[]): Record<string, string> {
    const result: Record<string, string> = {};
    
    function divideAndCode(items: BlockProbability[], prefix: string = '') {
        if (items.length === 1) {
            result[items[0].block] = prefix;
            return;
        }
        
        if (items.length === 0) {
            return;
        }
        
        // Вычисляем общую вероятность
        const totalProb = items.reduce((sum, item) => sum + item.probability, 0);
        const targetProb = totalProb / 2;
        
        // Находим оптимальную точку разделения, минимизируя разницу между группами
        let bestSplitIndex = 1;
        let minDiff = Infinity;
        let sumLeft = 0;
        
        // Перебираем все возможные точки разделения
        for (let i = 0; i < items.length - 1; i++) {
            sumLeft += items[i].probability;
            const sumRight = totalProb - sumLeft;
            const diff = Math.abs(sumLeft - sumRight);
            
            // Находим разделение с минимальной разницей
            if (diff < minDiff) {
                minDiff = diff;
                bestSplitIndex = i + 1;
            }
        }
        
        // Гарантируем, что в каждой группе хотя бы один элемент
        if (bestSplitIndex <= 0) {
            bestSplitIndex = 1;
        }
        if (bestSplitIndex >= items.length) {
            bestSplitIndex = items.length - 1;
        }
        
        const firstHalf = items.slice(0, bestSplitIndex);
        const secondHalf = items.slice(bestSplitIndex);
        
        // Рекурсивно кодируем каждую половину
        divideAndCode(firstHalf, prefix + '0');
        divideAndCode(secondHalf, prefix + '1');
    }
    
    divideAndCode(blockProbs);
    return result;
}

// Функция для вычисления среднего числа элементарных сигналов
function calculateAverageCodeLength(
    blockProbs: BlockProbability[],
    codes: Record<string, string>
): number {
    let total = 0;
    
    for (const blockProb of blockProbs) {
        const code = codes[blockProb.block];
        if (code) {
            total += blockProb.probability * code.length;
        }
    }
    
    return total;
}

// Основная функция
function main() {
    console.log("Исходный текст:", text);
    console.log(`Длина блока: ${k}`);
    console.log(`Количество элементарных сигналов: ${elementarySignals}\n`);
    
    // Разбиваем на блоки
    const blocks = splitIntoBlocks(text, k);
    console.log(`Количество блоков: ${blocks.length}`);
    console.log("Блоки:", blocks.join(", "));
    console.log();
    
    // Вычисляем вероятности блоков
    const blockProbs = calculateBlockProbabilities(blocks);
    console.log("Вероятности блоков:");
    blockProbs.forEach(bp => {
        console.log(`  ${bp.block}: ${bp.count}/${blocks.length} = ${bp.probability.toFixed(4)}`);
    });
    console.log();
    
    // Кодируем методом Шеннона-Фано
    const codes = shannonFanoEncode(blockProbs);
    
    // Формируем результат с кодами
    const encodedBlocks: BlockCode[] = blockProbs.map(bp => ({
        block: bp.block,
        probability: bp.probability,
        code: codes[bp.block] || '',
        codeLength: codes[bp.block]?.length || 0
    }));
    
    console.log("Кодирование методом Шеннона-Фано:");
    console.log("┌─────────────┬──────────────┬─────────────┬──────────────┐");
    console.log("│   Блок      │ Вероятность  │     Код     │ Длина кода   │");
    console.log("├─────────────┼──────────────┼─────────────┼──────────────┤");
    
    encodedBlocks.forEach(eb => {
        const blockStr = eb.block.padEnd(11);
        const probStr = eb.probability.toFixed(4).padStart(12);
        const codeStr = eb.code.padEnd(11);
        const lenStr = eb.codeLength.toString().padStart(12);
        console.log(`│ ${blockStr} │ ${probStr} │ ${codeStr} │ ${lenStr} │`);
    });
    console.log("└─────────────┴──────────────┴─────────────┴──────────────┘");
    console.log();
    
    // Вычисляем среднее число элементарных сигналов на одну k-буквенную комбинацию
    const avgCodeLength = calculateAverageCodeLength(blockProbs, codes);
    console.log(`Среднее число элементарных сигналов, приходящееся на одну ${k}-буквенную комбинацию: ${avgCodeLength.toFixed(4)}`);
    
    // Вычисляем среднее число элементарных сигналов на одну букву
    const avgPerLetter = avgCodeLength / k;
    console.log(`Среднее число элементарных сигналов, приходящееся на одну букву: ${avgPerLetter.toFixed(4)}`);
    
    // Вычисляем энтропию одной буквы для сравнения
    const allChars = text.replace(/\s/g, "").replace(/—/g, "").replace(/-/g, "");
    const charCounts = new Map<string, number>();
    for (const char of allChars) {
        charCounts.set(char, (charCounts.get(char) || 0) + 1);
    }
    
    let entropy = 0;
    charCounts.forEach((count, char) => {
        const prob = count / allChars.length;
        entropy -= prob * Math.log2(prob);
    });
    
    console.log(`Энтропия одной буквы: ${entropy.toFixed(4)}`);
    console.log(`Теоретический минимум (H/log k): ${(entropy / Math.log2(elementarySignals)).toFixed(4)}`);
}

// Запуск программы
main();

