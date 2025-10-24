let word: string = "Интервьюер интервента интервьюировал".toLowerCase();

function moveWindowLZ77(word: string): Array<{ offset: number, length: number, next: string }> {

    const result: Array<{ offset: number, length: number, next: string }> = [];

    let wordLength: number = word.length;
    let uniqueSymbols: number = new Set(word).size;
    let library: string[] = [];
    console.log(wordLength, uniqueSymbols);

    for (let i = 0; i < wordLength; i++) {
        if (!library.includes(word[i])) {
            library.push(word[i]);
            console.log("запушил ", word[i]);
        } else {
            const window = word.slice(0, i);
            const buffer = word.slice(i);

            let bestOffset = 0;
            let bestLength = 0;

            for (let len = 1; len <= buffer.length; len++) {
                const sub = buffer.slice(0, len);
                const idx = window.lastIndexOf(sub);

                if (idx !== -1) {
                    bestOffset = window.length - idx;
                    bestLength = len;
                } else {
                    break;
                }
            }

            // берём следующий символ после совпадения
            const nextSymbol = word[i + bestLength] ?? "";

            result.push({
                offset: bestOffset,
                length: bestLength,
                next: nextSymbol
            });

            console.log(
                `словарь=[${window.slice(-9)}] буфер=[${buffer.slice(0, 7)}] -> (${bestOffset}, ${bestLength}, ${nextSymbol || "∅"})`
            );
        }
    }

    return result;
}

console.log(moveWindowLZ77(word));
