let word: string = "ЗЕЛЕНАЯ ЗЕЛЕНЬ ЗЕЛЕНЕЕТ".toLowerCase();

function moveWindowLZ77(word: string): Array<{ offset: number, length: number, next: string }> {

    const result: Array<{ offset: number, length: number, next: string }> = [];

    let wordLength: number = word.length;
    let uniqueSymbols: number = new Set(word).size;
    let library: string[] = [];
    console.log(wordLength, uniqueSymbols);

    let i = 0;
    while (i < wordLength) {
        if (!library.includes(word[i])) {
            library.push(word[i]);
            console.log("запушил ", word[i]);
            result.push({ offset: 0, length: 0, next: word[i] })
            i += 1;
        } else {
            const window = word.slice(0, i);
            const buffer = word.slice(i);

            let bestOffset = 0;
            let bestLength = 0;

            for (let len = 1; len <= buffer.length; len++) {
                const sub = buffer.slice(0, len);
                const idx = window.indexOf(sub);
                if (idx !== -1) {
                    bestOffset = idx + 1; // offset от начала строки, индекс с 1
                    bestLength = len;
                } else {
                    break;
                }
            }

            const nextSymbol = (i + bestLength) < wordLength ? word[i + bestLength] : "";

            result.push({
                offset: bestOffset,
                length: bestLength,
                next: nextSymbol
            });

            console.log(
                `словарь=[${window.slice(-9)}] буфер=[${buffer.slice(0, 7)}] -> (${bestOffset}, ${bestLength}, ${nextSymbol || "∅"})`
            );
            i += bestLength + 1;
        }
    }

    return result;
}

console.log(moveWindowLZ77(word));
