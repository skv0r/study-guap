const text: string = "В  чужой  монастырь  со  своим  уставом  не  ходят";
// считаем только непустые символы
const totalChars = text.replace(/\s/g, "").length;

interface PercentageType {
    percentage: number;
    text: string;
}

function recursiveGetCounts(text: string): Map<string, number> {
    const charCounts = new Map<string, number>();
    for (const char of text) {
        if (char !== " ") {
            charCounts.set(char.toUpperCase(), (charCounts.get(char.toUpperCase()) || 0) + 1);
        }
    }
    return charCounts;
}

function getPercentage(text: string): PercentageType[] {
    const result: PercentageType[] = [];
    const charCounts = recursiveGetCounts(text);
    charCounts.forEach((count, char) => {
        const percentage = Math.round((count / totalChars) * 1000) / 1000;
        result.push({
            percentage,
            text: char,
        });
    });
    return result.sort((a,b) => b.percentage - a.percentage);
}

console.log(getPercentage(text));

function getFano(charCounts: PercentageType[]): void {
    const result: Record<string, string> = {};

    function divideAndCode(items: PercentageType[], prefix: string = '') {
        if (items.length === 1) {
            result[items[0].text] = prefix;
            return;
        }
        let sumLeft = 0;
        let index = 0;
        let found = false;
        items.forEach((item, i) => {
            if (!found && sumLeft + item.percentage <= 0.5) {
                sumLeft += item.percentage;
                index = i + 1;
            } else if (!found && sumLeft <= 0.5) {
                found = true;
            }
        });

        if (index <= 0 || index >= items.length) {
            index = 1; // миниальный сплит
        }

        const firstHalf = items.slice(0, index);
        const secondHalf = items.slice(index);

        divideAndCode(firstHalf, prefix + '0');
        divideAndCode(secondHalf, prefix + '1');
    }

    divideAndCode(charCounts);

    Object.keys(result).forEach(key => {
        console.log(`Символ '${key}' получил код: ${result[key]}`);
    });
}
const percentages = getPercentage(text);
getFano(percentages);