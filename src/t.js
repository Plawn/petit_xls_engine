const quoteRegex = function(str) {
    return str.replace(/([.?*+^$[\]\\(){}|-])/g, "\\$1");
};

/**
 * 
 * @param {string} string 
 * @param {{start:string;end:string}} delimiters
 */
const extractPlaceholders = (string, delimiters) => {
    // Yes, that's right. It's a bunch of brackets and question marks and stuff.
    // const re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?}/g;
    const { start, end } = delimiters;
    const re = new RegExp(quoteRegex(start) + '(?:(.+?):)?(.+?)(?:\\.(.+?))?' + quoteRegex(end), 'g');
    // const re = new RegExp('\\{\\{(?:(.+?):)?(.+?)(?:\\.(.+?))?\\}\\}', 'g');


    let match = null
    let matches = [];
    while ((match = re.exec(string)) !== null) {
        matches.push({
            // removing the last char
            placeholder: match[0],
            type: match[1] || 'normal',
            name: match[2],
            key: match[3],
            full: match[0].length === string.length
        });
    }

    return matches;
};


const test_str = '{{mission.company.name}} {{mission.name}}';

const delimiters = { start: '{{', end: '}}' };

console.log(extractPlaceholders(test_str, delimiters));