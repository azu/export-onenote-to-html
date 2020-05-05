#!/usr/bin/env node
'use strict';
const fs = require('fs');
const path = require('path');
const meow = require('meow');
const TurndownService = require('turndown')
const turndownService = new TurndownService()
const normalizeAlt = (alt) => {
    return alt
        .replace(/(\r\n|\r|\n)/g, " ")
        .replace(/([\p{sc=Hiragana}\p{sc=Han}\p{sc=Katakana}])\s+/ug, "$1")
}
turndownService.addRule('alt', {
    filter: ['img'],
    replacement: function (content, node) {
        if (node.src.includes(".cdn.office.net")) {
            return "";
        }
        const alt = normalizeAlt(node.alt || '')
        const src = node.getAttribute('src') || ''
        const title = node.title || ''
        const titlePart = title ? ' "' + title + '"' : ''
        const imageTag = '![' + alt + ']' + '(' + src + titlePart + ')'
            + (
                alt ? `
> ${alt}
`
                    : "")
        return src ? imageTag : ''
    }
})
const sharp = require('sharp');
const cli = meow(`
	Usage
	  $ export-onenote-to-html <input>

	Options
	  --output path to output directory

	Examples
	  $ export-onenote-to-html path/to/index.html --output output/
`, {
    flags: {
        output: {
            type: 'string'
        }
    }
});
const convertImage = (fileName, base64) => {
    const decodedFile = Buffer.from(base64, 'base64')
    return sharp(decodedFile)
        .toFile(fileName);
}
const cleanupHTML = (html) => {
    return html
        .split(`<img class="one_OutlineElementHandle_16x16x32" unselectable="on" role="presentation" src="https://c1-onenote-15.cdn.office.net:443/o/s/161290131900_resources/1033/one.png">`).join("")
        .split(`<img class="one_Resize_8x5x32" unselectable="on" role="presentation" src="https://c1-onenote-15.cdn.office.net:443/o/s/161290131900_resources/1033/one.png" alt="Resize the Outline" title="Resize the Outline">`).join("")
        .split(`<span unselectable="on" class="cui-taskPaneTitle" id="AppForOfficePanel0-title">Immersive Reader</span>`).join("")
}
const run = async (inputFile, {
    output
}) => {
    fs.mkdirSync(output, {
        recursive: true
    });
    const inputContent = fs.readFileSync(inputFile, "utf-8");
    let outputContent = inputContent;
    // ="data:image/tif;base64,TU0AKgAdNTz/////////////"
    const results = inputContent.matchAll(/(data:)([\w\/+]+);(charset=[\w-]+|base64).*?,([^"']+)/gi);
    const items = Array.from(results).reverse();
    const tasks = items.map(async (result, index) => {
        const matchIndex = result.index || 0;
        const match = result[0] || "";
        // force png
        const fileName = (items.length - index) + ".png"
        await convertImage(path.join(output, fileName), result[4]);
        outputContent = outputContent.replace(match, fileName)
    });
    await Promise.all(tasks);
    let cleanHTML = cleanupHTML(outputContent);
    const doc = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>${path.basename(inputFile, ".html")}</title>
</head>
<body>
${cleanHTML}
</body>
</html>
`
    fs.writeFileSync(path.join(output, "index.html"), doc, "utf-8");
    fs.writeFileSync(path.join(output, "README.md"), turndownService.turndown(cleanHTML), "utf-8");
}

if (!module.parent) {
    run(cli.input[0], cli.flags)
} else {
    module.exports = run;
}
