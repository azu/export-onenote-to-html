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
	  --yearDirectory Create {year}/ directory and output into the year directory

	Examples
	  $ export-onenote-to-html path/to/index.html --output output/
`, {
    flags: {
        output: {
            type: "string"
        },
        yearDirectory: {
            type: "boolean"
        }
    }
});
const convertImage = (fileName, base64) => {
    const decodedFile = Buffer.from(base64, 'base64')
    return sharp(decodedFile)
        .toFile(fileName);
}
const metaMarkdown = (markdown) => {
    return markdown.replace(/Page Contents\n\n(?:(.*)\s*\n\n){1,2}(\d+.*日)\s*\n\n(\d+:\d+)/, `---
title: "$1"
---

# $1

> $2 $3
`)
}
const cleanupHTML = (html) => {
    return html
        .split(`<img class="one_OutlineElementHandle_16x16x32" unselectable="on" role="presentation" src="https://c1-onenote-15.cdn.office.net:443/o/s/161290131900_resources/1033/one.png">`).join("")
        .split(`<img class="one_Resize_8x5x32" unselectable="on" role="presentation" src="https://c1-onenote-15.cdn.office.net:443/o/s/161290131900_resources/1033/one.png" alt="Resize the Outline" title="Resize the Outline">`).join("")
        .split(`<span unselectable="on" class="cui-taskPaneTitle" id="AppForOfficePanel0-title">Immersive Reader</span>`).join("")
}
const run = async (inputFile, {
    output,
    yearDirectory
}) => {
    let outputBaseDir = output
    const inputContent = fs.readFileSync(inputFile, "utf-8");
    if (yearDirectory) {
        let year;
        {
            const dateMatch = /(20\d+)年(\d+)月\d+日.*日<\/span>/;
            const matchResult = inputContent.match(dateMatch);
            if (matchResult) {
                year = matchResult[1]
            }
        }
        {
            const dateMatch = /(20\d+)(!?<.*>)年(!?<.*>)(\d+)(!?<.*>)月/;
            const matchResult = inputContent.match(dateMatch);
            if (matchResult) {
                year = matchResult[1]
            }
        }
        if (year === undefined) {
            throw new Error("Not found year. --yearDirectory does not work: " + output);
        }
        console.log("year", year);
        const relativeOutput = path.relative(process.cwd(), output);
        outputBaseDir = path.join(process.cwd(), year, relativeOutput);
    }
    fs.mkdirSync(outputBaseDir, {
        recursive: true
    });
    let outputContent = inputContent;
    // ="data:image/tif;base64,TU0AKgAdNTz/////////////"
    const results = inputContent.matchAll(/(data:)([\w\/+]+);(charset=[\w-]+|base64).*?,([^"']+)/gi);
    const items = Array.from(results).reverse();
    if (items.length > 0) {
        fs.mkdirSync(path.join(outputBaseDir, "img/"), {
            recursive: true
        })
    }
    const tasks = items.map(async (result, index) => {
        const match = result[0] || "";
        // force convert png
        const filePath = path.join("img/", (items.length - index) + ".png")
        // img/{number}.png
        await convertImage(path.join(outputBaseDir, filePath), result[4]);
        outputContent = outputContent.replace(match, filePath)
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
    const outputMarkdown = metaMarkdown(turndownService.turndown(cleanHTML));
    fs.writeFileSync(path.join(outputBaseDir, "index.html"), doc, "utf-8");
    fs.writeFileSync(path.join(outputBaseDir, "README.md"), outputMarkdown, "utf-8");
    fs.unlinkSync(inputFile);
}

if (!module.parent) {
    run(cli.input[0], cli.flags)
} else {
    module.exports = run;
}
