// ==UserScript==
// @name         OneNote: downloader
// @namespace    http://efcl.info/
// @version      0.1
// @description  Download onenote content
// @author       azu
// @match        https://onenote.officeapps.live.com/o/onenoteframe.aspx?*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';
    const requestSet = new Set();
    const downloadContent = (fileName) => {
        const WACDocumentPanelContent = document.querySelector("#WACDocumentPanelContent");
        const downLoadLink = document.createElement("a");
        downLoadLink.download = fileName + ".html";
        downLoadLink.href = URL.createObjectURL(new Blob([WACDocumentPanelContent.innerHTML], { type: "text/html" }));
        downLoadLink.click();
    }

    const download = () => {
        const title = document.querySelector(".Title").ariaLabel.replace("Page Title ","");
        downloadContent(title);
    }
    const downloadButton = document.createElement("button");
    downloadButton.textContent = "Download";
    downloadButton.setAttribute("style", "font-size: 16px; color: gray; position: fixed; top:0; right: 0; padding: 2px;");
    downloadButton.addEventListener("click", () => {
        download();
    });
    document.body.appendChild(downloadButton);
    const origOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function (...args) {
        const originalURL = args[1];
        try {
            const url = new URL(originalURL);
            if (url.pathname === "/onenoteonlinesync/v1/GetBase64Image") {
                requestSet.add(originalURL);
                console.log("loading", originalURL);
                this.addEventListener("load", () => {
                    requestSet.delete(originalURL);
                    console.log("loaded", originalURL, "still downloading count", requestSet.size);
                });
            }
        } catch (error) {
            // console.error(error);
        } finally {
            origOpen.apply(this, args);
        }
    };
    setInterval(() => {
        if (requestSet.size === 0){
            downloadButton.style.color = "blue";
        } else {
            downloadButton.style.color = "gray";
        }
    }, 1000);
})();
