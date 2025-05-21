// background.js

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "saveText") {
        const blob = new Blob([message.data], { type: "text/plain" });
        const url = URL.createObjectURL(blob);

        // Trigger the download
        chrome.downloads.download({
            url: url,
            filename: "notepad.txt", // File name
            conflictAction: "overwrite" // Overwrite if the file exists
        });

        // Revoke the object URL to free memory
        URL.revokeObjectURL(url);
    }
});