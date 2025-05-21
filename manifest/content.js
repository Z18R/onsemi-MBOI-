const text = document.body.innerText;

const blob = new Blob([text], { type: "text/plain" });
const url = URL.createObjectURL(blob);

const a = document.createElement("a");
a.href = url;
a.download = "notepad.txt"; // filename
a.click();

URL.revokeObjectURL(url);
