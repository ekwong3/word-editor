const form = document.querySelector("#text-form");
const select = document.querySelector("#select");
const folderName = document.querySelector("#folder-name");
const doc = document.querySelector("#doc");
const replaceInput = document.querySelector("#replace");
const findInput = document.querySelector("#find");
const caseCheck = document.getElementById("case");
const wordCheck = document.getElementById("whole-word");

function reverseString(str) {
  return str.split("").reverse().join("");
}

function getFolderPath() {
  if (doc.files.length < 1) {
    alertError("Please upload a folder with files");
    select.innerHTML = "Select a folder of files to edit";
    folderName.innerHTML = "";
    return;
  }
  const reversedPath = reverseString(doc.files[0].path);
  const firstSlash = reversedPath.search("/");
  const folderPath = reverseString(reversedPath.slice(firstSlash));
  return folderPath;
}

// Load image and show form
function loadFile() {
  const folderPath = getFolderPath();
  select.innerHTML = "Editing files in&nbsp";
  folderName.innerHTML = folderPath;
}

function reset() {
  findInput.value = "";
  replaceInput.value = "";
  doc.file = null;
  caseCheck.checked = false;
  wordCheck.checked = false;
  select.innerHTML = "Select a folder of files to edit";
  folderName.innerHTML = "";
}
// Edit text
function editText(e) {
  e.preventDefault();

  if (!doc.files[0]) {
    alertError("Please upload a Word file");
    return;
  }

  if (findInput.value === "" || replaceInput.value === "") {
    alertError("Please enter valid strings");
    return;
  }

  // Electron adds a bunch of extra properties to the file object including the path
  const folderPath = getFolderPath();
  const find = findInput.value;
  const replace = replaceInput.value;
  const matchCase = caseCheck.checked;
  const matchWord = wordCheck.checked;

  ipcRenderer.send("file:edit", {
    folderPath,
    find,
    replace,
    matchCase,
    matchWord,
  });
}

// When done, show message
ipcRenderer.on("file:done", () => {
  reset();
  alertSuccess("Text edited!");
});

function alertSuccess(message) {
  Toastify.toast({
    text: message,
    duration: 5000,
    close: false,
    style: {
      background: "green",
      color: "white",
      textAlign: "center",
    },
  });
}

function alertError(message) {
  Toastify.toast({
    text: message,
    duration: 5000,
    close: false,
    style: {
      background: "red",
      color: "white",
      textAlign: "center",
    },
  });
}

// File select listener
doc.addEventListener("input", loadFile);
// Form submit listener
form.addEventListener("submit", editText);
