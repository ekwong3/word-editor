const form = document.querySelector("#text-form");
const select = document.querySelector("#select");
const folderName = document.querySelector("#folder-name");
const folder = document.querySelector("#folderpath");
const replaceInput = document.querySelector("#replace");
const findInput = document.querySelector("#find");
const caseCheck = document.querySelector("#case");
const matchWordCheck = document.querySelector("#match-word");
const processSubCheck = document.querySelector("#subfolders");
const modal = document.querySelector("#modal");
const modalHeader = document.querySelector("#modal-header");
const span = document.querySelector("#close");
const title = document.querySelector("#modal-title");
const content = document.querySelector("#modal-content");

// When the user clicks on <span> (x), close the modal
span.onclick = function () {
  modal.style.display = "none";
};

// When the user clicks anywhere outside of the modal, close it
window.onclick = function (event) {
  if (event.target == modal) {
    modal.style.display = "none";
  }
};

function reverseString(str) {
  return str.split("").reverse().join("");
}

function getFolderPath() {
  if (folder.files.length < 1) {
    alertError("missing-files");
    select.innerHTML = "Select a folder of files to edit";
    folderName.innerHTML = "";
    return;
  }
  const reversedPath = reverseString(folder.files[0].path);
  const firstSlash = reversedPath.search("/");
  const folderPath = reverseString(reversedPath.slice(firstSlash));
  return folderPath;
}

// Load image and show form
function loadFile() {
  if (!folder.value) {
    resetFolder();
    return;
  }

  const folderPath = getFolderPath();
  select.innerHTML = "Editing files in&nbsp";
  folderName.innerHTML = folderPath;
}

function resetFolder() {
  folder.value = "";
  select.innerHTML = "Select a folder of files to edit";
  folderName.innerHTML = "";
}

function reset() {
  findInput.value = "";
  replaceInput.value = "";
  caseCheck.checked = false;
  processSubCheck.checked = false;
  matchWordCheck.checked = false;
  resetFolder();
}

// Edit text
function editText(e) {
  e.preventDefault();

  if (!folder.files[0]) {
    alertError("missing-files");
    return;
  }

  if (findInput.value === "" || replaceInput.value === "") {
    alertError("missing-strings");
    return;
  }

  // Electron adds a bunch of extra properties to the file object including the path
  const folderPath = `"${getFolderPath()}"`;
  const find = `"${findInput.value}"`;
  const replace = `"${replaceInput.value}"`;
  const keepCase = caseCheck.checked;
  const matchWord = matchWordCheck.checked;
  const processSub = processSubCheck.checked;

  alertEditing();

  ipcRenderer.send("file:edit", {
    folderPath,
    find,
    replace,
    keepCase,
    matchWord,
    processSub,
  });
}

// When done, show message
ipcRenderer.on("file:done", (message) => {
  reset();
  console.log(message);
  alertSuccess("Text edited!");
});

ipcRenderer.on("file:error", (error) => {
  alertError("file-error", error);
});

function alertEditing() {
  modalHeader.style.backgroundColor = "rgb(124 124 124)";
  modal.style.display = "block";
  title.innerHTML = "Editing in process";
  content.innerHTML = "Files are currently being edited";
}

function alertSuccess(message) {
  modalHeader.style.backgroundColor = "rgb(15 118 110)";
  modal.style.display = "block";
  title.innerHTML = message;
  content.innerHTML = "Files were successfully edited";
}

function alertError(error, errMsg = "") {
  modalHeader.style.backgroundColor = "rgb(194 32 14)";
  modal.style.display = "block";
  switch (error) {
    case "missing-files":
      title.innerHTML = "Missing Files";
      content.innerHTML = "Please upload a folder with valid files";
      break;
    case "missing-strings":
      title.innerHTML = "Missing Text";
      content.innerHTML = "Please enter words to find and replace";
      break;
    case "file-error":
      title.innerHTML = "File Error";
      content.innerHTML = errMsg;
      break;
    default:
      title.innerHTML = "Error";
      content.innerHTML = "Please try again!";
  }
}

// File select listener
folder.addEventListener("input", loadFile);
// Form submit listener
form.addEventListener("submit", editText);
