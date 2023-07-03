const form = document.querySelector("#text-form");
const select = document.querySelector("#select");
const folderName = document.querySelector("#folder-name");
const doc = document.querySelector("#doc");
const replaceInput = document.querySelector("#replace");
const findInput = document.querySelector("#find");
const caseCheck = document.getElementById("case");
const wordCheck = document.getElementById("whole-word");
const modal = document.getElementById("modal");
const modalHeader = document.getElementById("modal-header");
const span = document.getElementById("close");
const title = document.getElementById("modal-title");
const content = document.getElementById("modal-content");

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
  if (doc.files.length < 1) {
    alertError("missing-files");
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
    alertError("missing-files");
    return;
  }

  if (findInput.value === "" || replaceInput.value === "") {
    alertError("missing-strings");
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

ipcRenderer.on("file:error", () => {
  alertError("file-error");
});

function alertSuccess(message) {
  modalHeader.style.backgroundColor = "rgb(15 118 110)";
  modal.style.display = "block";
  title.innerHTML = message;
  content.innerHTML = "Files were successfully edited";
}

function alertError(error) {
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
      content.innerHTML =
        "There was an error editing the files. Make sure all files are valid Word files.";
      break;
    default:
      title.innerHTML = "Error";
      content.innerHTML = "Please try again!";
  }
}

// File select listener
doc.addEventListener("input", loadFile);
// Form submit listener
form.addEventListener("submit", editText);
