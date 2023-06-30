const form = document.querySelector("#img-form");
const doc = document.querySelector("#doc");
const filename = document.querySelector("#filename");
const replaceInput = document.querySelector("#replace");
const findInput = document.querySelector("#find");

// Load image and show form
function loadFile(e) {
  const file = e.target.files[0];

  // Check if file is an image
  if (!isFileText(file)) {
    alertError("Please select an image");
    return;
  }

  findInput.value = "";
  replaceInput.value = "";
  // Show form, image name
  form.style.display = "block";
  filename.innerHTML = doc.files[0].name;
}

// Make sure file is text
function isFileText(file) {
  const acceptedFileTypes = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ];
  return file && acceptedFileTypes.includes(file["type"]);
}

// Resize image
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
  const filePath = doc.files[0].path;
  const find = findInput.value;
  const replace = replaceInput.value;

  ipcRenderer.send("file:edit", {
    filePath,
    find,
    replace,
  });
}

// When done, show message
ipcRenderer.on("file:done", () => {
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
doc.addEventListener("change", loadFile);
// Form submit listener
form.addEventListener("submit", editText);
