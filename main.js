const { PythonShell } = require("python-shell");
const path = require("path");
const { app, BrowserWindow, Menu, ipcMain } = require("electron");

// const isDev = process.env.NODE_ENV !== "production";
const isDev = false;
const isMac = process.platform === "darwin";

let mainWindow;
let aboutWindow;

// Main Window
function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: isDev ? 1000 : 600,
    height: 600,
    icon: `${__dirname}/assets/icons/Icon_256x256.png`,
    resizable: isDev,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  // Show devtools automatically if in development
  if (isDev) {
    mainWindow.webContents.openDevTools();
  }

  // mainWindow.loadURL(`file://${__dirname}/renderer/index.html`);
  mainWindow.loadFile(path.join(__dirname, "./renderer/index.html"));
}

// About Window
function createAboutWindow() {
  aboutWindow = new BrowserWindow({
    width: 300,
    height: 300,
    title: "About Electron",
    icon: `${__dirname}/assets/icons/Icon_256x256.png`,
  });

  aboutWindow.loadFile(path.join(__dirname, "./renderer/about.html"));
}

// When the app is ready, create the window
app.on("ready", () => {
  createMainWindow();

  const mainMenu = Menu.buildFromTemplate(menu);
  Menu.setApplicationMenu(mainMenu);

  // Remove variable from memory
  mainWindow.on("closed", () => (mainWindow = null));
});

// Menu template
const menu = [
  ...(isMac
    ? [
        {
          label: app.name,
          submenu: [
            {
              label: "About",
              click: createAboutWindow,
            },
          ],
        },
      ]
    : []),
  {
    role: "fileMenu",
  },
  ...(!isMac
    ? [
        {
          label: "Help",
          submenu: [
            {
              label: "About",
              click: createAboutWindow,
            },
          ],
        },
      ]
    : []),
  ...(isDev
    ? [
        {
          label: "Developer",
          submenu: [
            { role: "reload" },
            { role: "forcereload" },
            { type: "separator" },
            { role: "toggledevtools" },
          ],
        },
      ]
    : []),
];

ipcMain.on("file:edit", (e, options) => {
  editText(options);
});

function editText({ folderPath, find, replace }) {
  let options = {
    mode: "text",
    args: [folderPath, find, replace],
  };
  PythonShell.run("editor.py", options, function (err) {
    console.log(err);
  }).then((message) => {
    console.log(message);
    mainWindow.webContents.send("file:done");
  });
}

// Quit when all windows are closed.
app.on("window-all-closed", () => {
  if (!isMac) app.quit();
});

// Open a window if none are open (macOS)
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
});
