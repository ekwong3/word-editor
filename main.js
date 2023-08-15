const child = require("child_process").execFile;
const path = require("path");
const { app, BrowserWindow, Menu, ipcMain } = require("electron");

const isDev = false;

function editText({
  folderPath,
  find,
  replace,
  keepCase,
  matchWord,
  processSub,
}) {
  const execPath = path.join(__dirname, "./scripts/editor");
  const params = [folderPath, find, replace, keepCase, matchWord, processSub];
  child(execPath, params, { shell: true }, function (err, message) {
    if (err) {
      mainWindow.webContents.send("file:error", err);
    } else {
      mainWindow.webContents.send("file:done", message);
    }
  });
}

let mainWindow;
let aboutWindow;

// Main Window
function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: isDev ? 1000 : 600,
    height: 580,
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
  ...[
    {
      role: "fileMenu",
    },
    {
      label: app.name,
      submenu: [
        {
          label: "About",
          click: createAboutWindow,
        },
      ],
      label: "Application",
      submenu: [
        {
          label: "About Application",
          selector: "orderFrontStandardAboutPanel:",
        },
        { type: "separator" },
        {
          label: "Quit",
          accelerator: "Command+Q",
          click: function () {
            app.quit();
          },
        },
      ],
    },
    {
      label: "Edit",
      submenu: [
        { label: "Undo", accelerator: "CmdOrCtrl+Z", selector: "undo:" },
        {
          label: "Redo",
          accelerator: "Shift+CmdOrCtrl+Z",
          selector: "redo:",
        },
        { type: "separator" },
        { label: "Cut", accelerator: "CmdOrCtrl+X", selector: "cut:" },
        { label: "Copy", accelerator: "CmdOrCtrl+C", selector: "copy:" },
        { label: "Paste", accelerator: "CmdOrCtrl+V", selector: "paste:" },
        {
          label: "Select All",
          accelerator: "CmdOrCtrl+A",
          selector: "selectAll:",
        },
      ],
    },
  ],
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

// Quit when all windows are closed.
app.on("window-all-closed", () => {
  app.quit();
});

// Open a window if none are open (macOS)
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
});
