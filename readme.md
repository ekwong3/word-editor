# Word Editor

Electron application that allows you to select Word files and find and replace

## Usage

Install dependencies:

```bash

npm install
```

Run:

```bash
npm start
```

You can also use `Electronmon` to constantly run and not have to reload after making changes

```bash
npx electronmon .
```

## Developer Mode

If your `NODE_ENV` is set to `development` then you will have the dev tools enabled and available in the menu bar. It will also open them by default.

When set to `production`, the dev tools will not be available.

## Packaging

Currently on MacOS, use `npx electron-packager . WordEditor --prune=true --icon=./build/icon.icns` to package.

## Icon

Icon from https://www.iconfinder.com/icons/1626830/arrows_blue_circle_refresh_reload_sync_icon

## Original Files

Original files taken from https://github.com/bradtraversy/image-resizer-electron
