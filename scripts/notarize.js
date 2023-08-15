require("dotenv").config();
const { notarize } = require("@electron/notarize");

exports.default = async function notarizing(context) {
  const { electronPlatformName, appOutDir } = context;
  if (electronPlatformName !== "darwin") {
    return;
  }

  const appName = context.packager.appInfo.productFilename;
  const appPath = `${appOutDir}/${appName}.app`;

  try {
    await notarize({
      appBundleId: "com.yourcompany.yourAppId",
      appPath: appPath,
      appleId: process.env.APPLEID,
      appleIdPassword: process.env.APPLEIDPASS,
      teamId: process.env.TEAMID,
    });
  } catch (error) {
    if (error.message?.includes("Failed to staple")) {
      const { spawn } = require("child_process");
      spawn(`xcrun`, ["stapler", "staple", appPath]);
    } else {
      throw error;
    }
  }
};
