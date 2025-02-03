$("#register-event-handlers").click(() => tryCatch(beginTimer));

async function beginTimer() {
  await Excel.run(async (context) => {
    // Add a selection changed event handler for the workbook.
    setInterval(updateInfo, 2000);
    console.log("updating cells on a timer");
    await context.sync();
  });
}

let clicked = false;

i = 0;

async function updateInfo() {
  await Excel.run(async (context) => {
    //get selected cells value:

    const worksheets = [];

    await Excel.run(async (context) => {
      let sheets = context.workbook.worksheets;
      sheets.load("items/name");

      await context.sync();

      console.log(sheets.items, "sheets.items");

      if (sheets.items.length > 1) {
        console.log(
          `There are ${sheets.items.length} worksheets in the workbook:`
        );
      } else {
        console.log(`There is one worksheet in the workbook:`);
      }

      sheets.items.forEach(function (sheet) {
        console.log(sheet.name);
        worksheets.push(sheet.name);
        console.log(worksheets, "worksheets");
      });
    });

    let sheet = context.workbook.worksheets.getItem(worksheets[i]);
    let range = sheet.getRange("B1:B6");
    range.load("text");
    await context.sync();
    let className = range.text[0][0];
    let firstPlace = range.text[1][0];
    let secondPlace = range.text[2][0];
    let thirdPlace = range.text[3][0];

    //iterate:
    i === worksheets.length - 1 ? (i = 0) : i++;

    if (!firstPlace) {
      updateInfo();
      return;
    }

    //connect to OBS Websocket localhost
    //Get websocket connection info
    //Enter the websocketIP address
    const websocketIP = document.getElementById("IP").value;

    //Enter the OBS websocket port number
    const websocketPort = document.getElementById("Port").value;

    //Enter the OBS websocket server password
    const websocketPassword = document.getElementById("PW").value;

    var obs = new OBSWebSocket();
    console.log(`ws://${websocketIP}:${websocketPort}`);
    // connect to OBS websocket
    try {
      const { obsWebSocketVersion, negotiatedRpcVersion } = await obs.connect(
        `ws://${websocketIP}:${websocketPort}`,
        websocketPassword,
        {
          rpcVersion: 1,
        }
      );
      console.log(
        `Connected to server ${obsWebSocketVersion} (using RPC ${negotiatedRpcVersion})`
      );
    } catch (error) {
      console.error("Failed to connect", error.code, error.message);
    }
    obs.on("error", (err) => {
      console.error("Socket error:", err);
    });

    //set OBS Scene
    await obs.call("SetCurrentProgramScene", {
      sceneName: document.getElementById("Scene").value,
    });

    //set OBS source text

    await obs.call(
      "SetInputSettings",
      {
        inputName: document.getElementById("ClassField").value,
        inputSettings: {
          text: className,
        },
      },
      (err, data) => {
        /* Error message and data. */
        // console.log('Using call SetInputSettings:', err, data);
      }
    );

    await obs.call(
      "SetInputSettings",
      {
        inputName: document.getElementById("Field1").value,
        inputSettings: {
          text: firstPlace,
        },
      },
      (err, data) => {
        /* Error message and data. */
        // console.log('Using call SetInputSettings:', err, data);
      }
    );

    await obs.call(
      "SetInputSettings",
      {
        inputName: "2nd",
        inputSettings: {
          text: secondPlace ? "2nd:" : " ",
        },
      },
      (err, data) => {
        /* Error message and data. */
        // console.log('Using call SetInputSettings:', err, data);
      }
    );
    await obs.call(
      "SetInputSettings",
      {
        inputName: document.getElementById("Field2").value,
        inputSettings: {
          text: secondPlace,
        },
      },
      (err, data) => {
        /* Error message and data. */
        // console.log('Using call SetInputSettings:', err, data);
      }
    );

    await obs.call(
      "SetInputSettings",
      {
        inputName: "3rd",
        inputSettings: {
          text: thirdPlace ? "3rd:" : " ",
        },
      },
      (err, data) => {
        /* Error message and data. */
        // console.log('Using call SetInputSettings:', err, data);
      }
    );
    await obs.call(
      "SetInputSettings",
      {
        inputName: document.getElementById("Field3").value,
        inputSettings: {
          text: thirdPlace,
        },
      },
      (err, data) => {
        /* Error message and data. */
        // console.log('Using call SetInputSettings:', err, data);
      }
    );

    await obs.disconnect();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
