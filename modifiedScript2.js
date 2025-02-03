$("#register-event-handlers").click(() => tryCatch(beginTimer));

async function beginTimer() {
  await Excel.run(async (context) => {
    // Add a selection changed event handler for the workbook.
    setInterval(updateInfo, 2000);
    console.log("updating cells on a timer");
    await context.sync();
  });
}

// async function registerEventHandlers() {
//   await Excel.run(async (context) => {
//     // Add a selection changed event handler for the workbook.
//     context.workbook.worksheets.onSelectionChanged.add(updateInfo);
//     console.log("Change the selected cell");
//     await context.sync();
//   });
// }

let clicked = false;
const worksheets = ["Sheet1", "Sheet2", "Sheet3"];
i = 0;

async function updateInfo() {
  await Excel.run(async (context) => {
    //get selected cell value

    // console.log(worksheets[i], "worksheets[i]");
    // //toggle variable to alternate cell selection:
    // clicked = !clicked;
    // console.log(clicked);

    let sheet = context.workbook.worksheets.getItem(worksheets[i]);
    let range = sheet.getRange("B1:B6");
    range.load("text");
    await context.sync();
    let className = range.text[0][0];

    // let secondPlace = sheet.getRange("B3");
    // clicked ? (range = sheet.getRange("B2")) : (range = sheet.getRange("B3"));
    // console.log(cellText, "cellText");
    // let cellText = activeCell.values.toString();
    // console.log("The active cell is " + cellText);
    let firstPlace = range.text[1][0];
    console.log(firstPlace, "firstPlace");
    // let fullRangeText = range.text;
    // console.log(fullRangeText[0], "fullRangeText[0]");

    // firstPlace.load("text");
    // await context.sync();

    //iterate:
    i === worksheets.length - 1 ? (i = 0) : i++;
    console.log(i, "<<<<<< i = ");

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
        inputName: document.getElementById("Source").value,
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
        inputName: document.getElementById("Source2").value,
        inputSettings: {
          text: firstPlace,
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
