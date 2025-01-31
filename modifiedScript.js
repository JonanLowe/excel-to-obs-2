$("#register-event-handlers").click(() => tryCatch(beginTimer));

async function beginTimer() {
  await Excel.run(async (context) => {
    // Add a selection changed event handler for the workbook.
    setInterval(updateInfo, 1000)
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

async function updateInfo(args: Excel.WorksheetSelectionChangedEventArgs) {
  await Excel.run(async (context) => {
    //get selected cell value
    clicked = !clicked;
    console.log(clicked)

    let myWorkbook = context.workbook;
    let activeCell = myWorkbook.getActiveCell();
    activeCell.load("values");
    await context.sync();
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let range = sheet.getRange("A1");
    clicked ? range = sheet.getRange("B2") : range = sheet.getRange("B3") 
    range.load("text");
    await context.sync();
    let cellText = range.text[0][0];
    console.log(cellText, "cellText");
    // let cellText = activeCell.values.toString();
    console.log("The active cell is " + cellText);
    sheet = context.workbook.worksheets.getItem("Sheet1");
    clicked ? range = sheet.getRange("B3") : range = sheet.getRange("B4") 
    range.load("text");
    await context.sync();
    let cellText2 = range.text[0][0];
    console.log(cellText2, "cellText2");

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
          rpcVersion: 1
        }
      );
      console.log(`Connected to server ${obsWebSocketVersion} (using RPC ${negotiatedRpcVersion})`);
    } catch (error) {
      console.error("Failed to connect", error.code, error.message);
    }
    obs.on("error", (err) => {
      console.error("Socket error:", err);
    });

    //set OBS Scene
    await obs.call("SetCurrentProgramScene", { sceneName: document.getElementById("Scene").value });

    //set OBS source text
    await obs.call(
      "SetInputSettings",
      {
        inputName: document.getElementById("Source").value,
        inputSettings: {
          text: cellText
        }
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
          text: cellText2
        }
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
