$("#register-event-handlers").click(() => tryCatch(registerEventHandlers));
$("#startTimer").click(() => tryCatch(startTimer));

async function registerEventHandlers() {
  await Excel.run(async (context) => {
    // Add a selection changed event handler for the workbook.
    context.workbook.worksheets.onSelectionChanged.add(
      onWorksheetSelectionChange
    );
    console.log("script running");
    await context.sync();
  });
}

function startTimer() {
  let intervalId;
  function oneSecondLog() {
    // check if an interval has already been set up
    intervalId = setInterval(writeLog, 1000);
  }
  function writeLog() {
    console.log("1 second has passed");
  }
  oneSecondLog();
}

async function changeCellTimer() {
  console.log("change cell timer started");
  let intervalId;
  function oneSecondCellChange() {
    // check if an interval has already been set up
    console.log("change cell one second timer");
    intervalId = setInterval(changeCells, 1000);
  }
  oneSecondCellChange();
}

async function changeCells() {
  await Excel.run(async (context) => {
    console.log("changing cells");
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let range = sheet.getRange("B3");
    range.load("text");
    await context.sync();
    let cellText3 = range.text[0][0];
    console.log(cellText3, "cellText3");
  });
}

async function onWorksheetSelectionChange(
  args: Excel.WorksheetSelectionChangedEventArgs
) {
  await Excel.run(async (context) => {
    console.log("clicked - connected to OBS");
    await changeCellTimer();

    //get selected cell value
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let range = sheet.getRange("B2");
    range.load("text");
    await context.sync();
    let cellText2 = range.text[0][0];
    console.log(cellText2, "cellText2");
    // changeCells();
    //

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
          text: cellText2,
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
          text: cellText3,
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
