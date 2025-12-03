// Office Add-in command handlers
Office.onReady((info) => {
    if (info.host === Office. HostType.Outlook) {
        console.log("Email to PDF Exporter - Commands loaded");
    }
});

function showTaskpane(event) {
    return Office.addin.showAsTaskpane()
        .then(() => {
            event.completed();
        })
        .catch(error => {
            console. error("Error:", error);
            event.completed();
        });
}

// Register the command handler
if (typeof FunctionRegistry !== "undefined") {
    FunctionRegistry.associate("showTaskpane", showTaskpane);
}